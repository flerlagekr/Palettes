#  Written by Ken Flerlage, April, 2021
#
#  Generate a Tableau color palette file based on crowdsourced palettes from the Tableau "Datafam"
#  Credit to Rodrigo Calloni for the original concept.
#  File is written to a public S3 bucket and can be downloaded here: https://flerlage-apps.s3.us-east-2.amazonaws.com/Preferences.tps
#
#  This code is in the public domain

import boto3
import json
import os
import datetime
import time
import gspread
import re
from PIL import ImageColor
from oauth2client.service_account import ServiceAccountCredentials
import webcolors

senderAddress = "Ken Flerlage <ken@flerlagetwins.com>"
ownerAddress = "flerlagekr@gmail.com"
paletteList = []

#------------------------------------------------------------------------------------------------------------------------------
# Write a message to the log (or screen). When running in AWS, print will write to Cloudwatch.
#------------------------------------------------------------------------------------------------------------------------------
def log (msg):
    logTimeStamp = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")
    print(str(logTimeStamp) + ": " + msg)

#------------------------------------------------------------------------------------------------------------------------------
# Tableau will ignore palettes with the same name (except first), so add spaces to make the name unique.
#------------------------------------------------------------------------------------------------------------------------------
def uniqueName(palName):
    tempName = palName

    while tempName in paletteList:
        tempName = tempName + " "

    return tempName

#------------------------------------------------------------------------------------------------------------------------------
# Get the closest matching color name.
#------------------------------------------------------------------------------------------------------------------------------
def closestColor(color):
    minColors = {}

    rd = 255
    gd = 255
    bd = 255

    for colorHex, colorName in webcolors.CSS3_HEX_TO_NAMES.items():
        try:
            rC, gC, bC = webcolors.hex_to_rgb(colorHex)
            rD = (rC - color[0]) ** 2
            gD = (gC - color[1]) ** 2
            bD = (bC - color[2]) ** 2

        except ValueError:
            pass

        minColors[(rD + gD + bD)] = colorName

    return minColors[min(minColors.keys())]

#------------------------------------------------------------------------------------------------------------------------------
# Get the color name.
#------------------------------------------------------------------------------------------------------------------------------
def getColorName(color):
    try:
        # Does this exact color have a name? 
        closestName = webcolors.rgb_to_name(color, "css3")
    
    except ValueError:
        # Exact color does not have a name. Find the closest color.
        closestName = closestColor(color)
        
    return closestName

#------------------------------------------------------------------------------------------------------------------------------
# Send message to Ken
#------------------------------------------------------------------------------------------------------------------------------
def phone_home (subject, msg):
    sender = senderAddress
    recipient = ownerAddress

    region = "us-east-2"

    # The email body for recipients with non-HTML email clients.
    bodyText = (msg)
                
    # The HTML body of the email.
    bodyHTML = """
    <html>
    <head></head>
    <body>
    <p style="font-family:Georgia;font-size:15px">""" + msg + """</p>
    </body>
    </html>
    """            

    # The character encoding for the email.
    charSet = "UTF-8"

    # Create a new SES resource and specify a region.
    client = boto3.client('ses',region_name=region)

    # Try to send the email.
    try:
        #Provide the contents of the email.
        response = client.send_email(
            Destination={
                'ToAddresses': [
                    recipient,
                ],
            },
            Message={
                'Body': {
                    'Html': {
                        'Charset': charSet,
                        'Data': bodyHTML,
                    },
                    'Text': {
                        'Charset': charSet,
                        'Data': bodyText,
                    },
                },
                'Subject': {
                    'Charset': charSet,
                    'Data': subject,
                },
            },
            Source=sender,
        )
    # Display an error if something goes wrong.	
    except ClientError as e:
        log (e.response['Error']['Message'])

#------------------------------------------------------------------------------------------------------------------------------
# Check a string to make sure it's a valid hex color code.
#------------------------------------------------------------------------------------------------------------------------------
def validHex(hexString):
 
    # Regex to check for valid. hexadecimal color code.
    regexPattern = "^([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$"
 
    # Compile the ReGex
    p = re.compile(regexPattern)
 
    # If the string is empty, return false
    if hexString == None:
        return False
 
    # Check if the string matches the pattern.
    if re.search(p, hexString):
        return True

    else:
        return False

#------------------------------------------------------------------------------------------------------------------------------
# Main lambda handler
#------------------------------------------------------------------------------------------------------------------------------
def lambda_handler(event, context):
    # Get the Google Sheets credentials from S3
    s3 = boto3.client('s3')
    bucket = "flerlage-lambda"
    prefBucket = "flerlage-apps"

    key = "creds.json"
    object = s3.get_object(Bucket=bucket, Key=key)
    content = object['Body']
    creds = json.loads(content.read())

    # Open Google Sheet
    scope =['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # Read your Google API key from a local json file.
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
    gc = gspread.authorize(credentials) 

    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/15TgNrC84NVp9XX5UTXDwCidty3YKWlezTdHDL1gA3_w')
    worksheet = sheet.worksheet("Form Responses 1")

    # Read the columns from the sheet.
    submitName = worksheet.col_values(3)
    colorType = worksheet.col_values(4)
    paletteName = worksheet.col_values(5)
    hexList = worksheet.col_values(6)

    # Open the output preferences file and write the beginning.
    prefFile = "Preferences.tps"

    out = "<?xml version='1.0'?>\n"
    out = out + "\n"
    out = out + "<workbook>\n"
    out = out + "    <preferences>\n"

    outString = ""

    # Initialize a matrix for generating the detailed data set.
    matrix = {}

    rowCount = 0

    # Loop through each palette, generate the XML, and write to the file.
    for i in range(1, len(paletteName)):
        # Generate the palette name, cleaning invalid characters along the way.
        if not(submitName[i] == "Ken Flerlage" and paletteName[i] == "All Colors"):
            # All Colors is used in Tableau only.
            submitted = submitName[i]
            submitted = submitted.replace('"', "'")
            submitted = submitted.replace('&', "and")

            palette = paletteName[i]
            palette = palette.replace('"', "'")
            palette = palette.replace('&', "and")

            pName = palette + " by " + submitted

            pName = uniqueName(pName)
            paletteList.append(pName)

            # Write the appropriate string based on the palette type.
            pType = colorType[i][0:3].upper()

            outString = '    	<color-palette name="' + pName + '" type="'

            if pType=="CAT":
                outString = outString + 'regular">'

            if pType=="SEQ":
                outString = outString + 'ordered-sequential">'

            if pType=="DIV":
                outString = outString + 'ordered-diverging">'

            outString = outString + "\n"
            out = out + outString

            colors = hexList[i].split(",")

            colorNum = 1

            for j in range(0, len(colors)):
                # Clean up the hex code and verify that it is a hex code.
                hexColor = colors[j].strip()
                hexColor = hexColor.replace("#", "")
                hexColor = hexColor.lower()

                if validHex(hexColor)==True:
                    # Write the hex code.
                    outString = "            <color>#" + hexColor + "</color>\n"
                    out = out + outString

                    # We need to round the color to the nearest 5 of R, G, and B.
                    RGB = ImageColor.getcolor("#" + hexColor, "RGB")
                    rOriginal = RGB[0]
                    gOriginal = RGB[1]
                    bOriginal = RGB[2]

                    # Get closest name
                    colorName = getColorName(RGB)

                    # Round to nearest 5.
                    R = 5 * round(rOriginal/5)
                    G = 5 * round(gOriginal/5)
                    B = 5 * round(bOriginal/5)

                    # Convert rounded values to hex.
                    RGB = (R, G, B)
                    hexColorRounded = '%02x%02x%02x' % RGB

                    # Add the color to the matrix.
                    matrix[rowCount, 0]  = submitName[i]
                    matrix[rowCount, 1]  = paletteName[i]
                    matrix[rowCount, 2]  = colorNum
                    matrix[rowCount, 3]  = hexColor
                    matrix[rowCount, 4]  = hexColorRounded
                    matrix[rowCount, 5]  = colorName
                    matrix[rowCount, 6]  = rOriginal
                    matrix[rowCount, 7]  = gOriginal
                    matrix[rowCount, 8]  = bOriginal

                    rowCount += 1
                    colorNum += 1

                else:
                    # Don't write the hex code, log an error, and send a notification.
                    if paletteName[i] != "All Colors":
                        errSubject = "Invalid Hex Code"
                        errMsg = "Invalid hex color code found in palette, '" + pName + "'. The hex code is '" + hexColor + "'."
                        log(errMsg)
                        phone_home(errSubject, errMsg)


            # Close out the palette.
            out = out + "        </color-palette>\n"

    # Write the end of the preferences file.
    out = out + "    </preferences>\n"
    out = out + "</workbook>\n"

    # Upload to S3
    response = s3.put_object(Bucket=prefBucket, Key=prefFile, Body=out) 
    log("Wrote preferences file to S3.")

    # Write data to Google Sheet
    rangeString = "A2:I" + str(rowCount+1)

    cell_list = worksheet.range(rangeString)

    row = 0
    column = 0

    for cell in cell_list: 
        cell.value = matrix[row,column]
        column += 1
        if (column > 8):
            column=0
            row += 1

    # Update in batch   
    worksheet = sheet.worksheet("Detail")
    worksheet.update_cells(cell_list)

    log("Wrote color details to Detail sheet.")


#------------------------------------------------------------------------------------------------------------------------------
# Labmda will always call the lambda handler function, so this will not get run unless you are running locally.
# This code will connect to AWS locally. This requires a credentials file in C:\Users\<Username>\.aws\
# For further details, see: https://boto3.amazonaws.com/v1/documentation/api/latest/guide/quickstart.html
#------------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    log("Code is running locally..............................................................")
    context = []
    event = {"state": "DISABLED"}
    boto3.setup_default_session(region_name="us-east-2", profile_name="default")
    lambda_handler(event, context)
    log("Code is complete.....................................................................")
