import json
import os
import datetime
import time
import gspread
import requests
import re
from PIL import ImageColor
from oauth2client.service_account import ServiceAccountCredentials
import webcolors

#------------------------------------------------------------------------------------------------------------------------------
# Write a message to the log (or screen). When running in AWS, print will write to Cloudwatch.
#------------------------------------------------------------------------------------------------------------------------------
def log (msg):
    logTimeStamp = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")
    print(str(logTimeStamp) + ": " + msg)

#------------------------------------------------------------------------------------------------------------------------------
# Get the closest matching color name.
#------------------------------------------------------------------------------------------------------------------------------
def getClosestColorName(RGB, Step):
    colorName = getColorName(RGB)

    if colorName == "Unknown":
        R = RGB[0]
        G = RGB[1]
        B = RGB[2]

        log("Finding closest match for RGB color: (" + str(R) + ", " + str(B) + ", " + str(G) + ")")

        # Try ranges of each color.
        rMin = R
        rMax = min(R + Step, 255)

        gMin = G
        gMax = min(G + Step, 255)

        bMin = B
        bMax = min(B + Step, 255)

        # Loop through, adding 1 to each color until we get a match.
        for rNew in range(rMin, rMax + 1):
            for gNew in range(gMin, gMax + 1):
                for bNew in range(bMin, bMax + 1):
                    RGBNew = (rNew, gNew, bNew)
                    colorName = getColorName(RGBNew)
                    
                    if colorName != "Unknown":
                        return colorName

    else:
        return colorName

#------------------------------------------------------------------------------------------------------------------------------
# Get the color name.
#------------------------------------------------------------------------------------------------------------------------------
def getColorName(RGB):
    colorName = "Unknown"
    colorHex = '%02x%02x%02x' % RGB
    urlAPI = "https://colornames.org/search/json/?hex=" + colorHex.upper()
    response = requests.post(urlAPI)

    if response.status_code == 200:
        colorName = response.json()['name']

        if colorName == None:
            colorName = "Unknown"

    return colorName

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

    # Open the All Colors sheet so that we can populate the color names.
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/15TgNrC84NVp9XX5UTXDwCidty3YKWlezTdHDL1gA3_w')
    worksheet = sheet.worksheet("All Colors")

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

            log("Processing palette, '" + pName + "'")

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
                    colorName = "" #getClosestColorName(RGB, 5)

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
