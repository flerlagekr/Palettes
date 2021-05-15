#  Written by Ken Flerlage, April, 2021
#
#  Look up color names from a list of hex colors in Google Sheets
#  Uses the API from https://npm.io/package/color-name-list
#
#  This code is in the public domain

import boto3
import json
import os
import datetime
import time
import gspread
import requests
import re
import textwrap
from PIL import ImageColor
from oauth2client.service_account import ServiceAccountCredentials
import webcolors

overwriteAllColors = False

#------------------------------------------------------------------------------------------------------------------------------
# Write a message to the log (or screen). When running in AWS, print will write to Cloudwatch.
#------------------------------------------------------------------------------------------------------------------------------
def log (msg):
    logTimeStamp = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")
    print(str(logTimeStamp) + ": " + msg)

#------------------------------------------------------------------------------------------------------------------------------
# Get the color name using the API from https://npm.io/package/color-name-list
# API will automatcally get the closest match.
#------------------------------------------------------------------------------------------------------------------------------
def getColorName(colorHex):
    colorName = "Unknown"
    urlAPI = "https://api.color.pizza/v1/" + colorHex.upper()
    response = requests.post(urlAPI)

    if response.status_code == 200:
        colorName = response.json()['colors'][0]['name']

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
    colorIDs = worksheet.col_values(3)
    hexColors = worksheet.col_values(4)
    colorNames = worksheet.col_values(6)

    rowCount = 2

    # Loop through each color, get the name, and write back to the Google Sheet.
    for i in range(1, len(hexColors)):
        processColorName = False

        if i < len(colorNames):
            # Column has a value for this row.
            if overwriteAllColors == True or colorNames[i] == "":
                # Name needs to be populated.colorNames
                processColorName = True
            else:
                processColorName = False

        else:
            processColorName = True
        
        if processColorName == True:
            log("Processing color # " + str(colorIDs[i]) + " (" + hexColors[i] + ")")
            colorName = getColorName(hexColors[i])

            # Wrap the text so that it fits nicely.
            wrappedText = textwrap.fill(colorName, 16)

            # If no wrapping, add an artificial line feed.
            if wrappedText == colorName:
                wrappedText = wrappedText + "\n"

            worksheet.update_cell(rowCount, 6, colorName)
            worksheet.update_cell(rowCount, 7, wrappedText)

        else:
            log("Name already populated for Color # " + str(colorIDs[i]) + " (" + hexColors[i] + ")")


        rowCount += 1

    log("Wrote color names to Detail sheet.")


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
