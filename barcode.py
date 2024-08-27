"""
*********************************************************
*                                                       *
*  Fatima Khan                                           *
*  Date: 2024-08-27                                      *
*                                                       *
*  Purpose:                                              *
*  To extract barcodes and ECC numbers from meat         *
*  packaging using OpenAI API                            *
*                                                       *
*  Version: 1.0                                          *
*                                                       *
*********************************************************
"""

import os
import base64
import requests
import openpyxl
import json

# OpenAI API Key
api_key = "API_KEY_HERE"

# Local folder path
folder_path = "C:\\Users\\Fatima.Khan\\Downloads\\iCloud Photos\\iCloud Photos"

# Excel file path (new file will be created)
excel_file = "C:\\Users\\Fatima.Khan\\Downloads\\barcode_extraction_results.xlsx"

# Initialize a new workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Extraction Results"

# Add headers to the Excel sheet
ws.append(["File Name", "Barcode Number", "Barcode Confidence", "Oval Text", "Oval Text Confidence"])

# Function to encode the image
def encode_image(file_path):
    with open(file_path, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
    return encoded_string

# Loop through each file in the local folder
for file_name in os.listdir(folder_path):
    if file_name.lower().endswith((".jpeg", ".jpg", ".png")):  # Process only image files
        print("Processing "+file_name)
        image_path = os.path.join(folder_path, file_name)
        
        # Encode the image
        base64_image = encode_image(image_path)
        
        # Prepare the headers for the API request
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        
        # Prepare the payload for the API request with structured output format
        payload = {
            "model": "gpt-4o-mini",
            "temperature" : "0.1",
            "messages": [
                {
                    "role": "system",
                    "content": "You are an AI highly specialized in accurately reading and interpreting barcodes and regulatory marks (like the EC or EEC mark) from images. You understand that barcodes can vary in length and format, with common lengths being 8, 12, or 13 digits. When extracting information: - Prioritize clarity and accuracy. - If the barcode is partially obscured, provide the most complete reading possible and indicate any uncertainties. - For the EC or EEC mark, ensure you extract all relevant text within the oval. If the text is unclear or missing, clearly state this. - Always validate the extracted barcode against common formats if possible, and specify if the format deviates from standard expectations. - Focus on extracting the numbers underneath the barcode rather than barcode widths. - Remove any spaces in your output. - Provide a confidence level (low, medium, high) for both the barcode and the oval text separately, based on the clarity and completeness of the data. - If any part of the image is unclear or likely to cause errors, mention this explicitly and adjust your confidence level accordingly."
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Please extract the barcode number and the text within the oval (EEC mark) from this image. Barcodes may vary in length, so extract the full sequence regardless of its digit count. Clearly indicate your confidence level (low, medium, high) for both the barcode and oval text, considering potential image quality issues or obstructions. If any part of the barcode or oval text is unreadable or uncertain, provide the best estimate and explain the source of uncertainty."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            "functions": [
                {
                    "name": "extract_barcode_and_text",
                    "description": "Extract the barcode number, text within the oval, and provide confidence levels.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "barcode_number": {
                                "type": "string",
                                "description": "The barcode number extracted from the image."
                            },
                            "barcode_confidence": {
                                "type": "string",
                                "enum": ["low", "medium", "high"],
                                "description": "Confidence level for the barcode extraction."
                            },
                            "oval_text": {
                                "type": "string",
                                "description": "The text extracted from the oval within the image."
                            },
                            "oval_text_confidence": {
                                "type": "string",
                                "enum": ["low", "medium", "high"],
                                "description": "Confidence level for the oval text extraction."
                            }
                        },
                        "required": ["barcode_number", "barcode_confidence", "oval_text", "oval_text_confidence"]
                    }
                }
            ],
            "function_call": {
                "name": "extract_barcode_and_text"
            },
            "max_tokens": 500
        }
        
        # Make the API request
        response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
        response_json = response.json()

        # Extract structured data from the response
        if 'choices' in response_json and len(response_json['choices']) > 0:
            choice = response_json['choices'][0]
            if 'message' in choice and 'function_call' in choice['message']:
                function_call = choice['message']['function_call']
                arguments = json.loads(function_call.get('arguments', '{}'))  # Parse the arguments JSON string
                
                # Extract fields from the structured output
                barcode_number = arguments.get('barcode_number', 'No Code Found')
                barcode_confidence = arguments.get('barcode_confidence', 'N/A')
                oval_text = arguments.get('oval_text', 'No Text Found')
                oval_text_confidence = arguments.get('oval_text_confidence', 'N/A')
                
                # Append the results to the Excel sheet
                ws.append([file_name, barcode_number, barcode_confidence, oval_text, oval_text_confidence])
        
        else:
            # Handle case where no valid response is received
            ws.append([file_name, "No Response", "N/A", "No Response", "N/A"])

# Save the new Excel file
wb.save(excel_file)

print(f"Results have been saved to {excel_file}")