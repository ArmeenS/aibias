# Excel-to-DALL-E Image Generator:

This Python script is designed to automate the process of generating images using OpenAI's DALL-E based on prompts listed in an Excel sheet. The script reads prompts from specified cells, generates corresponding images using the DALL-E API, and inserts these images into designated cells in the same or different Excel sheets. This tool is particularly useful for research projects or any tasks where large-scale image generation is needed based on structured inputs.


Key Features:

Automated Image Generation: Reads prompts from an Excel sheet and generates images using OpenAI's DALL-E API.

Excel Integration: Automatically inserts generated images into specified cells in the Excel workbook.

Error Handling: Includes error handling for API errors with a retry mechanism.

Customizable: Easily configurable to handle different sheets, cells, and prompts within the Excel workbook.


Requirements:

Python 3.x

openpyxl for handling Excel files.

requests for downloading images.

openai package for interacting with OpenAI's API.

An OpenAI API key.


Usage:

Prepare an Excel file (input.xlsx) with prompts in the designated cells.

Run the script to generate and insert images into the workbook.

The output is saved as output.xlsx.
