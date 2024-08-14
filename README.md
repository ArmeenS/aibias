# Excel-to-DALL-E Image Generator:

This Python script is designed to automate the process of generating images using OpenAI's DALL-E based on prompts listed in an Excel sheet. The script reads prompts from specified cells, generates corresponding images using the DALL-E API, and inserts these images into designated cells in the same or different Excel sheets. This tool is particularly useful for research projects or any tasks where large-scale image generation is needed based on structured inputs.


Key Features:

1. Automated Image Generation: Reads prompts from an Excel sheet and generates images using OpenAI's DALL-E API.

2. Excel Integration: Automatically inserts generated images into specified cells in the Excel workbook.

3. Error Handling: Includes error handling for API errors with a retry mechanism.

4. Customizable: Easily configurable to handle different sheets, cells, and prompts within the Excel workbook.


Requirements:

1. Python 3.11

2. openpyxl for handling Excel files.

3. requests for downloading images.

4. openai package for interacting with OpenAI's API.

5. An OpenAI API key.


Usage:

1. Prepare an Excel file (input.xlsx) with prompts in the designated cells.

2. Run the script to generate and insert images into the workbook.

3. The output is saved as output.xlsx.
