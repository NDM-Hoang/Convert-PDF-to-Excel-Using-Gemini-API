# Convert PDF file or image to excel using Gemini API
📌 **Overview**

This application allows you to automatically convert PDF documents/images to Excel files using Google's Gemini API. Users can interact through a simple graphical interface. **This software is programmed with AI**.

![Application Demo](https://github.com/user-attachments/assets/9f34f6e1-e265-443c-92ad-38ca5054981d)

🚀 **Key Features**

- ✅ Select input PDF or image file

- ✅ Specify Excel file save directory

- ✅ Interact with Gemini API through interface

- ✅ Automatically generate Python code from AI

- ✅ Preview and edit code

- ✅ Execute code directly in the application

- ✅ Automatically open Excel file after creation

⚙️ **Installation**

  System Requirements
  
    Python 3.7+
    
    Operating System: Windows/macOS/Linux
  
  Install Libraries
  
    pip install -r requirements.txt
  
  Contents of requirements.txt:
  
    tkinter
    openpyxl>=3.1.2
    requests>=2.31.0
    python-dotenv>=1.0.0
    Pillow>=10.0.0
  🔑 **API Configuration**
  
  1. Get API Key from Google AI Studio
  
  2. Enter API Key in the corresponding field in the application
  
  3. API Key will be automatically saved at:
  
    Windows: C:\Users\[Username]\.excel_converter\config.json
  
    macOS/Linux: ~/.excel_converter/config.json

🖥️ **How to Use**

1. Launch the application:

       python gemini_excel_converter.py
  
2. Interface operations:

  - Enter your API Key

  - Select PDF/image file to process

  - Choose directory to save Excel file

  - Enter processing requirements (example: "Maintain table formatting")

  - Click "Run Prompt" to generate code

  - View and check code

  - Click "Run Code" to create Excel file

  - The result file will be saved at:


        [Selected_directory]/[Original_filename].xlsx
