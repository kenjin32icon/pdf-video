# Document Converter

This Python program provides a versatile document conversion utility that can transform various document formats into audio, video, or text. It features a user-friendly graphical interface built with tkinter.

## Features

- Convert multiple document formats:
  - PDF files (*.pdf)
  - Word documents (*.docx)
  - PowerPoint presentations (*.pptx)
  - Text files (*.txt)
  - Image files (*.png, *.jpg, *.jpeg)
- Extract text from documents
- Convert text to audio
- Modern and intuitive GUI interface
- Progress tracking for conversions
- Support for threading to prevent GUI freezing

## Prerequisites

Before running this program, make sure you have Python installed and the following dependencies:

```bash
pip install PyPDF2
pip install python-docx
pip install python-pptx
pip install Pillow
pip install pytesseract
pip install gTTS
pip install opencv-python
pip install numpy
```

Additionally, you'll need to install Tesseract OCR on your system for image text extraction:
- Windows: Download and install from [Tesseract GitHub](https://github.com/UB-Mannheim/tesseract/wiki)
- Make sure to add Tesseract to your system PATH

## Installation

1. Clone this repository:
```bash
git clone https://github.com/kenjin32icon/pdf-video.git
cd pdf-video
```

2. Install the required Python packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the program:
```bash
python pdf.py
```

2. Using the application:
   - Click "Select File" to choose your input document
   - The program will automatically extract text from the document
   - Review the extracted text in the display window
   - Click "Convert to Audio" to generate an audio version
   - Progress bar will show conversion status
   - Status updates appear at the bottom of the window

## Supported File Formats

- PDF (*.pdf)
- Microsoft Word (*.docx)
- Microsoft PowerPoint (*.pptx)
- Text files (*.txt)
- Images (*.png, *.jpg, *.jpeg)

## Error Handling

The program includes comprehensive error handling for:
- File access issues
- Unsupported file formats
- Conversion failures
- OCR processing errors

## Contributing

Feel free to fork this repository and submit pull requests for any improvements.

## License

MIT License

Copyright (c) 2024 kenjin32icon

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Author

- kenjin32icon
- Contact: kariukilewis04@gmail.com
