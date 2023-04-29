# Excel Reviews to PDF

This Python script converts the text in the "review" column of an Excel file into a series of PDF documents. Each PDF file will contain up to 1000 pages, and if the content exceeds 1000 pages, the script will automatically create additional PDF files (e.g., reviews_part1.pdf, reviews_part2.pdf, etc.).

## Requirements

- Python 3.6 or higher
- pandas
- openpyxl
- reportlab

## Installation Guide

1. Clone the repository or download the script:

```bash
git clone https://github.com/jellzone/Excel-Reviews-to-PDF.git
```

2. Change the directory
```cd excel-reviews-to-pdf```

3. Install the required Python libraries:

```pip install pandas openpyxl reportlab```

4. Edit the main.py file and replace example.xlsx with the path to your Excel file and 'Sheet1' with the name of the worksheet containing the "review" column.

5. Run the script:
```python main.py```

The generated PDF files will be saved in the same directory as the script.

Open Source Agreement
This project is open source under the MIT License. By using or contributing to this project, you agree to the terms of the MIT License. For more information, please refer to the LICENSE file included in the repository.

License
MIT License

Copyright (c) 2023 Jellzone

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
