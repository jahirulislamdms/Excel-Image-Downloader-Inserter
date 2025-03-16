# Excel Image Downloader & Inserter

## Overview
**Excel Image Downloader & Inserter** is a Python script that automates the process of downloading images from URLs in an Excel file and embedding them into a new column in the same file. It provides an easy-to-use interface for selecting the Excel file, choosing the image URL column, and automatically downloading and inserting the images.

## Features
- **Graphical User Interface (GUI)** to select the Excel file and the column containing image URLs.
- **Automatic Image Downloading** from the URLs and saving them to a local folder.
- **Inserting Images** directly into a new column in the Excel file.
- **Save a New Excel File** with images embedded.

## Requirements
- Python 3.x
- Libraries:
  - `requests`
  - `openpyxl`
  - `tkinter` (pre-installed with Python)

## Installation
To install and run the script, follow these steps:

1. **Clone the repository:**
   ```bash
   git clone https://github.com/jahirulislamdms/Excel-Image-Downloader-Inserter.git
   cd excel-image-downloader

2. **Install required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

   Alternatively, you can install the required libraries individually:
   ```bash
   pip install requests openpyxl
   ```

## Usage
1. **Run the script** by executing the following command in your terminal:
   ```bash
   python script.py
   ```

2. **Select the Excel file** when the file dialog pops up.

3. **Choose the column** that contains the image URLs.

4. The script will **download the images** from the URLs and insert them into a new column in the Excel file.

5. A new file will be saved with "_with_images" added to the file name.

## Example

- **Input Excel File:**
  | Product | Image URL                      |
  |---------|--------------------------------|
  | Item A  | https://example.com/a.jpg      |
  | Item B  | https://example.com/b.jpg      |

- **Output Excel File:**
  | Product | Image URL                      | Downloaded Image |
  |---------|--------------------------------|------------------|
  | Item A  | https://example.com/a.jpg      | 🖼️ Image A       |
  | Item B  | https://example.com/b.jpg      | 🖼️ Image B       |

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
