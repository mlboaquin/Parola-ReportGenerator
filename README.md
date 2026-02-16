# Patent Report Generator

This tool automates the generation of patent reports (Invalidity, FTO, etc.) by processing Excel data files and populating Word document templates.

## Features

- **Automated Report Generation**: Generates comprehensive reports from Excel data.
- **Support for Multiple Report Types**: Currently supports Invalidity and FTO reports.
- **Update Mode**: Ability to update existing reports while preserving manual edits.
- **Web Scraping**: Fetches abstract and claim data from Google Patents when not available in the input.
- **GUI Interface**: User-friendly graphical interface built with PyQt6.

## Prerequisites

- Python 3.8+
- pip (Python package installer)

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```

2. The GUI will open.
3. Select the **Excel File** containing the patent data.
4. Select the **Word Template** (e.g., `Invalidity_Template.docx`).
5. Choose the **Report Type** (e.g., Invalidity).
6. Click **Generate Report**.

## File Structure

- `main.py`: Main application script containing the GUI and logic.
- `requirements.txt`: List of Python dependencies.
- `README.md`: This file.

## Dependencies

- **pandas**: Data manipulation and analysis.
- **requests**: HTTP library for web scraping.
- **beautifulsoup4**: HTML parsing for web scraping.
- **python-docx**: Creating and updating Microsoft Word files.
- **openpyxl**: Reading Excel files.
- **msoffcrypto-tool**: Handling password-protected Office files.
- **PyQt6**: GUI toolkit.

## License

[License Information]
