# 📊 Sales Report Generator

A web application that automatically generates comprehensive sales reports with visualizations from Excel data.

## Features

- 📤 **Upload Excel Files**: Support for .xls and .xlsx formats
- 📊 **Automatic Calculations**: MTD and YTD Gross Sales and GP%
- 📈 **Year-over-Year Comparison**: Compare current year performance against previous year
- 🎨 **Professional Visualizations**: 
  - 4-panel dashboard chart
  - Year-over-year comparison chart
- 📥 **Excel Report Export**: Download formatted Excel file with:
  - Styled summary table
  - Embedded charts
  - Multiple worksheets

## Installation

### Using UV (Recommended)

```bash
# Install dependencies
uv pip install -r requirements.txt
```

### Using pip

```bash
# Install dependencies
pip install -r requirements.txt
```

## Usage

### Running the Application

```bash
streamlit run app.py
```

The app will open in your default web browser at `http://localhost:8501`

### Using the Application

1. **Upload File**: Click "Upload Sales Excel File" and select your Excel file
2. **Configure Settings** (in sidebar):
   - Select target month (1-12)
   - Select target year (e.g., 2025)
   - Select comparison year (e.g., 2024)
3. **View Results**: See the calculated metrics and summary table
4. **Download Report**: Click "Download Excel Report" to get the formatted report

## Expected File Format

Your Excel file should have:
- Header row at row 9 (0-indexed)
- Columns: Year, Period, Sales Amount, Cost of Sales
- Items identified by alphanumeric codes (e.g., AMO0002, YUH0019)
- "Item Total:" rows to separate different items

## Calculations

- **MTD Gross Sales**: Sum of sales amount for the target month
- **MTD GP%**: (MTD Sales - MTD Cost) / MTD Sales × 100
- **YTD Gross Sales**: Sum of sales amount from month 1 to target month
- **YTD GP%**: (YTD Sales - YTD Cost) / YTD Sales × 100
- **% Achieved**: (Current Year / Previous Year) × 100

## Output

The generated Excel report includes:
- **Sheet 1**: Formatted summary table with professional styling
- **Sheet 2**: Dashboard chart (4 panels showing MTD/YTD metrics)
- **Sheet 3**: Year-over-year comparison chart

## Project Structure

```
report-generator/
├── app.py                 # Main Streamlit application
├── main.py               # Original processing script
├── test.ipynb           # Jupyter notebook for testing
├── requirements.txt      # Python dependencies
└── README.md            # This file
```

## Technologies Used

- **Streamlit**: Web application framework
- **Pandas**: Data manipulation and analysis
- **Matplotlib**: Data visualization
- **OpenPyXL**: Excel file creation and formatting
- **NumPy**: Numerical computations

## License

MIT License

## Support

For issues or questions, please create an issue in the repository.
