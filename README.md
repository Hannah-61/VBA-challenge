# VBA Stock Market Data Analysis Project

## 📄 Overview
This project automates the analysis of quarterly stock market data using VBA scripting in Excel. It simplifies the process of summarizing data, applying conditional formatting, and identifying key metrics across multiple worksheets.

---

## 📊 Features
### Key Outputs:
- **Ticker Symbol**: Extracted for each stock.
- **Quarterly Change**: Difference between the opening and closing prices for the quarter.
- **Percentage Change**: Percentage difference between opening and closing prices.
- **Total Stock Volume**: Sum of all trades for each stock.

### Advanced Insights:
- **Greatest % Increase**: Stock with the highest percentage increase.
- **Greatest % Decrease**: Stock with the highest percentage decrease.
- **Greatest Total Volume**: Stock with the highest total volume.

### Conditional Formatting:
- Positive changes highlighted in **green**.
- Negative changes highlighted in **red**.

### Multi-Sheet Processing:
- Analyzes stock data across all worksheets in the workbook with a single run.

---

## 🛠 Requirements
### Data Retrieval:
The script processes and extracts:
- Ticker Symbol
- Volume of Stock
- Open Price
- Close Price

### New Columns:
- **Ticker Symbol**
- **Total Stock Volume**
- **Quarterly Change ($)**
- **Percentage Change (%)**

### Conditional Formatting:
- Applied to both **Quarterly Change** and **Percentage Change** columns for easy visualization.

### Metrics Calculated:
1. Greatest Percentage Increase
2. Greatest Percentage Decrease
3. Greatest Total Volume

### Multi-Sheet Capability:
- The script runs consistently across all worksheets.

---

## 🚀 Getting Started
### Files Included:
1. **VBA Script**: The main `.vbs` file for the analysis.
2. **Screenshots**: Demonstrating the output of the script.
3. **README File**: This documentation.

### Dataset:
- Use the provided `alphabetical_testing.xlsx` file for initial testing. It allows for quicker validation of the script.

### How to Use:
1. Open the workbook containing the stock data in Excel.
2. Open the VBA editor using `Alt + F11`.
3. Insert the provided VBA script into a module.
4. Run the script to analyze data for all worksheets.
5. Review the generated results, including:
   - Summarized statistics for each stock.
   - Conditional formatting to highlight positive/negative changes.
   - Identified best and worst-performing stocks.

---

## 📈 Example Output
Below is a sample of the expected output:

**Conditional Formatting Applied**  
- Positive values in **green**  
- Negative values in **red**

**Identified Metrics**  
- Greatest Percentage Increase  
- Greatest Percentage Decrease  
- Greatest Total Volume  

---

## 📂 Repository Structure
```plaintext
VBA-Challenge/
│
├── VBA_Script.vbs        
├── Screenshots/        
│   ├── Example1.png
│   └── Example2.png
├── alphabetical_testing.xlsx  
└── README.md            
