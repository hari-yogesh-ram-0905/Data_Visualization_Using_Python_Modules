# Project Title : (Read Excel Files) - [Data Cleaning , Data Integration & Visualization using Python modules]

---
## Project Overview :

This project is a command-line-based Python application that automates the process of:

- Loading multiple Excel files (`.xls` or `.xlsx`) from a given folder
- Merging and cleaning the data (removing duplicates)
- Saving the cleaned result to a new Excel file
- Performing interactive visualizations on selected numeric columns

It uses powerful Python libraries like `pandas`, `seaborn`, and `matplotlib` for data processing and visualization.

---

## Steps Involved in the Project :

### 1. Importing Required Libraries

Import essential Python libraries like `pandas`, `seaborn`, `matplotlib`, `os`, and `glob` to handle data processing, visualization, and file operations.

### 2. Get Excel Files from Folder

The user inputs a folder path. The code collects all `.xls` and `.xlsx` files from the given path using the `glob` module.

### 3. Load Data from Excel Files

Each Excel file is read into a DataFrame using `pandas.read_excel()`. All valid DataFrames are appended to a list.

### 4. Merge and Clean DataFrames

All loaded DataFrames are concatenated into one using `pd.concat()`. Duplicate rows are removed using `drop_duplicates()`.

### 5. Save the Final Cleaned File

The final merged and cleaned DataFrame is saved as a new Excel file using `to_excel()`.

### 6. Load Final File for Visualization

The user is prompted to enter the path of the final cleaned Excel file for visualization. The file is loaded into a DataFrame.

### 7. Visualize the Data

User chooses two numeric columns and a plot type. Supported plots:

- 1) Histogram with KDE
- 2) Scatter Plot
- 3) Line Plot
- 4) Correlation Heatmap

### 8. Repeat or Exit

The user can continue to visualize more plots or exit the visualization interface.

---

## Features

### Excel Data Handling

- Automatically detects and reads all Excel files from a specified folder
- Combines them into a single DataFrame
- Removes duplicate rows to ensure clean output
- Saves the final DataFrame to a new Excel file

###  Interactive Visualization

- Load the cleaned Excel file and explore its structure
- Select any two numeric columns to plot

- Supported Plots:

  - Histogram with KDE (distribution)
  - Scatter Plot (relationship)
  - Line Plot (trend over a sequence)
  - Correlation Heatmap

### Modular Code Structure

- Functions are neatly separated by task:

  - File extraction
  - DataFrame merging
  - Excel I/O
  - Plotting routines
  - Interactive CLI-based plotting interface

---

## Tech Stack (OR) Tools & Technologies Used 

| Tool/Library | Purpose                 |
| ------------ | ----------------------- |
| Python       | Core language           |
| pandas       | Excel data handling     |
| seaborn      | Statistical plots       |
| matplotlib   | Basic visualizations    |
| os, glob     | File handling utilities |

---

## Getting Started

### Prerequisites

- Python 3.x installed
- Install required libraries:

```bash
pip install pandas matplotlib seaborn openpyxl
```

### Run the Project

```Vs Code Terminal :
python <your-main-filename>.py
```

Follow the terminal prompts to:

1. Input a folder path containing Excel files
2. Merge and clean the data
3. Save the merged file
4. Select columns for plotting

---

## Sample Use Cases

- Combine monthly sales reports
- Analyze survey results from multiple sheets
- Perform quick statistical visualizations on combined Excel data

---

## Project Structure

```text
├── main.py                # Main execution script
├── README.txt             # Project overview and instructions
└── output/
    └── merged_file.xlsx   # Final cleaned and saved Excel file
```

---

## Future Improvements

- GUI with Streamlit
- Support for categorical column plots
- Export plots as images or PDFs
- Add missing value handling

---

## Author

**HARI YOGESH RAM B**


---

## License

This project is open-source and available under the [MIT License](LICENSE).

---

## Note

This project is a great addition to your Data Science or Python Developer portfolio. It shows practical skills in data integration, cleaning, and visualization using real-world Excel data.

---
