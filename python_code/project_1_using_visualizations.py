# ---------------------------------------------------------------------------------------------------------------------------
# Import Required Modules
# ---------------------------------------------------------------------------------------------------------------------------


import pandas as pd

import seaborn as sns

import matplotlib.pyplot as plt

import os

from glob import glob


# ---------------------------------------------------------------------------------------------------------------------------
# Function: [ Get Excel Files from Folder ]
# ---------------------------------------------------------------------------------------------------------------------------


def get_excel_files(folder_path):


    xls_files = glob(os.path.join(folder_path, "*.xls"))

    xlsx_files = glob(os.path.join(folder_path, "*.xlsx"))


    return xls_files + xlsx_files


# ---------------------------------------------------------------------------------------------------------------------------
# Function: [ Load DataFrames from Excel Files ]
# ---------------------------------------------------------------------------------------------------------------------------


def load_excel_files(file_list):


    dataframes = []


    for file in file_list:

        try:

            df = pd.read_excel(file)

            dataframes.append(df)

            print(f"Loaded File: {file} | Rows: {len(df)}")


        except Exception as e:

            print(f"Failed to load file: {file} | Error: {e}")


    return dataframes


# ---------------------------------------------------------------------------------------------------------------------------
# Function: [ Merge and Clean DataFrames (drop duplicate rows) ]
# ---------------------------------------------------------------------------------------------------------------------------


def merge_dataframes(dataframes):


    if not dataframes:

        print("No valid Excel files could be read.")


    merged_df = pd.concat(dataframes, ignore_index=True)

    merged_df = merged_df.drop_duplicates()


    return merged_df


# ---------------------------------------------------------------------------------------------------------------------------
# Function: [ Save DataFrame to Excel ]
# ---------------------------------------------------------------------------------------------------------------------------


def save_to_excel(df, output_path):


    df.to_excel(output_path, index=False)

    print(f"\nFinal Merged file saved at: {output_path}")


# ---------------------------------------------------------------------------------------------------------------------------
# Function: [ Load Final Merged Excel File ]
# ---------------------------------------------------------------------------------------------------------------------------


def load_final_excel_file():


    file_path = input("\nEnter the full path of the final cleaned Excel file for visualization: ").strip()


    if not os.path.exists(file_path):

        print("File not found. Please check the path and try again.")

        return None


    try:

        df = pd.read_excel(file_path)

        print(f"Loaded final Excel file successfully with {df.shape[0]} rows and {df.shape[1]} columns.")

        return df
    

    except Exception as e:

        print(f"Error reading the file: {e}")
        

# ---------------------------------------------------------------------------------------------------------------------------
# Visualization Functions (total = 4 plots)
# ---------------------------------------------------------------------------------------------------------------------------


def plot_histograms(df, col_1, col_2):   #  To visualize the distribution of both the columns
   

    # Identify numeric columns :
    # ---------------------------

    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()


    if col_1 in numeric_cols and col_2 in numeric_cols:

        fig, axes = plt.subplots(1, 2, figsize=(14, 5))


        # Histogram for col_1 :
        # ----------------------

        sns.histplot(df[col_1], kde=True, ax=axes[0])

        axes[0].set_title(f"Distribution of {col_1}")

        axes[0].set_xlabel(col_1)
        axes[0].set_ylabel("Frequency")

        axes[0].grid(True)


        # Histogram for col_2 :
        # ---------------------

        sns.histplot(df[col_2], kde=True, ax=axes[1])

        axes[1].set_title(f"Distribution of {col_2}")

        axes[1].set_xlabel(col_2)
        axes[1].set_ylabel("Frequency")

        axes[1].grid(True)

        plt.tight_layout()

        plt.show()


    else:

        print(f"Error: One or both columns '{col_1}' and '{col_2}' are not numeric.")


# ---------------------------------------------------------------------------------------------------------------------------


def scatter_plot(df, col_1, col_2):    # To visualize the relationship or correlation between two numeric variables.
  

    # Check if columns are numeric:
    # -----------------------------

    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()

    
    if col_1 in numeric_cols and col_2 in numeric_cols:

        plt.figure(figsize=(8, 5))

        sns.scatterplot(x=df[col_1], y=df[col_2])

        plt.title(f"Scatterplot: {col_1} vs {col_2}")

        plt.xlabel(col_1)
        plt.ylabel(col_2)

        plt.grid(True)

        plt.tight_layout()

        plt.show()


    else:

        print(f"Error: One or both columns '{col_1}' and '{col_2}' are not numeric.")


# ---------------------------------------------------------------------------------------------------------------------------


def line_plot(df, col_1, col_2):  # To visualize the trend or relationship between the two numeric columns.

    # check if the columns are numeric :
    # ----------------------------------

    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()


    if col_1 not in numeric_cols or col_2 not in numeric_cols:

        print("Both columns must be numeric.")

        return
    

    plt.figure(figsize=(8, 5))

    sns.lineplot(data=df, x=col_1, y=col_2, marker='o')

    plt.title(f"Line Plot: {col_1} vs {col_2}")

    plt.tight_layout()

    plt.show()


# ---------------------------------------------------------------------------------------------------------------------------


def heatmap(df, col_1, col_2):   #  To visually represent the correlation — strong positive is red, strong negative is blue.

 
    # Check if both columns are numeric :
    # -----------------------------------

    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()


    if col_1 in numeric_cols and col_2 in numeric_cols:


        plt.figure(figsize=(10, 6))

        corr = df[[col_1, col_2]].corr()

        sns.heatmap(corr, annot=True, cmap='coolwarm')

        plt.title(f"Correlation Heatmap: {col_1} vs {col_2}")

        plt.tight_layout()

        plt.show()


    else:

        print(f"Error: One or both columns '{col_1}' and '{col_2}' are not numeric.")


# ---------------------------------------------------------------------------------------------------------------------------
# Function: [ Visualization Interface ]
# ---------------------------------------------------------------------------------------------------------------------------


def visualization_interface(df):


    while True:


        print("\nAvailable Columns in Data:")

        print(df.columns.tolist())


        col_1 = input("\nEnter the first column name (X-axis): ").strip()
        col_2 = input("Enter the second column name (Y-axis): ").strip()


        print("\nChoose Visualization Type:")
        

        print("1. Histogram with KDE")
        print("2. Scatter Plot")
        print("3. Line Plot")
        print("4. Heatmap")


        choice = input("Enter your choice (1–4): ").strip()

        
        if choice == "1":

            plot_histograms(df, col_1, col_2)

        elif choice == "2":

            scatter_plot(df, col_1, col_2)

        elif choice == "3":

            line_plot(df, col_1, col_2)

        elif choice == "4":

            heatmap(df, col_1, col_2)
        

        else:

            print("Invalid choice!. Please try again!.")

            continue


        cont = input("\nDo you want to visualize another plot? (yes/no): ").strip().lower()


        if cont not in ['yes', 'y']:

            print("Exit from the visualization session.")

            return


# ---------------------------------------------------------------------------------------------------------------------------
# Function: [ Visualize After Merge ]
# ---------------------------------------------------------------------------------------------------------------------------


def visualize_after_merge():


    df = load_final_excel_file()

    if df is not None:

        visualization_interface(df)


# ---------------------------------------------------------------------------------------------------------------------------
# Main Function
# ---------------------------------------------------------------------------------------------------------------------------


def main():


    folder_path = input("Enter the folder path containing Excel files: ").strip()


    if not os.path.isdir(folder_path):

        print("Invalid folder path. Please check and try again.")

        return


    excel_files = get_excel_files(folder_path)


    if not excel_files:

        print("No Excel files found in the folder. Please check the path and file extensions.")

        return
    

    all_files_data = load_excel_files(excel_files)


    if not all_files_data:

        print("No valid Excel files could be read.")

        return


    merged_files = merge_dataframes(all_files_data)


    if merged_files is None:

        print("Your Merged File returns None!....")

        return


    output_filename = input("Enter the name for the final merged Excel file (e.g: file_name.xlsx): ").strip()


    if not output_filename.endswith(".xlsx"):

        output_filename += ".xlsx"


    output_file = os.path.join(folder_path, output_filename)


    save_to_excel(merged_files, output_file)


    print(f"\nSuccessfully Merged {len(all_files_data)} files.")

    print(f"Total rows after Merging (excluding duplicates) = {len(merged_files)}")


# ---------------------------------------------------------------------------------------------------------------------------
# Execute Main + Visualization
# ---------------------------------------------------------------------------------------------------------------------------


if __name__ == "__main__":

    main()

    visualize_after_merge()
    

# ---------------------------------------------------------------------------------------------------------------------------