import pandas as pd

# Function to filter applicants based on qualification details (e.g., 60 in 'Result' column)
def filter_applicants_by_qualification(excel_file, min_qualification=60):
    # Read the Excel file into a pandas DataFrame
    try:
        df = pd.read_excel(excel_file, engine="openpyxl")  # Specify engine="openpyxl" for .xlsx files
        
        # Check if the 'Result' column exists
        if 'Result' not in df.columns:
            print("Error: 'Result' column not found in the Excel file.")
            return
        
        # Convert 'Result' column to numeric if it's not already (this will also handle any non-numeric values)
        df['Result'] = pd.to_numeric(df['Result'], errors='coerce')
        
        # Filter applicants with Result greater than or equal to the minimum qualification
        filtered_df = df[df['Result'] >= min_qualification]
        
        # Check if any applicants meet the qualification criteria
        if filtered_df.empty:
            print(f"No applicants found with Result >= {min_qualification}.")
        else:
            print(f"Applicants with Result >= {min_qualification}:")
            print(filtered_df)

            # Save the filtered results to a new Excel file (optional)
            output_file = f"filtered_applicants_{min_qualification}_percent.xlsx"
            filtered_df.to_excel(output_file, index=False)
            print(f"Filtered data saved to {output_file}")

    except FileNotFoundError:
        print(f"Error: The file '{excel_file}' was not found.")
    except pd.errors.EmptyDataError:
        print("Error: The file is empty or could not be read.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Usage example
if __name__ == "__main__":
    excel_file_path = "C:/Users/Cyber/python/applicants.xlsx"
    filter_applicants_by_qualification(excel_file_path, min_qualification=60)
