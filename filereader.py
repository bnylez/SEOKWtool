import os
import pandas as pd
import sys

def calculate_priority(ratio):
    if ratio > 10:
        return "High"
    elif 5 <= ratio <= 10:
        return "Medium"
    else:
        return "Low"

def main(folder_path):
    dataframes = {}
    summary_data = []

    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        
        if os.path.isfile(file_path) and file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
            dataframes[file_name] = df

    for file_name, df in dataframes.items():
        avg_keyword_difficulty = avg_volume = None
        row_count_minus_one = len(df) - 1

        if "Keyword Difficulty" in df.columns:
            avg_keyword_difficulty = df["Keyword Difficulty"].mean()
        
        if "Volume" in df.columns:
            avg_volume = df["Volume"].mean()

        ratio = (avg_volume / avg_keyword_difficulty) if avg_volume is not None and avg_keyword_difficulty is not None and avg_keyword_difficulty != 0 else 0
        priority = calculate_priority(ratio)

        # Switch the order of 'Words Per Cluster' and 'Priority'
        summary_data.append([file_name, avg_keyword_difficulty, avg_volume, ratio, priority, row_count_minus_one])

    # Update the DataFrame column names to reflect the new order
    summary_df = pd.DataFrame(summary_data, columns=['File Name', 'Average Keyword Difficulty', 'Average Volume', 'Ratio', 'Priority', 'Words Per Cluster'])
    return summary_df


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python script.py <folder_path> <month> <week>")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    month = sys.argv[2]  # Get the month from command line
    week = sys.argv[3]   # Get the week from command line
    result_df = main(folder_path)
    
    # Save the dataframe to an Excel file in the same folder
    output_file_name = os.path.join(folder_path, f"{month}_Week_{week}_KW_Research_Report.xlsx")
    result_df.to_excel(output_file_name, index=False)
    print(f"Report saved as '{output_file_name}'")

