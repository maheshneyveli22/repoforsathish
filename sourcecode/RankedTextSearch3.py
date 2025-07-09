import pandas as pd

def analyze_keywords(data):
    """
    Analyzes a list of keywords to classify them as primary/secondary
    and generate master entries.

    Args:
        data (list of dict): A list of dictionaries, where each dictionary
                             represents a keyword with 'Keyword string (KH)' and 'Ranking'.

    Returns:
        pandas.DataFrame: A DataFrame containing the original, classified,
                          and master keyword entries.
    """

    # Convert the input list of dictionaries to a Pandas DataFrame
    df = pd.DataFrame(data)
    df.rename(columns={'Keyword string (KH)': 'Keyword', 'Ranking': 'Ranking'}, inplace=True)

    # Ensure the 'Keyword' column is treated as strings to prevent iteration errors
    df['Keyword'] = df['Keyword'].astype(str)

    # --- Step 1: Check if its primary or secondary ---
    # Initialize 'Class' column
    df['Class'] = ''

    # Get all keywords for comparison
    all_keywords = df['Keyword'].tolist()

    # Determine if each keyword is primary or secondary
    for index, row in df.iterrows():
        current_keyword = row['Keyword']
        is_primary = False
        for other_keyword in all_keywords:
            # A keyword is primary if it contains another *different* keyword as a substring
            if current_keyword != other_keyword and other_keyword in current_keyword:
                is_primary = True
                break
        df.at[index, 'Class'] = 'primary' if is_primary else 'Secondary'

    print("--- Step 1: Classified Keywords ---")
    print(df)
    print("\n" + "="*50 + "\n")

    # --- Step 2: Deep analysis (Create Master rows) ---
    master_rows = []
    # This map will store which master keyword each primary/secondary keyword belongs to for sorting
    keyword_master_group_map = {}

    primary_keywords_df = df[df['Class'] == 'primary']

    for index, primary_row in primary_keywords_df.iterrows():
        master_keyword_name = primary_row['Keyword']
        master_ranking_sum = primary_row['Ranking'] # Start with the primary keyword's own ranking

        # The primary keyword itself belongs to its own master group
        keyword_master_group_map[master_keyword_name] = master_keyword_name

        # Find all other keywords (primary or secondary) that are substrings of this master_keyword_name
        for _, other_row in df.iterrows():
            if other_row['Keyword'] != master_keyword_name and other_row['Keyword'] in master_keyword_name:
                master_ranking_sum += other_row['Ranking']
                # Assign this substring keyword to the current master group
                # If a keyword is a substring of multiple masters, the last one processed will "win"
                keyword_master_group_map[other_row['Keyword']] = master_keyword_name

        # Create the master row entry
        master_rows.append({
            'Keyword': master_keyword_name,
            'Ranking': master_ranking_sum,
            'Class': 'Master'
        })
        print(f"Master Keyword: '{master_keyword_name}' (Original Ranking: {primary_row['Ranking']})")
        print(f"  Master Ranking: {master_ranking_sum}\n")

    # Convert master rows to a DataFrame
    df_master = pd.DataFrame(master_rows)

    print("--- Step 2: Generated Master Rows ---")
    print(df_master)
    print("\n" + "="*50 + "\n")

    # --- Step 3: Make a table which has all the rows (primary, master and secondary) ---
    final_df = pd.concat([df, df_master], ignore_index=True)

    # Add a 'MasterGroup' column for hierarchical sorting
    # For each keyword, try to find its master group from the map.
    # If not found (e.g., a secondary keyword that isn't a substring of any primary),
    # it will group by itself (its own keyword).
    final_df['MasterGroup'] = final_df['Keyword'].apply(lambda x: keyword_master_group_map.get(x, x))

    # Define the desired order for the 'Class' column
    class_order = ['Master', 'primary', 'Secondary']
    # Convert 'Class' column to a categorical type with the specified order
    final_df['Class'] = pd.Categorical(final_df['Class'], categories=class_order, ordered=True)

    # Sort the final DataFrame by 'MasterGroup' (to group related keywords),
    # then by 'Class' (Master, primary, Secondary order),
    # then by 'Keyword' for alphabetical order within the same class and group.
    final_df = final_df.sort_values(by=['MasterGroup', 'Class', 'Keyword'], ascending=[True, True, True]).reset_index(drop=True)

    # Drop the temporary 'MasterGroup' column as it's only for sorting
    final_df = final_df.drop(columns=['MasterGroup'])

    print("--- Step 3: Final Combined Table ---")
    print(final_df)
    return final_df

# --- Example Usage ---
# Define input and output Excel file names
input_excel_file = "keyword_data.xlsx"
output_excel_file = "analyzed_keywords.xlsx"

try:
    # Read the data from an Excel file
    # Ensure your Excel file has columns named 'Keyword string (KH)' and 'Ranking'
    # You might need to install openpyxl: pip install openpyxl
    keyword_data_from_excel = pd.read_excel(input_excel_file).to_dict(orient='records')
    print(f"Successfully read data from '{input_excel_file}'\n")

    # Run the analysis with data from Excel
    final_table = analyze_keywords(keyword_data_from_excel)

    # Write the final table to a new Excel file
    # index=False prevents Pandas from writing the DataFrame index as a column
    final_table.to_excel(output_excel_file, index=False)
    print(f"\nSuccessfully wrote the final analyzed table to '{output_excel_file}'")

except FileNotFoundError:
    print(f"Error: The input file '{input_excel_file}' was not found.")
    print("Please make sure 'keyword_data.xlsx' exists in the same directory as the script.")
except Exception as e:
    print(f"An error occurred: {e}")

