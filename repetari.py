import pandas as pd  # Import pandas for data manipulation
from collections import defaultdict  # Import defaultdict for creating dictionaries with default values
from itertools import combinations  # Import combinations for generating all possible combinations

def find_repeated_numbers(filename):
    """
    Find numbers that repeat in multiple columns of the Excel file.

    Parameters:
    filename (str): The path to the Excel file.

    Returns:
    dict: A dictionary where keys are repeated numbers and values are lists of columns where they appear.
    """
    df = pd.read_excel(filename)  # Read the Excel file into a DataFrame
    repeated_numbers = defaultdict(list)  # Create a defaultdict to store repeated numbers and their columns
    
    # Iterate through each column and record the occurrences of each number
    for col in df.columns:
        col_data = df[col]  # Get the data for the current column
        for idx, value in col_data.items():  # Iterate through each value in the column
            repeated_numbers[value].append(col)  # Append the column name to the list of columns for the current value
    
    return repeated_numbers  # Return the dictionary of repeated numbers

def save_repeated_numbers_to_csv(repeated_numbers, filename):
    """
    Save the repeated numbers to a CSV file in a mathematical graph format.

    Parameters:
    repeated_numbers (dict): A dictionary where keys are repeated numbers and values are lists of columns where they appear.
    filename (str): The name of the CSV file.
    """
    data = []  # Initialize an empty list to store the data
    for number, columns in repeated_numbers.items():  # Iterate through the repeated numbers
        for r in range(2, len(columns) + 1):  # Generate all possible combinations of columns
            for combo in combinations(columns, r):
                data.append([number, ', '.join(combo)])  # Append the number and its column combination to the data list
    
    df = pd.DataFrame(data, columns=['Number', 'Columns'])  # Create a DataFrame from the data
    df.to_csv(filename, index=False)  # Save the DataFrame to a CSV file
    print(f"Repeated numbers saved to {filename}")  # Print a message indicating the file was saved

def main():
    """
    Main function to find repeated numbers and save them to a CSV file.
    """
    filename = "numbers.xlsx"  # Path to the Excel file
    repeated_numbers = find_repeated_numbers(filename)  # Find repeated numbers in the Excel file
    
    if repeated_numbers:
        print("Repeated numbers and their locations:")  # Print a message indicating repeated numbers were found
        for number, columns in repeated_numbers.items():
            print(f"Number {number} found in columns: {', '.join(columns)}")  # Print the repeated numbers and their columns
        
        save_repeated_numbers_to_csv(repeated_numbers, "repeated_numbers.csv")  # Save the repeated numbers to a CSV file
    else:
        print("No repeated numbers found in multiple columns.")  # Print a message indicating no repeated numbers were found

if __name__ == "__main__":
    main()  # Call the main function if the script is executed directly
