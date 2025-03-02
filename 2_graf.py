import pandas as pd  # Import pandas for data manipulation
import graphviz  # Import graphviz for graph creation
import random  # Import random for generating random values
from collections import defaultdict  # Import defaultdict for creating dictionaries with default values

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
    
    # Filter out numbers that do not appear in more than one column
    repeated_numbers = {num: cols for num, cols in repeated_numbers.items() if len(set(cols)) > 1}
    
    return repeated_numbers  # Return the dictionary of repeated numbers

def random_color():
    """
    Generate a random color in hexadecimal format.

    Returns:
    str: A random color in hexadecimal format.
    """
    return f"#{random.randint(0, 0xFFFFFF):06X}"  # Generate a random color

def random_polygon():
    """
    Generate random polygon attributes.

    Returns:
    tuple: A tuple containing sides, distortion, orientation, and skew.
    """
    sides = random.randint(5, 10)  # Random number of sides between 5 and 10
    distortion = random.uniform(-1, 1)  # Random distortion between -1 and 1
    orientation = random.randint(0, 360)  # Random orientation between 0 and 360 degrees
    skew = random.uniform(-1, 1)  # Random skew between -1 and 1
    return sides, distortion, orientation, skew  # Return the polygon attributes

def create_graph(repeated_numbers):
    """
    Create a graph using Graphviz with nodes representing columns and edges representing repeated numbers.

    Parameters:
    repeated_numbers (dict): A dictionary where keys are repeated numbers and values are lists of columns where they appear.

    Returns:
    graphviz.Digraph: The created graph.
    """
    dot = graphviz.Digraph(comment='Repeated Numbers Graph')  # Create a new directed graph
    dot.attr('node', shape='polygon', color='white', style='filled', fontname='Arial')  # Set default node attributes
    
    # Add nodes and edges to the graph
    for number, columns in repeated_numbers.items():
        for col in columns:
            sides, distortion, orientation, skew = random_polygon()  # Generate random polygon attributes
            dot.node(col, col, sides=str(sides), distortion=str(distortion), orientation=str(orientation), skew=str(skew), color=random_color())  # Add a node for the column
        for i in range(len(columns)):
            for j in range(i + 1, len(columns)):
                dot.edge(columns[i], columns[j], label=str(number), color=random_color())  # Add an edge between columns with the repeated number
    
    return dot  # Return the created graph

def main():
    """
    Main function to find repeated numbers and create a graph.
    """
    filename = "numbers.xlsx"  # Path to the Excel file
    repeated_numbers = find_repeated_numbers(filename)  # Find repeated numbers in the Excel file
    
    if repeated_numbers:
        print("Repeated numbers and their locations:")  # Print a message indicating repeated numbers were found
        for number, columns in repeated_numbers.items():
            print(f"Number {number} found in columns: {', '.join(columns)}")  # Print the repeated numbers and their columns
        
        dot = create_graph(repeated_numbers)  # Create a graph of the repeated numbers
        dot.save('repeated_numbers_graph.dot')  # Save the graph in DOT format
        dot.render('repeated_numbers_graph', format='svg', cleanup=True)  # Render the graph in SVG format
        print("Graph saved as 'repeated_numbers_graph.dot' and 'repeated_numbers_graph.svg'")  # Print a message indicating the graph was saved
    else:
        print("No repeated numbers found in multiple columns.")  # Print a message indicating no repeated numbers were found

if __name__ == "__main__":
    main()  # Call the main function if the script is executed directly
