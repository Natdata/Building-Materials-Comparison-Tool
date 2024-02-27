Building Materials Comparison Tool

This Python script is designed to automate the process of comparing building materials listed in two separate Excel sheets and assigning corresponding prices where matches are found. The primary goal is to streamline the process of matching materials from a cost estimate sheet with a supplier's offer sheet, thus facilitating quicker and more accurate cost estimation for construction projects.

Features:

Excel Sheets Input: The tool takes two Excel sheets as input. The first sheet (Offer Sheet) contains two columns: one for the building materials and another for their prices. The second sheet (Cost Estimate Sheet) includes a list of building materials for which prices need to be assigned.

Material Matching: The script compares material names between the Offer Sheet and the Cost Estimate Sheet. It looks for exact matches or partial matches within the names and assigns the corresponding prices from the Offer Sheet to the materials in the Cost Estimate Sheet.

Flexible Matching Logic: Utilizes string manipulation and comparison techniques to identify matches even when the material names are not identical, catering for variations in naming conventions.

Output with Prices Assigned: Produces an updated version of the Cost Estimate Sheet with a new column for prices, where each material from the cost estimate that matches a material in the offer sheet has the relevant price assigned.

Implementation Details
Pandas Library: Leverages the Pandas library for efficient data manipulation and Excel file handling. This includes reading Excel files into DataFrame objects, iterating over rows, and applying conditional logic for matching materials.

String Comparison: Employs basic string comparison techniques for matching material names. This approach can be further enhanced with more sophisticated text similarity algorithms for improved matching accuracy.

Excel File Output: Utilizes Pandas' capabilities to export the modified DataFrame back into an Excel file, providing a user-friendly output that can be immediately used for further analysis or reporting.

Usage
To use this script, users should have Python installed on their system along with the Pandas library. Input Excel files must be prepared according to the specified format, with the Offer Sheet containing materials and prices and the Cost Estimate Sheet listing the materials to be priced.

Future Enhancements
Implementation of advanced text matching algorithms (e.g., fuzzy matching) to improve the accuracy of material comparison.
Integration with a user interface for easier file selection and parameter configuration.
Development of a feedback loop to learn from manual adjustments and improve automatic matching over time.
This tool aims to save time and reduce errors in the process of costing building materials, providing a solid foundation for more accurate and efficient project budgeting.
