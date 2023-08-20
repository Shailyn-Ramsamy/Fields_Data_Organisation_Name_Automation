# Fields_Data_Organisation_Name_Automation

### Matching and Highlighting Tool
This script is designed to perform matching between strings in two Excel files and highlight certain cells based on matching criteria. It utilizes the pandas, thefuzz, openpyxl, and concurrent.futures libraries.

### Installation
Before running the script, you need to ensure that you have the required libraries installed. You can install them using the following commands:

```bash
pip install pandas thefuzz openpyxl
```

### Usage
Place your input Excel file named "Input.xlsx" in the same directory as the script.
Place the reference Excel file named "FD_Orgs.xlsx" in the same directory as the script.
Run the script using your preferred Python interpreter.

### Purpose
1. The script performs the following tasks:

2. Read the input Excel files: It reads the "Input.xlsx" and "FD_Orgs.xlsx" files using the pandas library.

3. Matching Logic: The script matches values from the "Input" column of the "Input.xlsx" file against a set of unique organization names from the "FD_Orgs.xlsx" file. The matching logic is based on fuzzy string comparison methods provided by the thefuzz library.
  * If a match is found with high confidence (similarity score greater than or equal to 93), the script assigns the matched organization name to the "Matches" column of the DataFrame.
  * For cases where the similarity score is in a certain range (77 to 92) and additional conditions are met, the script considers the match as uncertain and stores these values in the uncertain and uncertain_2 sets.
  * A special case handling logic is also applied, where if a certain similarity score threshold is met, the script assigns a predefined organization name.

7. Highlighting Cells: The script uses the openpyxl library to style and highlight cells in the output DataFrame. Cells containing certain values (e.g., uncertain matches) are highlighted with specific background colors using the highlight_cells function.

6. Parallel Processing: The script uses the concurrent.futures.ThreadPoolExecutor to perform the matching in parallel, which can significantly speed up the process for larger datasets.

### Customization
You can modify the similarity score thresholds for different matching confidence levels by adjusting the numerical values in the code.
You can customize the background colors for highlighting by modifying the highlight_cells function.
To add more special case handling, you can extend the handle_special_cases function and adjust the corresponding threshold.
Output
The script displays the final DataFrame with matched organizations and highlighted cells based on the defined logic.

### Note
+ The script assumes that the input files are in the same directory as the script. If they are located in a different directory, you should provide the correct file paths when reading the files.
+ Make sure to adjust the script if your input files have different column names or structures.
  
Important: Before using this script on critical data, ensure to thoroughly test it with sample data to verify its behavior and suitability for your use case.
