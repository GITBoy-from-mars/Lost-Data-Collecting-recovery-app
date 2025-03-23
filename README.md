##Step 1:##

**Python script (main.py)** automates the process of searching for .xml files on an Android device, transferring them to a local computer, and uploading them to Google Drive. It uses ADB (Android Debug Bridge) to communicate with the connected Android device and Google Drive API for file uploads. Below is a step-by-step breakdown of how the script works:

Check ADB Availability: The script verifies whether ADB is installed and accessible by running adb --version. If ADB is not found, the script terminates.

Check Device Connection: It runs adb devices to check if an Android device is connected and authorized. If a device is found, it retrieves the device's name using adb shell getprop ro.product.model.

Google Drive Authentication: The script authenticates with Google Drive using OAuth 2.0 credentials stored in wapcos.json. If a valid token does not exist, it prompts the user to log in and saves the token for future use.

Find .xml Files on the Device: Using the ADB find command, it searches for all .xml files in the path /sdcard/Android/data/your directory/.

Download .xml Files to Local Storage: For each file found, it:

Extracts the filename and ensures no duplicate names by appending an index if necessary.

Pulls the file from the device to the Downloads/path directory using adb pull.

Upload to Google Drive: The script checks if the file already exists in the specified Google Drive folder. If found, it updates the existing file; otherwise, it creates a new one.

Error Handling: The script handles various errors, such as missing ADB, no connected devices, authentication failures, and file transfer issues.

Execution Flow: If ADB is available and a device is connected, the script initiates the file search, transfers files, uploads them to Google Drive, and notifies the user when the process is complete.



#Step 2:

One of these Python script (xmltoexcel.py) automates the conversion of XML files to Excel spreadsheets, ensuring both individual and combined outputs. It scans a specified directory for .xml files, extracts their contents while preserving hierarchy, and saves the structured data in Excel format. The process includes error handling, timestamped file naming, and an aggregated Excel file containing all parsed data.

Key Features & Workflow:
Recursively Parses XML Data: Extracts text content and attributes from nested XML elements while maintaining a hierarchical structure.

Finds XML Files: Searches for all .xml files within a given folder.

Creates Individual Excel Files: Converts each XML file into an Excel sheet, ensuring unique filenames by appending a timestamp.

Generates a Combined Excel File: Merges all converted XML data into a single Excel file (combined result.xlsx).

Handles Errors Gracefully: Reports failed file conversions and logs any issues encountered.

Supports Multiple Folders: The script can process multiple directory paths by iterating through a predefined list of input folders.

Configurable Input & Output Paths: Allows easy customization of the input folder and output folder locations.

Ensures Unique Column Headers: Aggregates all possible XML field names across multiple files to maintain consistency in the final dataset.

Loop-Based Batch Processing: Can iterate through multiple project folders (path3 to path16 in the Wapcos project) for bulk conversions.


**Note** - The xmltoexcel.py, xmltoexcel1.py, xmltoexcel2.py the work of these files are same as mentioned above but the key diffrence is some of my data contain complex .xml data and child data so i divided this in three parts and do some updates also according to data






#Step 3:

The script named arrange.py automates the collection, renaming, and consolidation of Excel files from multiple directories into a single master sheet. It searches for "combined result.xlsx" files in structured folder paths, renames them based on their source path, and merges them into a multi-sheet Excel workbook.

Key Features & Workflow
Fetch & Rename Excel Files:

Searches for combined result.xlsx in path1 to path N directories inside the Wapcos project.

Renames files to "Combined result of Output of pathX.xlsx" for clarity.

Copies renamed files to a dedicated combined sheet folder.

Directory Processing & Automation:

Iterates through directories path1, path2, ..., stopping when a directory is missing.

Ensures new folders are created if necessary.

Logs missing files to prevent errors.

Master Sheet Generation:

Collects all renamed Excel files.

Merges them into a single Excel workbook with separate sheets for each file.

Uses openpyxl for efficient writing.

This script simplifies data management for large-scale XML-to-Excel conversions, providing organized, consolidated reports for analysis





# Overall Goal

This Python script automates the process of:

Step 1:
Retrieving XML files from an Android device connected via ADB.
Saving those XML files to your local computer's Downloads folder.
Uploading those XML files to a specific folder in your Google Drive.

Step 2:
Converting the XML files to Excel files.

Step3:
Processing the Excel files to clean and combine the data.

Step4: 
Matching the Column IDs to main data Header
