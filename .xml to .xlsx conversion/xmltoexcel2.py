import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime

def parse_element(element, parent_path, data_dict):
    """Recursively parse XML elements and collect data in dictionary"""
    current_path = f"{parent_path}/{element.tag}" if parent_path else element.tag
    
    # Add text content
    if element.text and element.text.strip():
        data_dict[f"{current_path}/text"] = element.text.strip()
    
    # Add attributes
    for attr, value in element.attrib.items():
        data_dict[f"{current_path}/@{attr}"] = value
    
    # Process children
    for child in element:
        parse_element(child, current_path, data_dict)

def convert_xml_folder(input_folder, output_folder):
    """Convert all XML files and create combined results"""
    os.makedirs(output_folder, exist_ok=True)
    xml_files = glob.glob(os.path.join(input_folder, "*.xml"))
    
    if not xml_files:
        print(f" No XML files found in {INPUT_FOLDER},{input_folder}")
        return
    
    success_count = 0
    failure_count = 0
    all_data = []
    all_columns = set()

    # First pass: Collect all possible columns
    for xml_file in xml_files:
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            temp_dict = {}
            parse_element(root, "", temp_dict)
            all_columns.update(temp_dict.keys())
        except Exception as e:
            print(f" Column detection failed for {os.path.basename(xml_file)}: {str(e)}")
            continue

    # Second pass: Process files and collect data
    for xml_file in xml_files:
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            data_dict = {col: None for col in all_columns}
            parse_element(root, "", data_dict)
            
            # Create individual Excel file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.basename(xml_file).replace('.xml', '')
            output_file = os.path.join(output_folder, f"{base_name}_{timestamp}.xlsx")
            
            df = pd.DataFrame([data_dict])
            df.to_excel(output_file, index=False)
            all_data.append(df)
            
            print(f" Converted: {os.path.basename(xml_file)}")
            success_count += 1
            
        except Exception as e:
            print(f" Failed: {os.path.basename(xml_file)} - {str(e)}")
            failure_count += 1

    # Create combined results
    if all_data:
        try:
            combined_df = pd.concat(all_data, ignore_index=True)
            combined_file = os.path.join(output_folder, "combined result.xlsx")
            
            # Write combined file with headers
            combined_df.to_excel(combined_file, index=False)
            print(f"\n Combined results saved to: {combined_file}")
        except Exception as e:
            print(f"\n Failed to create combined file: {str(e)}")
    else:
        print("\n No data to combine")

    print(f"\n Conversion Summary:")
    print(f"Successfully converted: {success_count} files")
    print(f"Failed conversions: {failure_count} files")



if __name__ == "__main__":
    
    INPUT_FOLDER = (f"S:\\Downloads\\your xml files directory")
    OUTPUT_FOLDER = (f"{INPUT_FOLDER}\\Outputs")

    print(" Starting XML to Excel conversion with combined results...")
    convert_xml_folder(INPUT_FOLDER, OUTPUT_FOLDER)

