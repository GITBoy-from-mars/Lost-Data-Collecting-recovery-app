import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime

def parse_element(element, parent_path, data_dict):
    """Parse XML elements with proper path construction"""
    current_path = f"{parent_path}/{element.tag}" if parent_path else element.tag
    
    # Handle attributes
    for attr, value in element.attrib.items():
        data_dict[f"{current_path}/@{attr}"] = value
    
    # Handle text content
    if element.text and element.text.strip():
        data_dict[f"{current_path}/text"] = element.text.strip()
    
    # Process children with indexing
    tag_counts = {}
    for child in element:
        tag = child.tag
        tag_counts[tag] = tag_counts.get(tag, 0) + 1
        child_index = tag_counts[tag]
        new_path = f"{current_path}/{tag}[{child_index}]" if tag_counts[tag] > 1 else f"{current_path}/{tag}"
        parse_element(child, new_path, data_dict)

def process_xml_file(xml_file):
    """Process XML file and separate header data from repeating elements"""
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    # First parse to get all data
    full_data = {}
    parse_element(root, "", full_data)
    
    # Separate header data (non-repeating elements)
    header_data = {k: v for k, v in full_data.items() if "Q3_cropping" not in k}
    
    # Process repeating Q3_cropping elements
    line_items = []
    for crop in root.findall('.//Q3_cropping'):
        item_data = {}
        parse_element(crop, "", item_data)
        line_items.append(item_data)
    
    # Create records with header only in first row
    records = []
    if line_items:
        # First record combines header + first line item
        records.append({**header_data, **line_items[0]})
        # Subsequent records contain only line items
        records.extend(line_items[1:])
    else:
        # If no line items, just use header data
        records.append(header_data)
    
    return records

def convert_xml_folder(input_folder, output_folder):
    """Convert XML files with proper header/lineitem separation"""
    os.makedirs(output_folder, exist_ok=True)
    xml_files = glob.glob(os.path.join(input_folder, "*.xml"))
    
    if not xml_files:
        print(f" No XML files found in {input_folder}")
        return
    
    all_columns = set()
    all_records = []
    success_count = 0
    failure_count = 0
    
    # First pass: Discover all columns
    for xml_file in xml_files:
        try:
            records = process_xml_file(xml_file)
            for record in records:
                all_columns.update(record.keys())
        except Exception as e:
            print(f" Column discovery failed for {os.path.basename(xml_file)}: {str(e)}")
    
    # Convert to sorted list for consistent ordering
    columns = sorted(all_columns)
    
    # Second pass: Process files
    for xml_file in xml_files:
        try:
            records = process_xml_file(xml_file)
            
            # Create DataFrame with consistent columns
            df = pd.DataFrame(records, columns=columns)
            
            # Save individual file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.basename(xml_file).replace('.xml', '')
            output_file = os.path.join(output_folder, f"{base_name}_{timestamp}.xlsx")
            df.to_excel(output_file, index=False)
            
            all_records.extend(records)
            success_count += 1
            print(f" Converted: {os.path.basename(xml_file)}")
            
        except Exception as e:
            failure_count += 1
            print(f" Failed: {os.path.basename(xml_file)} - {str(e)}")
    
    # Save combined results
    if all_records:
        combined_df = pd.DataFrame(all_records, columns=columns)
        combined_file = os.path.join(output_folder, "combined_results.xlsx")
        combined_df.to_excel(combined_file, index=False)
        print(f"\n Combined results saved to: {combined_file}")
    
    print(f"\n Conversion Summary:")
    print(f"Success: {success_count} files")
    print(f"Failed: {failure_count} files")


# Example usage
if __name__ == "__main__":
    INPUT_FOLDER = "S:\Desktop\path1"
    OUTPUT_FOLDER = f"{INPUT_FOLDER}/Outputs_of path1"
    
    print(" Starting XML conversion with header/lineitem separation...")
    convert_xml_folder(INPUT_FOLDER, OUTPUT_FOLDER)

    INPUT_FOLDER = "S:\Desktop\path2"
    OUTPUT_FOLDER = f"{INPUT_FOLDER}/Outputs_of path2"
    
    print(" Starting XML cfor path2")
    convert_xml_folder(INPUT_FOLDER, OUTPUT_FOLDER)

