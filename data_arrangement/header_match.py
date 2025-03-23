import pandas as pd
import os

main_data_path = "S:\\Desktop\\master_sheet.xlsx" 
ids_directory = "S:\\Desktop\\Forms_IDs"    
output_path = "S:\\Desktop\\renamed_data.xlsx"  

#first one is worksheet name and second one is form IDs why which its changes
sheet_id_mapping = {
    'path': 'Form.xlsx',
    'path1': 'Form2.xlsx',
    'path2': 'Form3.xlsx',
}


all_sheets = pd.read_excel(main_data_path, sheet_name=None)

processed_sheets = {}


for sheet_name, df in all_sheets.items():
 
    id_file = sheet_id_mapping.get(sheet_name)
    
    if id_file:
 
        id_path = os.path.join(ids_directory, id_file)
        
        try:
            
            df_ids = pd.read_excel(id_path)
            id_to_name = df_ids.set_index('ID')['Name'].astype(str).to_dict()
            
            
            new_columns = []
            for col in df.columns:
            
                str_col = str(col)
                new_name = id_to_name.get(str_col, id_to_name.get(col, col))
                new_columns.append(new_name)
            
            df.columns = new_columns
            print(f"Processed {sheet_name} using {id_file}")
            
        except FileNotFoundError:
            print(f"ID file {id_file} not found for {sheet_name}, keeping original headers")
    
    processed_sheets[sheet_name] = df

# Save all processed sheets to new Excel file
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, df in processed_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\nAll sheets processed successfully! Saved to: {output_path}")
