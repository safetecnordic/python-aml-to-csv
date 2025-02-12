import pythonnet
pythonnet.load("coreclr")  # Ensures PythonNet loads .NET Core for .NET 8.0

import clr
import os

# Sets the dll directory, the aml directory and the output folder
dll_folder = r"C:\Users\WDAGUtilityAccount\Documents\python_aml_to_csv\dll_files"
aml_file_path = r"C:\Users\WDAGUtilityAccount\Documents\python_aml_to_csv\aml_file\SF 71.01.01-001 Safetec Update_2.aml"
output_folder = r"C:\Users\WDAGUtilityAccount\Documents\python_aml_to_csv\output_folder"


# Add references to the dlls
clr.AddReference(os.path.join(dll_folder, "Newtonsoft.Json.dll"))   
clr.AddReference(os.path.join(dll_folder, "EPPlus.dll")) 
clr.AddReference(os.path.join(dll_folder, "Aml.Engine.dll"))  
clr.AddReference(os.path.join(dll_folder, "AMLtoCSV.dll")) 

# Import C# classes 
from AMLtoCSV import AMLProcessor, AppAProcessor, AppB1Processor, AppB2Processor, CSVExporter

print("C# classes imported.")

# Initialize classes
aml_processor = AMLProcessor()
csv_exporter = CSVExporter()

# call method LoadAMLfile from aml_processor class (Reads aml file and returns caexdocument object)
caex_document = aml_processor.LoadAMLFile(aml_file_path)

# Process and export AppA data
app_a_processor = AppAProcessor()
app_a_data = app_a_processor.ExtractData(caex_document)
csv_exporter.ExportToCSV(app_a_data, os.path.join(output_folder, "AppA_Export.csv"))

# Process and export AppB1 data
app_b1_processor = AppB1Processor()
app_b1_data = app_b1_processor.ExtractData(caex_document)
csv_exporter.ExportToCSV(app_b1_data, os.path.join(output_folder, "AppB1_Export.csv"))

#  Process and export AppB2 data
app_b2_processor = AppB2Processor()
app_b2_data = app_b2_processor.ExtractData(caex_document)
csv_exporter.ExportToCSV(app_b2_data, os.path.join(output_folder, "AppB2_Export.csv"))

print("AML to CSV export complete.")
