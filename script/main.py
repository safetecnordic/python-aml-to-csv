
import pythonnet
pythonnet.load("coreclr")  # Ensure PythonNet loads .NET Core for .NET 8.0

import clr
import os

# Set DLL folder and AML paths
dll_folder = r"C:\Users\WDAGUtilityAccount\Documents\python-aml-to-csv\dll_files"
aml_file_path = r"C:\Users\WDAGUtilityAccount\Documents\python-aml-to-csv\aml_file\excel-to-aml_export-4.0.aml"
output_folder = r"C:\Users\WDAGUtilityAccount\Documents\python-aml-to-csv\output_folder"

# Add references to DLLs
clr.AddReference(os.path.join(dll_folder, "Newtonsoft.Json.dll"))
clr.AddReference(os.path.join(dll_folder, "EPPlus.dll"))
clr.AddReference(os.path.join(dll_folder, "Aml.Engine.dll"))
clr.AddReference(os.path.join(dll_folder, "AMLtoCSV.dll"))

# NOW import from C# classes namespace
import AMLtoCSV 

print("C# classes imported.")

# Initialize classes through the namespace
aml_processor = AMLtoCSV.AMLProcessor()
csv_exporter = AMLtoCSV.CSVExporter()

# Load AML file
caex_document = aml_processor.LoadAMLFile(aml_file_path)

# Process and export AppC data
app_c_processor = AMLtoCSV.AppCProcessor()
app_c_data = app_c_processor.ExtractData(caex_document)
csv_exporter.ExportToCSV(app_c_data, os.path.join(output_folder, "AppC_Export.csv"))

# Process and export AppB1 data
app_b1_processor = AMLtoCSV.AppB1Processor()
app_b1_data = app_b1_processor.ExtractData(caex_document)
csv_exporter.ExportToCSV(app_b1_data, os.path.join(output_folder, "AppB1_Export.csv"))

# Process and export AppB2 data
app_b2_processor = AMLtoCSV.AppB2Processor()
app_b2_data = app_b2_processor.ExtractData(caex_document)
csv_exporter.ExportToCSV(app_b2_data, os.path.join(output_folder, "AppB2_Export.csv"))


print("AML to CSV export complete.")
