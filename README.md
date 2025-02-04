# xlsxDecoder

## xlsxDataParser.py

Three classes to read an Excel file (xls, xlsm, xlsx) in a pandas dataframe considering only cell contents with valid format. For that, the .XML of the Excel file is read and evaluated. 

Focus: Read all data of a structured data set to receive all metadata information. The dataset consists of a header distinguidhing between different tests (column- or row-wise) and the metadata information (row- or column-wise) belonging to it.

🦾 Currently working on it!

### *class xlsxDecoder()*

Decoding the .XML file of the Excel file to read the cell formats.

### *class xlsxParser()*

Parse all cell entries in a pandas dataframe and recognising the orientation of the header in the Excel file and combined columns. Uses xlsxDecoder() class to remove not valid cell entries.

### *class ParallelPathFinder()*

Uses parallel computing to walk through all subfolders of a given folder finding files with a given name.
