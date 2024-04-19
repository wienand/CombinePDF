# PDF and Image Combiner

This command-line tool allows you to combine PDF and image files into a single PDF file
based on an Excel list. You can specify the input files, output file names, and other
options using command-line arguments.

## Installation

1. Just download or clone the repository.

   ```
   git clone https://github.com/your-username/pdf-image-combiner.git
   cd pdf-image-combiner
   ````

2. Install the required dependencies:

   ```
   pip install -r requirements.txt
   ```

## Usage

```
python combine.py --source-excel path/to/excel_file.xlsx \
                  --column-input-files input_column_header \
                  --column-output-file output_column_header \
                  [--source-sheet sheet_name \]
                  [--root-path-input /path/to/input/files \]
                  [--root-path-output /path/to/output/files \]
```

## Arguments

```--source-excel```: Path to the Excel file containing IDs and file locations.

```--column-input-files```: Column containing the paths to the input files.

```--column-output-file```: Column containing the paths to the output files (combined PDFs).

```--source-sheet``` (optional): Specify the sheet name in the Excel file (if not provided, the active sheet will be
used).

```--root-path-input``` (optional): Root path for input file paths (useful if the paths in the Excel file are relative).

```--root-path-output``` (optional): Root path for output file paths.

## Examples

1. Combine PDFs and images listed in “Sheet1” of data.xlsx in columns Output and Input:

```
python combine.py --source-excel data.xlsx --column-input-files Input --column-output-file Output
```

2. Specify custom root paths for input and output files:

```
python combine.py --source-excel data.xlsx --column-input-files Input --column-output-file Output --root-path-input /path/to/input/files --root-path-output /path/to/output/files
```

## License
The Unlicense https://unlicense.org/
