# Excel-AI-project

## Overview
This project automates the workflow of reading and filtering Japanese CSV data, processing it in Excel via VBA macros, and post-processing results using OpenAI's API.

## Workflow
1. **CSV Preparation**: `main.py` reads and filters CSV files from the `public` directory, saving filtered results to `runMacro/target_macro_input.json`.
2. **Excel Automation**: `vba_simulation.py` loads the JSON, writes it to `runMacro.xlsm`, runs the macro, and saves results to `runMacro/target_macro_output.json`.
3. **AI Processing**: `ai_connect.py` sends macro output JSON to OpenAI for further processing (e.g., language conversion).

## Usage
1. Place your CSV files in the `public` directory.
2. Run `main.py` to filter and export data:
   ```powershell
   python main.py
   ```
3. Run `vba_simulation.py` to automate Excel and run the macro:
   ```powershell
   python vba_simulation.py
   ```
4. Run `ai_connect.py` to process the macro output with OpenAI:
   ```powershell
   python ai_connect.py
   ```

## Requirements
- Python 3.10+
- Install dependencies:
  ```powershell
  pip install -r requirements.txt
  ```
- Excel with macros enabled
- OpenAI API key in `.env` file

## File Descriptions
- `main.py`: CLI for CSV filtering and export
- `init_csv.py`: CSV reading and filtering utilities
- `vba_simulation.py`: Excel automation and macro execution
- `ai_connect.py`: OpenAI API integration for post-processing

## Troubleshooting
- Ensure macros are enabled and Excel security settings allow automation.
- Place your OpenAI API key in a `.env` file as `OPENAI_API_KEY=your_key_here`.
- If you encounter encoding errors, check your CSV file format.

## License
MIT