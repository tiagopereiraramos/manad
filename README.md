# MANAD Processor

A Python project for processing MANAD (Manual Normativo de Arquivos Digitais) files. This tool extracts and consolidates K150 (rubric descriptions) and K300 (rubric entries) records, generating detailed reports in Excel format.

## Features

- Parses K150 and K300 records from MANAD files.
- Consolidates rubric data by period and rubric code.
- Generates formatted Excel reports with rubric details and calculated values.
- Supports multiple input files for batch processing.

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/your-repo/MANADProcessor.git
   ```
2. Install the required Python packages:
   ```bash
   pip install pandas openpyxl
   ```

## Usage

1. Place your MANAD `.TXT` files in the input folder.
2. Update the paths for input (`pasta_entrada`) and output (`pasta_saida`) folders in the script.
3. Run the script:
   ```bash
   python manad_processor.py
   ```
4. The Excel reports will be generated in the specified output folder.

## Project Structure

### Classes and Dataclasses

#### `RegistroK300`

A dataclass representing rubric entries.

**Attributes:**
- `cod_rubrica` (str): Rubric code.
- `valor_rubrica` (float): Rubric value.
- `cod_reg_trab` (str): Employee code.
- `dt_comp` (str): Competency date.

#### `RegistroK150`

A dataclass representing rubric descriptions.

**Attributes:**
- `cod_rubrica` (str): Rubric code.
- `desc_rubrica` (str): Rubric description.

#### `MANADProcessor`

A class for processing MANAD files.

**Attributes:**
- `file_path` (str): Path to the MANAD file to be processed.
- `k300_data` (List[RegistroK300]): List containing all K300 records.
- `k150_data` (Dict[str, str]): Dictionary mapping rubric codes to their descriptions.

**Methods:**
- `load_data()`: Loads and parses MANAD file records.
- `parse_line(line: str)`: Processes each line based on its type (`K150` or `K300`).
- `parse_k300(line: str) -> RegistroK300`: Processes K300 records (rubric entries).
- `parse_k150(line: str) -> RegistroK150`: Processes K150 records (rubric descriptions).
- `process_data() -> Tuple[pd.DataFrame, pd.DataFrame]`: Consolidates and prepares data for reporting.
- `gerar_relatorio_formatado(df_bruto, agrupado, descricoes_rubricas, arquivo_saida)`: Creates a formatted Excel report.

## Example Workflow

```python
from manad_processor import MANADProcessor

# Initialize processor with the file path
processor = MANADProcessor("path/to/manad/file.txt")

# Load and parse data
processor.load_data()

# Process data
df_bruto, agrupado = processor.process_data()

# Generate the report
processor.gerar_relatorio_formatado(
    df_bruto,
    agrupado,
    processor.k150_data,
    "output_report.xlsx"
)
```

## Output Structure

The generated Excel report contains the following columns:
- `mes_ano`: Month and year of the rubric entry.
- `Rubrica`: Rubric code.
- `Nome da Rubrica`: Rubric description.
- `Nº Empregados/Contribuintes`: Number of unique employees or contributors.
- `Valor Informado`: A placeholder column for manual input.
- `Valor Calculado`: Total value for the rubric.

## Dependencies

- **Python 3.8+**
- pandas
- openpyxl

Install the dependencies with:
```bash
pip install pandas openpyxl
```

## License

This project is licensed under the MIT License.

