import pandas as pd

def get_excel_columns(filepath):
    """Read excel files and return columns."""
    try:
        df = pd.read_excel(filepath, nrows=0)
        return df.columns.tolist()
    except Exception as e:
        raise Exception(f"Error while reading : {filepath}: {str(e)}")

def merge_files(main_path, source_path, mapping_dict, output_path):
    """
    Merge both files.
    mapping_dict is a dict : { "Colonne_Main": "Colonne_Source" }
    """

    df_main = pd.read_excel(main_path)
    df_source = pd.read_excel(source_path)

    rename_dict = {source_col: main_col for main_col, source_col in mapping_dict.items() if source_col not in ["Ignor", "Ignorer", "Ignore"]}
    
    columns_to_keep = list(rename_dict.keys())
    df_source_filtered = df_source[columns_to_keep].copy()
    
    df_source_filtered.rename(columns=rename_dict, inplace=True)

    df_final = pd.concat([df_main, df_source_filtered], ignore_index=True)
    df_final.to_excel(output_path, index=False)