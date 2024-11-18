import pandas as pd

def load_input_data(file_name, sheet_name):
    return pd.read_excel(file_name, sheet_name=sheet_name)

def load_scenario_data(file_name, sheet_name):
    return pd.read_excel(file_name, sheet_name=sheet_name)

