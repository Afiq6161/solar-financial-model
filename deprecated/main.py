import pandas as pd
from input_handler import load_input_data, load_scenario_data
from scenario_handler import generate_consumption_scenarios
from financial_calculations import calculate_scenario_results

# Load input data from Excel file
input_data = load_input_data('solar_financial_model_inputs.xlsx', 'Input Details')
scenario_input_data = load_scenario_data('solar_financial_model_inputs.xlsx', 'Scenarios')

# Initialize results dictionary to store data for each input set and scenario
scenario_results = {}

# Loop through each set of inputs
for input_index, input_row in input_data.iterrows():
    scenarios = generate_consumption_scenarios(input_row, scenario_input_data)
    
    # Loop through each scenario to calculate results for the current set of inputs
    for scenario_name, annual_consumptions in scenarios.items():
        scenario_key = f"Input_{input_index + 1}_{scenario_name}"
        scenario_results[scenario_key] = calculate_scenario_results(input_row, scenario_name, annual_consumptions)

# Export the DataFrames to an Excel file with multiple sheets
with pd.ExcelWriter('solar_financial_model_scenarios_results.xlsx', engine='openpyxl') as writer:
    input_data.to_excel(writer, sheet_name='Input Details', index=False)
    for scenario_key, df in scenario_results.items():
        df.to_excel(writer, sheet_name=scenario_key, index=False)
