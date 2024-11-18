import pandas as pd

def generate_consumption_scenarios(input_row, scenario_input_data):
    years_projection = int(input_row['Years Projection'])
    scenarios = {}

    for index, row in scenario_input_data.iterrows():
        scenario_name = row['Scenario Name']
        baseline_consumption = row['Baseline Consumption (kWh/year)']
        percentage_changes = []

        for i in range(1, 6):  # Assuming up to 5 changes for simplicity
            change_key = f'Change {i} Percentage Change'
            start_year_key = f'Change {i} Start Year'
            duration_key = f'Change {i} Duration'
            if change_key in row and not pd.isna(row[change_key]):
                percentage_changes.append({
                    'Percentage Change': row[change_key],
                    'Start Year': int(row[start_year_key]),
                    'Duration': int(row[duration_key])
                })

        scenarios[scenario_name] = generate_scenario(years_projection, baseline_consumption, percentage_changes)
    return scenarios

def generate_scenario(years, baseline, percentage_changes):
    scenarios = []
    for year in range(1, years + 1):
        for change in percentage_changes:
            if year >= change['Start Year'] and year < change['Start Year'] + change['Duration']:
                baseline *= (1 + change['Percentage Change'] / 100)
        scenarios.append(baseline)
    return scenarios
