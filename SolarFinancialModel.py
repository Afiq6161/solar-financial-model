import numpy as np
import pandas as pd
import scipy.optimize as opt
import openpyxl

# Load input data from CSV file
input_data = pd.read_csv('solar_financial_model_inputs.csv')
scenario_input_data = pd.read_csv('solar_consumption_scenarios.csv')

# Solar Details
capacity_kWp = input_data['Capacity (kWp)'][0]
specific_yield = input_data['Specific Yield (kWh/kWp/year)'][0]
performance_drop = input_data['Annual Performance Drop (%)'][0] / 100
years_projection = input_data['Years Projection'][0]

# Financial Parameters
electricity_tariff = input_data['Electricity Tariff (RM/kWh)'][0]
tnb_buyback_rate = input_data['TNB Buyback Rate (RM/kWh)'][0]
cost_per_kwp = input_data['Cost per kWp (RM/kWp)'][0]
additional_structure_cost = input_data['Additional Structure Cost (RM)'][0]
opex = input_data['OPEX (RM)'][0]
tariff_hike_percentage = input_data['Tariff Hike Percentage (%)'][0] / 100
tariff_hike_interval = input_data['Tariff Hike Interval (years)'][0]
buyback_hike_percentage = input_data['Buyback Hike Percentage (%)'][0] / 100
buyback_hike_interval = input_data['Buyback Hike Interval (years)'][0]
opex_hike_percentage = input_data['OPEX Hike Percentage (%)'][0] / 100
opex_hike_interval = input_data['OPEX Hike Interval (years)'][0]
opex_start_year = input_data['OPEX Start Year'][0]

total_investment_cost = capacity_kWp * cost_per_kwp + additional_structure_cost

tax_saving_gita = 0.3 * total_investment_cost  # 30% tax saving from GITA

# Scenario Definitions
def generate_consumption_scenarios(years, baseline, percentage_changes):
    scenarios = []
    for year in range(1, years + 1):
        for change in percentage_changes:
            if year >= change['Start Year'] and year < change['Start Year'] + change['Duration']:
                baseline *= (1 + change['Percentage Change'] / 100)
        scenarios.append(baseline)
    return scenarios

# Generate Scenarios from CSV file
scenarios = {}
for index, row in scenario_input_data.iterrows():
    scenario_name = row['Scenario Name']
    baseline_consumption = row['Baseline Consumption (kWh/year)']
    percentage_changes = []
    for i in range(1, 6):  # Assuming up to 5 changes for simplicity, can be extended
        change_key = f'Change {i} Percentage Change'
        start_year_key = f'Change {i} Start Year'
        duration_key = f'Change {i} Duration'
        if change_key in row and not pd.isna(row[change_key]):
            percentage_changes.append({
                'Percentage Change': row[change_key],
                'Start Year': row[start_year_key],
                'Duration': row[duration_key]
            })
    scenarios[scenario_name] = generate_consumption_scenarios(years_projection, baseline_consumption, percentage_changes)

# Initialize results dictionary to store data for each scenario
scenario_results = {}

# Loop through each scenario to calculate results
for scenario_name, annual_consumptions in scenarios.items():
    # Initialize lists to store the results for this scenario
    pv_generation_rates = []
    pv_generations = []
    consumed_pv_powers = []
    excess_energy_exported = []
    energy_savings = []
    exported_energy_savings = []
    tax_savings = []
    electricity_tariffs = []
    solar_consumption_percentages = []
    cumulative_cash_flows = []
    tnb_buyback_rates = []
    opex_values = []

    # Calculate PV Generation and Annual Consumption for each year
    cumulative_savings = -total_investment_cost
    initial_investment = cumulative_savings

    for year in range(1, years_projection + 1):
        # Update electricity tariff based on hike interval
        if (year - 1) % tariff_hike_interval == 0 and year > 1:
            electricity_tariff *= (1 + tariff_hike_percentage)
        electricity_tariffs.append(electricity_tariff)

        # Update TNB buyback rate based on hike interval
        if (year - 1) % buyback_hike_interval == 0 and year > 1:
            tnb_buyback_rate *= (1 + buyback_hike_percentage)
        tnb_buyback_rates.append(tnb_buyback_rate)

        # Update OPEX based on hike interval, considering OPEX start year
        if year >= opex_start_year:
            if (year - opex_start_year) % opex_hike_interval == 0 and year > opex_start_year:
                opex *= (1 + opex_hike_percentage)
            opex_values.append(opex)
        else:
            opex_values.append(0)

        # PV Power Generation Rate Calculation
        pv_generation_rate = specific_yield * ((1 - performance_drop) ** (year - 1))
        pv_generation_rates.append(pv_generation_rate)

        # PV Power Generation (kWh/year)
        pv_generation = capacity_kWp * pv_generation_rate
        pv_generations.append(pv_generation)

        # Consumed PV Power and Excess Energy Exported
        consumed_pv_power = min(pv_generation, annual_consumptions[year - 1])
        excess_energy = max(0, pv_generation - annual_consumptions[year - 1])
        consumed_pv_powers.append(consumed_pv_power)
        excess_energy_exported.append(excess_energy)

        # Energy Savings Calculations
        consumption_saving = consumed_pv_power * electricity_tariff
        export_saving = excess_energy * tnb_buyback_rate
        energy_savings.append(consumption_saving)
        exported_energy_savings.append(export_saving)

        # Tax Savings from GITA (only applicable in the first year)
        tax_saving = tax_saving_gita if year == 1 else 0
        tax_savings.append(tax_saving)

        # Solar Consumption Percentage
        consumption_percentage = (consumed_pv_power / annual_consumptions[year - 1]) * 100
        solar_consumption_percentages.append(consumption_percentage)

        # Cumulative Cash Flow Calculation
        annual_cash_flow = consumption_saving + export_saving + tax_saving - opex_values[-1]
        cumulative_savings += annual_cash_flow
        cumulative_cash_flows.append(cumulative_savings)

    # Store scenario results in DataFrame
    data = {
        'Year': list(range(1, years_projection + 1)),
        'PV Generation Rate (kWh/kWp/year)': pv_generation_rates,
        'PV Generation (kWh/year)': pv_generations,
        'Annual Consumption (kWh/year)': annual_consumptions,
        'Consumed PV Power (kWh/year)': consumed_pv_powers,
        'Excess Energy Exported (kWh/year)': excess_energy_exported,
        'Electricity Tariff (RM/kWh)': electricity_tariffs,
        'TNB Buyback Rate (RM/kWh)': tnb_buyback_rates,
        'Energy Consumption Saving (RM)': energy_savings,
        'Exported Energy Saving (RM)': exported_energy_savings,
        'Tax Saving from GITA (RM)': tax_savings,
        'OPEX (RM)': opex_values,
        'Cumulative Cash Flow (RM)': cumulative_cash_flows,
        'Solar Consumption Percentage (%)': solar_consumption_percentages
    }

    scenario_results[scenario_name] = pd.DataFrame(data)

    # Calculate IRR using scipy's financial function
    cash_flows = [initial_investment] + [energy_savings[year] + exported_energy_savings[year] + tax_savings[year] - opex_values[year] for year in range(years_projection)]

    def npv(rate, cashflows):
        return sum([cf / (1 + rate) ** i for i, cf in enumerate(cashflows)])

    irr = opt.newton(lambda r: npv(r, cash_flows), 0.1)
    print(f"Scenario: {scenario_name}, IRR: {irr * 100:.2f}%")

    # Determine Payback Period
    payback_period = None
    for year, cumulative_cash_flow in enumerate(cumulative_cash_flows, start=1):
        if cumulative_cash_flow >= 0:
            previous_cash_flow = cumulative_cash_flows[year - 2] if year > 1 else initial_investment
            payback_period = year - 1 + (abs(previous_cash_flow) / (cumulative_cash_flow - previous_cash_flow))
            break

    if payback_period:
        print(f"Scenario: {scenario_name}, Payback Period: {payback_period:.2f} years")
    else:
        print(f"Scenario: {scenario_name}, Payback Period: Not achieved within the projection period")

    # Calculate average solar consumption percentage
    average_solar_consumption_percentage = sum(solar_consumption_percentages) / years_projection
    print(f"Scenario: {scenario_name}, Average Solar Consumption Percentage: {average_solar_consumption_percentage:.2f}%")

    # Store summary data for this scenario
    summary_data = {
        'Metric': ['Payback Period (years)', 'Internal Rate of Return (IRR, %)', 'Average Solar Consumption Percentage (%)'],
        'Value': [payback_period if payback_period else 'Not achieved', irr * 100, average_solar_consumption_percentage]
    }
    scenario_results[f"{scenario_name}_summary"] = pd.DataFrame(summary_data)

# Export the DataFrames to an Excel file with multiple sheets
with pd.ExcelWriter('solar_financial_model_scenarios_results.xlsx', engine='openpyxl') as writer:
    input_data.to_excel(writer, sheet_name='Input Details', index=False)
    for scenario_name, df in scenario_results.items():
        df.to_excel(writer, sheet_name=scenario_name, index=False)

print("Results have been exported to 'solar_financial_model_scenarios_results.xlsx'")
