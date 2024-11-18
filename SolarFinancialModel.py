import numpy as np
import pandas as pd
import numpy_financial as npf
import openpyxl

# Load input data from Excel file
input_data = pd.read_excel('solar_financial_model_inputs.xlsx', sheet_name='Input Details')
scenario_input_data = pd.read_excel('solar_financial_model_inputs.xlsx', sheet_name='Scenarios')

# Initialize results dictionary to store data for each input set and scenario
scenario_results = {}

# Loop through each set of inputs
for input_index, input_row in input_data.iterrows():
    # Solar Details
    capacity_kWp = input_row['Capacity (kWp)']
    specific_yield = input_row['Specific Yield (kWh/kWp/year)']
    performance_drop = input_row['Annual Performance Drop (%)'] / 100
    years_projection = int(input_row['Years Projection'])

    # Financial Parameters
    base_electricity_tariff = input_row['Electricity Tariff (RM/kWh)']
    tnb_buyback_rate = input_row['TNB Buyback Rate (RM/kWh)']
    cost_per_kwp = input_row['Cost per kWp (RM/kWp)']
    additional_structure_cost = input_row['Additional Structure Cost (RM)']
    opex = input_row['OPEX (RM)']
    tariff_hike_percentage = input_row['Tariff Hike Percentage (%)'] / 100
    tariff_hike_interval = int(input_row['Tariff Hike Interval (years)'])
    buyback_hike_percentage = input_row['Buyback Hike Percentage (%)'] / 100
    buyback_hike_interval = int(input_row['Buyback Hike Interval (years)'])
    opex_hike_percentage = input_row['OPEX Hike Percentage (%)'] / 100
    opex_hike_interval = int(input_row['OPEX Hike Interval (years)'])
    opex_start_year = int(input_row['OPEX Start Year'])

    total_investment_cost = capacity_kWp * cost_per_kwp + additional_structure_cost
    base_investment_cost = capacity_kWp * cost_per_kwp  # Without additional installation cost
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

    # Generate Scenarios from Excel file
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
                    'Start Year': int(row[start_year_key]),
                    'Duration': int(row[duration_key])
                })
        scenarios[scenario_name] = generate_consumption_scenarios(years_projection, baseline_consumption, percentage_changes)

    # Loop through each scenario to calculate results for the current set of inputs
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
        cumulative_cash_flows_total = []
        cumulative_cash_flows_base = []
        tnb_buyback_rates = []
        opex_values = []

        # Calculate PV Generation and Annual Consumption for each year
        cumulative_savings_total = -total_investment_cost
        cumulative_savings_base = -base_investment_cost
        initial_investment_total = cumulative_savings_total
        initial_investment_base = cumulative_savings_base

        # Initialize electricity tariff for each scenario
        electricity_tariff = base_electricity_tariff

        for year in range(1, years_projection + 1):
            # Update electricity tariff based on hike interval
            if (year - 1) % tariff_hike_interval == 0 and year > 1:
                electricity_tariff = base_electricity_tariff * ((1 + tariff_hike_percentage) ** ((year - 1) // tariff_hike_interval))
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

            # Cumulative Cash Flow Calculation (with and without additional installation cost)
            annual_cash_flow = consumption_saving + export_saving + tax_saving - opex_values[-1]
            cumulative_savings_total += annual_cash_flow
            cumulative_savings_base += (consumption_saving + export_saving - opex_values[-1])  # Without tax savings and additional installation cost
            cumulative_cash_flows_total.append(cumulative_savings_total)
            cumulative_cash_flows_base.append(cumulative_savings_base)

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
            'Cumulative Cash Flow with Additional Cost (RM)': cumulative_cash_flows_total,
            'Cumulative Cash Flow without Additional Cost (RM)': cumulative_cash_flows_base,
            'Solar Consumption Percentage (%)': solar_consumption_percentages
        }

        scenario_key = f"Input_{input_index + 1}_{scenario_name}"
        scenario_results[scenario_key] = pd.DataFrame(data)

        # Calculate IRR using numpy_financial's irr function (with and without additional installation cost)
        cash_flows_total = [initial_investment_total] + [energy_savings[year] + exported_energy_savings[year] + tax_savings[year] - opex_values[year] for year in range(years_projection)]
        cash_flows_base = [initial_investment_base] + [energy_savings[year] + exported_energy_savings[year] - opex_values[year] for year in range(years_projection)]

        irr_total = npf.irr(cash_flows_total)
        irr_base = npf.irr(cash_flows_base)
        print(f"Input Set {input_index + 1}, Scenario: {scenario_name}, IRR with Additional Cost: {irr_total * 100:.2f}%")
        print(f"Input Set {input_index + 1}, Scenario: {scenario_name}, IRR without Additional Cost: {irr_base * 100:.2f}%")

        # Determine Payback Period (with and without additional installation cost)
        payback_period_total = None
        payback_period_base = None
        for year, cumulative_cash_flow in enumerate(cumulative_cash_flows_total, start=1):
            if cumulative_cash_flow >= 0:
                previous_cash_flow = cumulative_cash_flows_total[year - 2] if year > 1 else initial_investment_total
                payback_period_total = year - 1 + (abs(previous_cash_flow) / (cumulative_cash_flow - previous_cash_flow))
                break

        for year, cumulative_cash_flow in enumerate(cumulative_cash_flows_base, start=1):
            if cumulative_cash_flow >= 0:
                previous_cash_flow = cumulative_cash_flows_base[year - 2] if year > 1 else initial_investment_base
                payback_period_base = year - 1 + (abs(previous_cash_flow) / (cumulative_cash_flow - previous_cash_flow))
                break

        if payback_period_total:
            print(f"Input Set {input_index + 1}, Scenario: {scenario_name}, Payback Period with Additional Cost: {payback_period_total:.2f} years")
        else:
            print(f"Input Set {input_index + 1}, Scenario: {scenario_name}, Payback Period with Additional Cost: Not achieved within the projection period")

        if payback_period_base:
            print(f"Input Set {input_index + 1}, Scenario: {scenario_name}, Payback Period without Additional Cost: {payback_period_base:.2f} years")
        else:
            print(f"Input Set {input_index + 1}, Scenario: {scenario_name}, Payback Period without Additional Cost: Not achieved within the projection period")

# Export the DataFrames to an Excel file with multiple sheets
with pd.ExcelWriter('solar_financial_model_scenarios_results.xlsx', engine='openpyxl') as writer:
    input_data.to_excel(writer, sheet_name='Input Details', index=False)
    for scenario_key, df in scenario_results.items():
        df.to_excel(writer, sheet_name=scenario_key, index=False)

print("Results have been exported to 'solar_financial_model_scenarios_results.xlsx'")
