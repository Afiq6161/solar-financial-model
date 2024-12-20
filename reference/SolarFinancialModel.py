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
    base_tnb_buyback_rate = input_row['TNB Buyback Rate (RM/kWh)']
    cost_per_kwp = input_row['Cost per kWp (RM/kWp)']
    additional_structure_cost = input_row['Additional Structure Cost (RM)']
    base_opex = input_row['OPEX (RM)']
    tariff_hike_percentage = input_row['Tariff Hike Percentage (%)'] / 100
    tariff_hike_interval = int(input_row['Tariff Hike Interval (years)'])
    buyback_hike_percentage = input_row['Buyback Hike Percentage (%)'] / 100
    buyback_hike_interval = int(input_row['Buyback Hike Interval (years)'])
    opex_hike_percentage = input_row['OPEX Hike Percentage (%)'] / 100
    opex_hike_interval = int(input_row['OPEX Hike Interval (years)'])
    opex_start_year = int(input_row['OPEX Start Year'])
    energy_export_allowed = input_row['Energy Export Allowed']

    base_investment_cost = capacity_kWp * cost_per_kwp  # Without additional installation cost
    total_investment_cost = base_investment_cost + additional_structure_cost
    tax_saving_gita = 0.3 * base_investment_cost  # 30% tax saving from GITA

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
        capital_expenses = []
        capital_expenses_base = []
        total_expenses = []
        total_expenses_base = []
        total_incomes = []

        # Calculate PV Generation and Annual Consumption for each year
        cumulative_savings_total = 0  # Starting at 0 since initial investment is accounted in capital expense
        cumulative_savings_base = 0  # Starting at 0 since initial investment is accounted in capital expense

        # Initialize tariff for each scenario
        electricity_tariff = base_electricity_tariff
        tnb_buyback_rate = base_tnb_buyback_rate
        opex = base_opex

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
            excess_energy = max(0, pv_generation - annual_consumptions[year - 1]) if energy_export_allowed else 0
            consumed_pv_powers.append(consumed_pv_power)
            excess_energy_exported.append(excess_energy)

            # Energy Savings Calculations
            consumption_saving = consumed_pv_power * electricity_tariff
            export_saving = excess_energy * tnb_buyback_rate if energy_export_allowed else 0
            energy_savings.append(consumption_saving)
            exported_energy_savings.append(export_saving)

            # Tax Savings from GITA (only applicable in the first year)
            tax_saving = tax_saving_gita if year == 1 else 0
            tax_savings.append(tax_saving)

            # Solar Consumption Percentage
            consumption_percentage = (consumed_pv_power / annual_consumptions[year - 1]) * 100
            solar_consumption_percentages.append(consumption_percentage)

            # Capital Expenses (initial investment cost)
            capital_expense = total_investment_cost if year == 1 else 0
            capital_expense_base = base_investment_cost if year == 1 else 0
            capital_expenses.append(capital_expense)
            capital_expenses_base.append(capital_expense_base)

            # Total Expenses for the year (OPEX + Capital Expenses)
            total_expense = opex_values[-1] + capital_expense
            total_expense_base = opex_values[-1] + capital_expense_base
            total_expenses.append(total_expense)
            total_expenses_base.append(total_expense_base)


            # Total Income for the year (Energy Savings + Exported Energy Savings + Tax Savings)
            total_income = consumption_saving + export_saving + tax_saving
            total_incomes.append(total_income)

            # Cumulative Cash Flow Calculation (with and without additional installation cost)
            annual_cash_flow = total_income - total_expense
            cumulative_savings_total += annual_cash_flow
            cumulative_savings_base += (total_income - total_expense_base)  # Without tax savings and additional installation cost
            cumulative_cash_flows_total.append(cumulative_savings_total)
            cumulative_cash_flows_base.append(cumulative_savings_base)

        # Store scenario results in DataFrame
        data = {
            'Year': list(range(1, years_projection + 1)),
            'PV Generation Rate (kWh/kWp/year)': pv_generation_rates,
            'PV Generation (kWh/year)': pv_generations,
            'Consumed PV Power (kWh/year)': consumed_pv_powers,
            'Excess Energy Exported (kWh/year)': excess_energy_exported,
            'Electricity Tariff (RM/kWh)': electricity_tariffs,
            'TNB Buyback Rate (RM/kWh)': tnb_buyback_rates,
            'Energy Consumption Saving (RM)': energy_savings,
            'Exported Energy Saving (RM)': exported_energy_savings,
            'Tax Saving from GITA (RM)': tax_savings,
            'OPEX (RM)': opex_values,
            'Capital Expense (base) (RM)': capital_expenses_base,
            'Capital Expense (RM)': capital_expenses,
            'Total Expense (base)(RM)': total_expenses_base,
            'Total Expense (RM)': total_expenses,
            'Total Income (RM)': total_incomes,
            'Cumulative Cash Flow (base) (RM)': cumulative_cash_flows_base,
            'Cumulative Cash Flow (RM)': cumulative_cash_flows_total,
        }

        scenario_key = f"Input_{input_index + 1}_{scenario_name}"
        scenario_results[scenario_key] = pd.DataFrame(data)

        # Calculate IRR using numpy_financial's irr function (with and without additional installation cost)
        cash_flows_total = [total_incomes[year] - total_expenses[year] for year in range(years_projection)]
        cash_flows_base = [total_incomes[year] - total_expenses_base[year] for year in range(years_projection)]

        irr_total = npf.irr(cash_flows_total)
        irr_base = npf.irr(cash_flows_base)
        print(f"Input Set {input_index + 1}, Scenario: {scenario_name}, IRR with Additional Cost: {irr_total * 100:.2f}%")
        print(f"Input Set {input_index + 1}, Scenario: {scenario_name}, IRR without Additional Cost: {irr_base * 100:.2f}%")

        # Determine Payback Period (with and without additional installation cost)
        payback_period_total = None
        payback_period_base = None
        for year, cumulative_cash_flow in enumerate(cumulative_cash_flows_total, start=1):
            if cumulative_cash_flow >= 0:
                previous_cash_flow = cumulative_cash_flows_total[year - 2] if year > 1 else -total_investment_cost
                payback_period_total = year - 1 + (abs(previous_cash_flow) / (cumulative_cash_flow - previous_cash_flow))
                break

        for year, cumulative_cash_flow in enumerate(cumulative_cash_flows_base, start=1):
            if cumulative_cash_flow >= 0:
                previous_cash_flow = cumulative_cash_flows_base[year - 2] if year > 1 else -base_investment_cost
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

        # Calculate average solar consumption percentage
        average_solar_consumption_percentage = sum(solar_consumption_percentages) / years_projection
        print(f"Scenario: {scenario_name}, Average Solar Consumption Percentage: {average_solar_consumption_percentage:.2f}%")

        # Store summary data for this scenario
        summary_data = {
            'Scenario': scenario_name,
            'Input Set': input_index + 1,
            'Cost per kWp (RM/kWp)': cost_per_kwp,
            'Installed Capacity (kWp)': capacity_kWp,
            'PV Investment Cost (RM)': base_investment_cost,
            'Overall Investment Cost with Structure (RM)': total_investment_cost,
            'Average Annual Building Load (kWh/year)': sum(annual_consumptions) / years_projection,
            'Performance Drop (%)': performance_drop * 100,
            'PV IRR (%)': irr_base * 100,
            'Overall IRR (%)': irr_total * 100,
            'PV Payback Period (years)': payback_period_base if payback_period_base else 'Not achieved',
            'Overall Payback Period (years)': payback_period_total if payback_period_total else 'Not achieved',
            'Average Annual Solar Yield (kWh/year)': sum(pv_generations) / years_projection,
            'Consumption Percentage (%)': average_solar_consumption_percentage,
            'Specific Yield (kWh/kWp/year)': specific_yield
        }
        scenario_results[f"{scenario_name}_summary_{input_index + 1}"] = pd.DataFrame([summary_data])

# Export the DataFrames to an Excel file with a summary table and scenario details
with pd.ExcelWriter('solar_financial_model_scenarios_results.xlsx', engine='openpyxl') as writer:
    input_data.to_excel(writer, sheet_name='Input Details', index=False)
    summary_list = []
    for scenario_key, df in scenario_results.items():
        if 'summary' in scenario_key:
            summary_list.append(df)
        else:
            df.to_excel(writer, sheet_name=scenario_key, index=False)
    if summary_list:
        pd.concat(summary_list).to_excel(writer, sheet_name='Summary', index=False)

print("Results have been exported to 'solar_financial_model_scenarios_results.xlsx'")
