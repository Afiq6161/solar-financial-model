import pandas as pd
import numpy_financial as npf
import openpyxl

def load_input_file(filepath):
    input_data = pd.read_excel(filepath, sheet_name='Input Details')
    scenario_input_data = pd.read_excel(filepath, sheet_name='Scenarios')
    return input_data, scenario_input_data

def generate_consumption_scenarios(years, baseline, percentage_changes):
    scenarios = []
    for year in range(1, years + 1):
        for change in percentage_changes:
            if year >= change['Start Year'] and year < change['Start Year'] + change['Duration']:
                baseline *= (1 + change['Percentage Change'] / 100)
        scenarios.append(baseline)
    return scenarios

def run_model(input_data, scenario_input_data, progress_var):
    scenario_results = {}

    for input_index, input_row in input_data.iterrows():
        capacity_kWp = input_row['Capacity (kWp)']
        specific_yield = input_row['Specific Yield (kWh/kWp/year)']
        performance_drop = input_row['Annual Performance Drop (%)'] / 100
        years_projection = int(input_row['Years Projection'])
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

        base_investment_cost = capacity_kWp * cost_per_kwp
        total_investment_cost = base_investment_cost + additional_structure_cost
        tax_saving_gita = 0.3 * base_investment_cost

        scenarios = {}
        for index, row in scenario_input_data.iterrows():
            scenario_name = row['Scenario Name']
            baseline_consumption = row['Baseline Consumption (kWh/year)']
            percentage_changes = []
            for i in range(1, 6):
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

        for scenario_name, annual_consumptions in scenarios.items():
            pv_generation_rates = []
            pv_generations = []
            consumed_pv_powers = []
            excess_energy_exported = []
            energy_savings = []
            exported_energy_savings = []
            tax_savings = []
            electricity_tariffs = []
            cumulative_cash_flows_total = []
            cumulative_cash_flows_base = []
            tnb_buyback_rates = []
            opex_values = []
            capital_expenses = []
            capital_expenses_base = []
            total_expenses = []
            total_expenses_base = []
            total_incomes = []

            cumulative_savings_total = 0
            cumulative_savings_base = 0

            electricity_tariff = base_electricity_tariff
            tnb_buyback_rate = base_tnb_buyback_rate
            opex = base_opex

            for year in range(1, years_projection + 1):
                if (year - 1) % tariff_hike_interval == 0 and year > 1:
                    electricity_tariff = base_electricity_tariff * ((1 + tariff_hike_percentage) ** ((year - 1) // tariff_hike_interval))
                electricity_tariffs.append(electricity_tariff)

                if (year - 1) % buyback_hike_interval == 0 and year > 1:
                    tnb_buyback_rate *= (1 + buyback_hike_percentage)
                tnb_buyback_rates.append(tnb_buyback_rate)

                if year >= opex_start_year:
                    if (year - opex_start_year) % opex_hike_interval == 0 and year > opex_start_year:
                        opex *= (1 + opex_hike_percentage)
                    opex_values.append(opex)
                else:
                    opex_values.append(0)

                pv_generation_rate = specific_yield * ((1 - performance_drop) ** (year - 1))
                pv_generation_rates.append(pv_generation_rate)

                pv_generation = capacity_kWp * pv_generation_rate
                pv_generations.append(pv_generation)

                consumed_pv_power = min(pv_generation, annual_consumptions[year - 1])
                excess_energy = max(0, pv_generation - annual_consumptions[year - 1]) if energy_export_allowed else 0
                consumed_pv_powers.append(consumed_pv_power)
                excess_energy_exported.append(excess_energy)

                consumption_saving = consumed_pv_power * electricity_tariff
                export_saving = excess_energy * tnb_buyback_rate if energy_export_allowed else 0
                energy_savings.append(consumption_saving)
                exported_energy_savings.append(export_saving)

                tax_saving = tax_saving_gita if year == 1 else 0
                tax_savings.append(tax_saving)

                capital_expense = total_investment_cost if year == 1 else 0
                capital_expense_base = base_investment_cost if year == 1 else 0
                capital_expenses.append(capital_expense)
                capital_expenses_base.append(capital_expense_base)

                total_expense = opex_values[-1] + capital_expense
                total_expense_base = opex_values[-1] + capital_expense_base
                total_expenses.append(total_expense)
                total_expenses_base.append(total_expense_base)

                total_income = consumption_saving + export_saving + tax_saving
                total_incomes.append(total_income)

                annual_cash_flow = total_income - total_expense
                cumulative_savings_total += annual_cash_flow
                cumulative_savings_base += (total_income - total_expense_base)
                cumulative_cash_flows_total.append(cumulative_savings_total)
                cumulative_cash_flows_base.append(cumulative_savings_base)

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

            progress_var.set(progress_var.get() + (30 / len(input_data) / len(scenarios)))

    return scenario_results

def save_results(results, filepath):
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        for scenario_key, df in results.items():
            df.to_excel(writer, sheet_name=scenario_key, index=False)
