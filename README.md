# Solar Financial Model Calculator - README

This README provides step-by-step instructions on how to use the Solar Financial Model Calculator. The code calculates solar financial metrics such as IRR, payback period, energy savings, and more based on input parameters provided by the user. The results are exported to an Excel file for easy analysis.

## Prerequisites

Before running the code, ensure you have the following installed:

1. **Python 3.7+**
2. **Libraries**:
   - `numpy`
   - `pandas`
   - `numpy_financial`
   - `openpyxl`

To install the required Python packages, run:

```sh
pip install numpy pandas numpy-financial openpyxl
```

## File Structure

- **solar_financial_model_inputs.xlsx**: An Excel file containing the input data for the solar model, including financial parameters and scenarios.
- **SolarFinancialModel.py**: The Python script that reads the input data, processes it, and exports the results.
- **solar_financial_model_scenarios_results.xlsx**: The output file that contains the results of the financial calculations for each input and scenario.

## Input File Structure

The input data is provided in the `solar_financial_model_inputs.xlsx` file, which has two sheets:

1. **Input Details**: This sheet contains the solar and financial parameters for each scenario, such as:
   - Capacity (kWp)
   - Specific Yield (kWh/kWp/year)
   - Annual Performance Drop (%)
   - Years Projection
   - Electricity Tariff (RM/kWh)
   - TNB Buyback Rate (RM/kWh)
   - Cost per kWp (RM/kWp)
   - Additional Structure Cost (RM)
   - OPEX (RM)
   - Tariff Hike Percentage (%) and Interval (years)
   - Buyback Rate Hike Percentage (%) and Interval (years)
   - OPEX Hike Percentage (%) and Interval (years)
   - OPEX Start Year

2. **Scenarios**: This sheet contains different electricity consumption scenarios, including:
   - Scenario Name
   - Baseline Consumption (kWh/year)
   - Up to 5 changes in percentage change, start year, and duration.

## Running the Code

1. **Prepare the Input File**: Fill in the `solar_financial_model_inputs.xlsx` file with your input data. You can provide multiple sets of inputs and scenarios.

2. **Run the Python Script**:
   - Open your command line or terminal.
   - Navigate to the directory containing the `SolarFinancialModel.py` file.
   - Activate your virtual environment (if you have one).
   - Run the script using the following command:

   ```sh
   python SolarFinancialModel.py
   ```

3. **View the Results**:
   - After running the script, the results will be saved in an Excel file called `solar_financial_model_scenarios_results.xlsx`.
   - The file will have multiple sheets: one for each scenario and a summary of the input details.

## Outputs Explained

- **Year-by-Year Analysis**: The output includes year-by-year details for each scenario, such as:
  - PV Generation Rate (kWh/kWp/year)
  - PV Generation (kWh/year)
  - Annual Consumption (kWh/year)
  - Consumed PV Power (kWh/year)
  - Excess Energy Exported (kWh/year)
  - Energy Consumption Savings, Exported Energy Savings, Tax Savings, etc.
  - Cumulative Cash Flow (with and without additional cost)

- **Summary Metrics**:
  - **IRR with Additional Cost**: The internal rate of return considering all investment costs.
  - **IRR without Additional Cost**: The internal rate of return excluding additional installation costs.
  - **Payback Period**: The estimated time required to recover the investment, both with and without additional costs.

## Example

To see how the calculator works, follow these steps:

1. Open `solar_financial_model_inputs.xlsx` and edit the values in the **Input Details** and **Scenarios** sheets to reflect your specific solar project data.
2. Save and close the Excel file.
3. Run the Python script (`python SolarFinancialModel.py`), and wait for it to complete.
4. Open the newly generated `solar_financial_model_scenarios_results.xlsx` to analyze the results.

## Troubleshooting

1. **Python not recognized**: Ensure Python is installed and added to your system's PATH.
2. **Missing Packages**: If the script fails due to missing packages, run the following command:

   ```sh
   pip install -r requirements.txt
   ```

3. **Runtime Errors**: If you encounter runtime errors, check your input data to ensure all required fields are correctly filled out and that data types are consistent.

## Notes

- Ensure the input data is realistic to avoid convergence issues when calculating the IRR.
- The IRR calculation uses the `numpy_financial` library, which is specifically designed for financial calculations.

## License

This project is open source and available under the [MIT License](LICENSE).

## Contact

If you have any questions or need further assistance, please contact the developer at [your_email@example.com].

