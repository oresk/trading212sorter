#!/usr/bin/env python3
import pandas as pd
import argparse
from openpyxl import load_workbook

# TODO:
# - do the fifo calculation
# - keep track of conversion gains/losses


# negative values decrement the cash on balance
# so buy trades are negative
# new card cost is negative
# currency conversion fee is negative
#
# positive values increment the cash on balance
# so sell trades are positive
# interest on cash is positive
# lending interest is positive
# deposit is positive
# dividend is positive
# dividend manufactured payment is positive


class trading_export_sorter:
    def __init__(self, input_file):
        self.df = pd.read_csv(input_file)
        self.prepare_data()
        self.buy_and_sell = self.df.loc[self.df['Action'].isin(["Limit buy", "Limit sell", "Market buy", "Market sell"])].groupby('Ticker')

    def prepare_data(self):
        # invert values of currency conversion fees
        self.df['Currency conversion fee'] *= -1
        # invert values of buy trades
        buy_columns_to_invert = ['Total', 'No. of shares']
        self.df.loc[self.df['Action'].str.lower().str.contains('buy'), buy_columns_to_invert] *= -1
        
        
    def prepare_main_sheet(self):
        def get_sum(column_name):
            return {column_name:self.df.loc[self.df['Action'] == column_name]['Total'].sum()}
        
        data = {}
        data.update(get_sum("Interest on cash"))
        data.update(get_sum('Lending Interest'))
        data.update(get_sum('Deposit'))
        data.update(get_sum('Dividend'))
        data.update(get_sum('Dividend Manufactured Payment'))
        data.update(get_sum('New Card Cost'))
        data.update({'Currency Conversion Fees': self.df["Currency conversion fee"].sum()})

        df_main = pd.DataFrame(list(data.items()))
        return df_main
        
    def prepare_ticker(self, ticker_name):
        columns_to_sum = ['Total', 'No. of shares', 'Result', 'Currency conversion fee']
        ticker = self.buy_and_sell.get_group(ticker_name)
        # add a total row
        sums = ticker[columns_to_sum].sum(numeric_only=True)
        sums.name = 'Total'
        return pd.concat([ticker, sums.to_frame().T])
    
    def export_all(self, output_file):
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            self.prepare_main_sheet().to_excel(writer, sheet_name='Main', index=False, header=False)
            for ticker_name in self.buy_and_sell.groups.keys():
                print(f"exporting ticker: {ticker_name}")
                self.prepare_ticker(ticker_name).to_excel(writer, sheet_name=ticker_name, index=True)
        self.adjust_column_widths(output_file)
        
    def adjust_column_widths(self, output_file): 
        # Load the workbook to adjust column widths
        workbook = load_workbook(output_file)
        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            for column_cells in worksheet.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = length + 3
        workbook.save(output_file)


def main():
    parser = argparse.ArgumentParser(description="Process trading data from CSV.")
    parser.add_argument("--output-file", default="output.xlsx", help="Path to the output Excel file")
    parser.add_argument("csv_file", help="Path to the trading212 export CSV file")
    args = parser.parse_args()

    sorter = trading_export_sorter(args.csv_file)
    sorter.export_all(args.output_file)

if __name__ == "__main__":
    main()
