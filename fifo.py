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
        print(self.df.columns)
        print(f"Actions in the file: {self.df['Action'].unique()}")
        # invert values of currency conversion fees
        self.df['Currency conversion fee'] *= -1
        # invert values of buy trades
        buy_columns_to_invert = ['Total', 'No. of shares']
        self.df.loc[self.df['Action'].str.lower().str.contains('buy'), buy_columns_to_invert] *= -1
        self.buy_and_sell = self.df.loc[self.df['Action'].isin(["Limit buy", "Limit sell", "Market buy", "Market sell"])].groupby('Ticker')

    def do_work(self, output_file):
        result_sum = 0
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for ticker_name in self.buy_and_sell.groups.keys():
                columns_to_sum = ['Total', 'No. of shares', 'Result', 'Currency conversion fee']
                ticker = self.buy_and_sell.get_group(ticker_name)
                # add a total row
                sums = ticker[columns_to_sum].sum(numeric_only=True)
                sums.name = 'Total'
                ticker = pd.concat([ticker, sums.to_frame().T])
                result_sum += sums['Result']
                print(f"exporting ticker: {ticker_name}")
                ticker.to_excel(writer, sheet_name=ticker_name, index=True)

            def get_sum(column_name):
                return {column_name:self.df.loc[self.df['Action'] == column_name]['Total'].sum()}
            
            data = dict()
            data.update(get_sum('Interest on cash'))
            data.update(get_sum('Lending interest'))
            data.update(get_sum('Deposit'))
            data.update(get_sum('Dividend (Dividend)'))
            data.update(get_sum('Dividend (Dividend manufactured payment)'))
            data.update(get_sum('New card cost'))
            data.update({'Currency Conversion Fees': self.df["Currency conversion fee"].sum()})

            data.update({'Total results': result_sum})

            df_main = pd.DataFrame(list(data.items()))
            df_main.to_excel(writer, sheet_name='Main', index=False, header=False)

        self.adjust_column_widths(output_file)
        self.move_main_sheet_to_front(output_file)
        
    def move_main_sheet_to_front(self, output_file):
        # Load the workbook to move the main sheet to the front
        workbook = load_workbook(output_file)
        sheets = workbook.sheetnames
        main_sheet = workbook['Main']
        workbook.remove(main_sheet)
        workbook.create_sheet('Main', 0)
        workbook.save(output_file)
    
        
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
    parser = argparse.ArgumentParser(description="Process exported data from trading212.")
    parser.add_argument("--output-file", default="output.xlsx", help="Path to the output Excel file")
    parser.add_argument("csv_file", help="Path to the trading212 export CSV file")
    args = parser.parse_args()

    sorter = trading_export_sorter(args.csv_file)
    sorter.do_work(args.output_file)

if __name__ == "__main__":
    main()
