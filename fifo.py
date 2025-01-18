#!/usr/bin/env python3
import pandas as pd
import argparse
from openpyxl import load_workbook

class trading_export_sorter:
    def __init__(self, input_file):
        self.df = pd.read_csv(input_file)
        self.buy_and_sell = self.df.loc[self.df['Action'].isin(["Limit buy", "Limit sell", "Market buy", "Market sell"])].groupby('Ticker')

        
    def prepare_main_sheet(self):
        interest_on_cash = self.df.loc[self.df['Action'] == "Interest on cash"]['Total'].sum()
        lending_interest = self.df.loc[self.df['Action'] == "Lending interest"]['Total'].sum()
        deposit = self.df.loc[self.df['Action'] == "Deposit"]['Total'].sum()
        dividend = self.df.loc[self.df['Action'] == "Dividend (Dividend)"]['Total'].sum()
        dividend_manufactured_payment = self.df.loc[self.df['Action'] == "Dividend (Dividend manufactured payment)"]['Total'].sum()
        new_card_cost = self.df.loc[self.df['Action'] == "New card cost"]['Total'].sum()
        all_currency_conversion_fees = self.df["Currency conversion fee"].sum()
        
        data = {
            'Interest on Cash': interest_on_cash,
            'Lending Interest': lending_interest,
            'Deposit': deposit,
            'Dividend': dividend,
            'Dividend Manufactured Payment': dividend_manufactured_payment,
            'New Card Cost': new_card_cost,
            'Currency Conversion Fees': all_currency_conversion_fees
        }
        df_main = pd.DataFrame(list(data.items()))
        return df_main
        
    def prepare_ticker(self, ticker_name):
        columns_to_sum = ['Total', 'No. of shares', 'Result', 'Currency conversion fee']
        columns_to_invert = ['Total', 'No. of shares']

        ticker = self.buy_and_sell.get_group(ticker_name)
        # invert values of buy trades
        ticker.loc[ticker['Action'].str.lower().str.contains('buy'), columns_to_invert] *= -1
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