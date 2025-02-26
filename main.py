import pandas as pd
from xlsxwriter import Workbook


class RentersInsuranceReport(object):
    def __init__(self, input_url: str) -> None:
        self.dataframe = pd.read_excel(input_url)
        self.workbook = Workbook("output.xlsx")

        self.columns = [
            'Property Address',
            'Unit',
            'Tenant',
            'Insurance Provider',
            'Policy ID',
            'Liability Coverage',
            'Expiration Date',
            'Days to Expiration',
            'Additional Notes'
        ]

        # Delete first 2 rows
        self.dataframe.drop([0, 1], inplace=True, axis=0)

        # Merge Second with first row
        self.dataframe.iloc[0] = self.dataframe.iloc[0].str.cat(
            self.dataframe.iloc[1].fillna(''), sep=' ').str.strip()

        # Rename columns with current first row
        self.dataframe.columns = self.dataframe.iloc[0]

        # Delete what were the rows with the names of the columns
        self.dataframe.drop([2, 3], inplace=True, axis=0)

        # Renaming tenant id to property id
        last_property = None
        for index, row in self.dataframe.iterrows():
            resident_val = row['Resident']

            if isinstance(resident_val, str) and resident_val.startswith('Property'):
                last_property = resident_val.split('Property: ')[1]
            else:
                row['Resident'] = last_property

        # Delete group rows like Unit and Property
        self.dataframe = self.dataframe.dropna(subset=['Unit'])

        # Delete unnecesary columns
        self.dataframe.drop([
            'Occupancy Status',
            'Policy Status',
            'Rental Type',
            'Policy Title',
            'Lease From',
            'Lease To',
            'Policy Begin',
            'Cancellation Date',
            'Reinstate Date',
            'Master Policy',
            'Pet Endorsement',
            'Interested Party'
        ], inplace=True, axis=1)
        self.dataframe.drop(self.dataframe.columns[7], axis=1, inplace=True)

        # Change Unit and Tenant name orders
        self.dataframe['Name'], self.dataframe['Unit'] = self.dataframe['Unit'], self.dataframe['Name']

        # Add an empty column for Days to Expiration
        self.dataframe.insert(7, "Dummy", None)

        self.dataframe.columns = self.columns

        self.dataframe.fillna('', inplace=True)
        self.dataframe.reset_index(inplace=True, drop=True)

        print("DATAFRAME")
        print(self.dataframe)

    def get_report(self) -> None:

        worksheet = self.workbook.add_worksheet("Renters Insurance MM-DD-YY")

        default_cell_format = self.workbook.add_format({'border': 1, 'bg_color': 'red', "align": "center", "valign": "vcenter"})
        expired_format = self.workbook.add_format({'border': 1, 'bg_color': 'red', "align": "center", "valign": "vcenter"})
        about_to_expire_format = self.workbook.add_format({'border': 1, 'bg_color': 'yellow', "align": "center", "valign": "vcenter"})
        under_insured_format = self.workbook.add_format({'border': 1, 'bg_color': 'blue', "align": "center", "valign": "vcenter"})
        column_name_format = self.workbook.add_format({
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "gray"
        })
        main_title_format = self.workbook.add_format({
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })

        worksheet.write(1, 0, 'About to Expire', about_to_expire_format)
        worksheet.write(2, 0, 'Master Policy', under_insured_format)
        worksheet.write(3, 0, 'Under Insured (<$300,000)', under_insured_format)
        worksheet.write(4, 0, 'Expired', expired_format)

        worksheet.merge_range("E2:H2", "Renter's Insurance", main_title_format)
        worksheet.merge_range("E5:G5", "Report Date:", main_title_format)

        worksheet.write_row(7, 0, self.columns, column_name_format)

        for index, row in self.dataframe.iterrows():
            worksheet.write_row(8+index, 0, row)

        worksheet.hide_gridlines(2)
        worksheet.autofit(max_width=180)

        self.workbook.close()


def main():
    report = RentersInsuranceReport(
        "/home/guilhe/Descargas/PolicySummary02_25_2025.xlsx")
    report.get_report()


if __name__ == "__main__":
    main()
