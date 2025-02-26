import pandas as pd
from xlsxwriter import Workbook


class RentersInsuranceReport(object):
    def __init__(self, input_url: str) -> None:
        self.dataframe = pd.read_excel(input_url)

        # Delete first 3 rows
        self.dataframe.drop([0, 1, 3], inplace=True, axis=0)

        # Rename columns with current first row
        self.dataframe.columns = self.dataframe.iloc[0]

        # Delete what was the row with the names of the columns
        self.dataframe.drop([2], inplace=True, axis=0)

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
        self.dataframe.drop(
            ['Occupancy', 'Policy', 'Rental Type', 'Insurance Company', 'Cancellation', 'Reinstate', 'Lease'], inplace=True, axis=1)

        print("DATAFRAME")
        print(self.dataframe)


def main():
    report = RentersInsuranceReport(
        "/home/guilhe/Descargas/PolicySummary02_25_2025.xlsx")


if __name__ == "__main__":
    main()
