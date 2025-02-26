import pandas as pd
from xlsxwriter import Workbook


class RentersInsuranceReport(object):
    def __init__(self, input_url: str) -> None:
        self.dataframe = pd.read_excel(input_url)

        # Delete first 2 rows
        self.dataframe.drop([0, 1], inplace=True, axis=0)

        # Merge Second with first row
        self.dataframe.iloc[0] = self.dataframe.iloc[0].str.cat(
            self.dataframe.iloc[1].fillna(''), sep=' ').str.strip()

        # Rename columns with current first row
        self.dataframe.columns = self.dataframe.iloc[0]

        # Delete what were the rows with the names of the columns
        self.dataframe.drop([2, 3], inplace=True, axis=0)

        print(self.dataframe)
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
            'Insurance Company',
            'Policy Number',
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
        self.dataframe.drop(self.dataframe.columns[5], axis=1, inplace=True)

        self.dataframe['Name'], self.dataframe['Unit'] = self.dataframe['Unit'], self.dataframe['Name']

        print("DATAFRAME")
        print(self.dataframe)


def main():
    report = RentersInsuranceReport(
        "/home/guilhe/Descargas/PolicySummary02_25_2025.xlsx")


if __name__ == "__main__":
    main()
