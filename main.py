import pandas as pd
from xlsxwriter import Workbook


class RentersInsuranceReport(object):
    def __init__(self, input_url: str) -> None:
        self.dataframe = pd.read_excel(input_url)

        self.dataframe.drop([0, 1, 3], inplace=True, axis=0)

        self.dataframe.columns = self.dataframe.iloc[0]

        self.dataframe.drop([2], inplace=True, axis=0)

        print("DATAFRAME")
        print(self.dataframe)


def main():
    report = RentersInsuranceReport(
        "/home/guilhe/Descargas/PolicySummary02_25_2025.xlsx")


if __name__ == "__main__":
    main()
