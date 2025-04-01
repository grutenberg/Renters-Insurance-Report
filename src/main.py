from report import RentersInsuranceReport


def main():
    
    # presentation
    print("Wellcome to the Renters Insurance Reporter")
    print("please, insert the absolute <path> to the PolicySummary Yardi report")
    
    selected_path = input(">")

    try:
        report = RentersInsuranceReport(selected_path)
    except Exception as e:
        print(f"Error while trying to open the PolicySummary: {str(e)}")

    try:
        report.get_report()
    except Exception as e:
        print("Fatal error while trying to make Renters Insurance Report")


if __name__ == "__main__":
    main()
