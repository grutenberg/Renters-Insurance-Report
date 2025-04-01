import customtkinter as tk
from report import RentersInsuranceReport


tk.set_appearance_mode("dark")  # options: "light", "dark", "system"
tk.set_default_color_theme("blue")

class ReporterWindow(tk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Renters Insurance Reporter")
        self.geometry("400x200")

        self.title_screen = tk.CTkLabel(self, text="Renters Insurance Reporter", wraplength=300)
        self.title_screen.pack(pady=20)

        self.file_button = tk.CTkButton(self, text="Select file", command=self.process_report)
        self.file_button.pack()

        self.info_screen = tk.CTkLabel(self, text="Some label", wraplength=300)
        self.info_screen.pack(pady=20)

    def process_report(self):
        filename = tk.filedialog.askopenfilename(title="Select PolicySummary report")

        self.alert(f"Selected file: {filename}")

        try:
            report = RentersInsuranceReport(filename)
        except Exception as e:
            self.alert(f"Error while loading file: {str(e)}")
            return

        try:
            report.get_report()
        except Exception as e:
            self.alert(f"Error while making report: {str(e)}")

        self.alert("Report created successfuly")

    def alert(self, text: str):
        self.info_screen.configure(text=text)


if __name__ == "__main__":
    window = ReporterWindow()

    window.mainloop()
