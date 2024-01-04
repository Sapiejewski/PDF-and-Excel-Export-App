import tkinter as tk
import os
from tkinter import filedialog
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Spacer, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

class ReportGenerator:
    file = ""
    Outputname=""
    
    def __init__(self, root):
        
        self.root = root
        self.root.title("Program do generowania raportów")

        self.label1 = tk.Label(root, text="Wybierz plik do eksportu danych")
        self.label1.grid(row=0, column=0,sticky="n")

        self.button = tk.Button(root, text="Wybierz plik", command=self.choose_file)
        self.button.grid(row=1, column=0,sticky="n")

        self.labelFile = tk.Label(root,text="")
        self.labelFile.grid(row=2,column=0,sticky="n")

        self.label2 = tk.Label(root, text="Podaj nazwę nowego pliku")
        self.label2.grid(row=3, column=0,sticky="n")

        self.entry2 = tk.Entry(root)
        self.entry2.grid(row=4, column=0,sticky="n")

        self.label3 = tk.Label(root, text="Podaj nazwę egzaminu")
        self.label3.grid(row=5, column=0,sticky="n")

        self.entry3 = tk.Entry(root)
        self.entry3.grid(row=6, column=0,sticky="n")

        self.button1= tk.Button(root,text="Eksportuj", command=self.on_button_click)
        self.button1.grid(row=7,column=0,sticky="n")

        self.button1= tk.Button(root,text="Zamknij program", command=self.shutdown)
        self.button1.grid(row=8,column=0,sticky="n")
    def shutdown(self):
        self.root.destroy()

    def choose_file(self):
        file = filedialog.askopenfilename(title="Wybierz plik Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
        self.file=file
        filename = os.path.basename(self.file)
        self.labelFile.config(text=f"Wybrano plik: {filename}")  

    def generate_report(self, data_frame, output_file):
        try:   
            pdf = SimpleDocTemplate(output_file, pagesize=letter)
            story = []
            styles = getSampleStyleSheet()
            pdfmetrics.registerFont(TTFont('Roboto-Black', './Roboto-Light.ttf'))  
            font_name ="Roboto-Black"
            custom_style = ParagraphStyle('CustomStyle', parent=styles['Normal'], fontName=font_name, fontSize=16)

            for index, row in data_frame.iterrows():
                imie = str(row['IMIE'])
                nazwisko = str(row['NAZWISKO'])
                nr_dowodu = str(row['nr dowodu osobistego/paszportu'])
                nr_rejestru = str(row['NR REJESTRU'])

                content = [
                    Paragraph("Urząd Komisji Nadzoru Finansowego", custom_style),
                    Spacer(1, 10),
                    Paragraph(f"{self.entry3.get()}", custom_style),
                    Spacer(1, 10),
                    Paragraph(f"{imie}", custom_style),
                    Spacer(1, 5),
                    Paragraph(f"{nazwisko}", custom_style),
                    Spacer(1, 10),
                    Paragraph(f"<b>Nr Rejestru:</b> {nr_rejestru}", custom_style),
                    Spacer(1, 50),
                ]

                if (index + 1) % 4 == 0:
                    content.append(PageBreak())

                story += content

            pdf.build(story)
        except Exception as e:
            self.show_error_message(f"Wystąpił błąd. Nie można wygenerować raportu")
    def show_error_message(self, message):
        tk.messagebox.showerror("Error", message)
    def on_button_click(self):
        try:
            if not self.file:
                raise ValueError("Nie ma wybranego pliku.")
            if not self.entry3.get():
                raise ValueError("Nie ma podanego egzaminu. ")
            if not self.entry2.get():
                raise ValueError("Nie ma podanej nazwy nowych plików. ")
            path = self.file
            data_frame = pd.read_excel(path)
            exam = self. entry3.get()
            output = self.entry2.get()

            data = data_frame[["NAZWISKO", "IMIE", "nr dowodu osobistego/paszportu", "NR REJESTRU"]]
            data["PODPIS WEJSCIE"] = " "
            data["PODPIS WYJSCIE"] = " "
            data["UWAGI"] = " "

            with pd.ExcelWriter(f"{str(self.entry2.get())}.xlsx", engine="xlsxwriter") as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet('Zeszyt1')
                bold_format = workbook.add_format({'bold': True})
                worksheet.merge_range('A1:G1', f"Nazwa Egzaminu {exam}", bold_format)
                worksheet.merge_range('A2:G2', 'Data', bold_format)
                worksheet.set_row(0, 30)
                worksheet.set_row(1, 30)
                center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                for col_num, value in enumerate(data.columns.values):
                    worksheet.set_column(col_num, col_num, len(str(value)) + 10)
                data.to_excel(writer, sheet_name='Zeszyt1', index=False, startrow=2)

            self.generate_report(data, f"{output}.pdf")
        except Exception as e:
            self.show_error_message(f"Wystąpił błąd w generowaniu raportu: {e}")


def main():
    root = tk.Tk()
    app = ReportGenerator(root)
    root.geometry("500x400")
    root.mainloop()


if __name__ == '__main__':
    main()
