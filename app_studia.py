import os
import tkinter as tk
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

        self.button1= tk.Button(root,text="Eksportuj", command=self.on_button_click)
        self.button1.grid(row=7,column=0,sticky="n")

        self.labelFinished = tk.Label(root, text="")
        self.labelFinished.grid(row=8, column=0,sticky="n")

        self.button1= tk.Button(root,text="Zamknij program", command=self.shutdown)
        self.button1.grid(row=9,column=0,sticky="s")
        root.grid_columnconfigure(0, weight=1)


    def shutdown(self):
        self.root.destroy()

    def choose_file(self):
        file = filedialog.askopenfilename(title="Wybierz plik Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
        self.file=file
        filename = os.path.basename(self.file)
        self.labelFile.config(text=f"Wybrano plik: {filename}")  

    def generate_report(self, data_frame, output_file):
            pdf = SimpleDocTemplate(output_file, pagesize=letter)
            story = []
            styles = getSampleStyleSheet()
            pdfmetrics.registerFont(TTFont('Roboto-Black', './Roboto-Black.ttf'))  
            font_name ="Roboto-Black"
            custom_style = ParagraphStyle('CustomStyle', parent=styles['Normal'], fontName=font_name, fontSize=16)
            
            labels=[]   
            content=[]

            for label,_ in data_frame.items():
                labels.append(label)
            print(labels)
            array= data_frame.to_numpy()
            for item in array:
                for i in range(len(item)):
                    content.append(Paragraph(f"{labels[i]}: {item[i]}",custom_style))
                    content.append(Spacer(1,10))
                content.append(PageBreak())
            story+=content
            pdf.build(story)

        
    def show_error_message(self, message):
        tk.messagebox.showerror("Error", message)
    def on_button_click(self):
        try:
            if not self.file:
                raise ValueError("Nie ma wybranego pliku.")
            if not self.entry2.get():
                raise ValueError("Nie ma podanej nazwy nowego pliku. ")
            path = self.file
            data_frame = pd.read_excel(path)
            output = self.entry2.get()
            self.generate_report(data_frame, f"{output}.pdf")
            self.labelFinished.config(text=f"Stowrzono plik: {self.Outputname}.pdf")  
        except Exception as e:
            self.show_error_message(f"Wystąpił błąd w generowaniu raportu: {e}")
def main():
    root = tk.Tk()
    app = ReportGenerator(root)
    root.geometry("300x330")
    root.minsize(200,220)
    root.mainloop()


if __name__ == '__main__':
    main()
