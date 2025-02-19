import PyPDF2
import re
import pandas as pd
from tkinter import Tk, Button, Label, filedialog, messagebox, ttk, Scrollbar, END, Entry, Frame
from datetime import datetime
import sqlite3
from tkinter import ttk

# Add this after your imports
def create_database():
    conn = sqlite3.connect('payslips.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS payslips
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  periodo TEXT,
                  valor_liquido DECIMAL(10,2),
                  subs_refeicao DECIMAL(10,2),
                  kms DECIMAL(10,2),
                  dias DECIMAL(10,2),
                  descontos DECIMAL(10,2))''')
    conn.commit()
    conn.close()


def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF file."""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''
            for page in reader.pages:
                text += page.extract_text() or ''
            return text.strip()
    except Exception as e:
        messagebox.showerror("Error", f"Could not read the PDF file: {str(e)}")
        return None

def convert_period_format(period):
    """Converts MM/YY to Month Year format"""
    if period == "Not Found":
        return period
        
    try:
        month, year = period.split('/')
        # Dictionary for month conversion to Portuguese
        months = {
            '01': 'Janeiro', '02': 'Fevereiro', '03': 'Março',
            '04': 'Abril', '05': 'Maio', '06': 'Junho',
            '07': 'Julho', '08': 'Agosto', '09': 'Setembro',
            '10': 'Outubro', '11': 'Novembro', '12': 'Dezembro'
        }
        return f"{months[month]} 20{year}"
    except:
        return period

def parse_pdf_data(text):
    def format_currency(value):
        try:
            # Convert string to float, handling both comma and dot separators
            value = value.replace('.', '').replace(',', '.')
            number = float(value)
            # Format with 2 decimal places and convert back to Portuguese format
            return f"{number:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
        except:
            return "0,00"
    """Parses the extracted text for the specific fields."""
    data = {}

    subs_values = ["10,20", "9,60", "8,32"]
    subs_refeicao_value = "0,00"

    # Improved regex patterns based on the payslip format
    periodo_pattern = r"Salário Base.*?(\d{2}/\d{2})"  # Período
    valor_liquido_pattern = r"Valor\s*Líquido\s*(\d+[.,]\d+)"
    for subs_value in subs_values:
        pattern = fr"{subs_value}.*?(\d+[.,]\d+)"
        match = re.search(pattern, text)
        if match:
            subs_refeicao_value = match.group(1)
            break
    dias_pattern = r"Subs\.\s*Refeição\s*\(Cartão\).*?(\d+[.,]\d+)"
    descontos_pattern = r"Totais.*?(\d+[.,]\d+)\s*$"

    # Extract values
    periodo_match = re.search(periodo_pattern, text, re.IGNORECASE)
    valor_liquido_match = re.search(valor_liquido_pattern, text, re.IGNORECASE)
    
    dias_matches = re.findall(dias_pattern, text, re.IGNORECASE)
    descontos_match = re.search(descontos_pattern, text, re.IGNORECASE | re.MULTILINE)

    # Store matched values
    data['Período'] = periodo_match.group(1) if periodo_match else "Not Found"
    data['Valor Líquido'] = format_currency(valor_liquido_match.group(1)) if valor_liquido_match else "0,00"
    data['Subs. Refeição'] = subs_refeicao_value
    data['Descontos'] = format_currency(descontos_match.group(1)) if descontos_match else "0,00"
    data['Dias'] = dias_matches[-1] if dias_matches else "0,00"
    period = periodo_match.group(1) if periodo_match else "Not Found"
    data['Período'] = convert_period_format(period)

    return data

class PDFAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Payslip Analyzer")
        self.root.geometry("1200x800")
        self.root.configure(bg="#2e2e2e")
        self.data = []

        # Configure style
        self.configure_style()

        # Top frame for controls
        self.control_frame = Frame(root, bg="#2e2e2e")
        self.control_frame.pack(fill="x", padx=10, pady=5)

        # Km multiplier input
        Label(self.control_frame, text="Km Multiplier:", fg="white", bg="#2e2e2e", font=("Arial", 12)).pack(side="left", padx=5)
        self.km_multiplier = Entry(self.control_frame, width=5, font=("Arial", 12))
        self.km_multiplier.pack(side="left", padx=5)
        self.km_multiplier.insert(0, "9")  # Default value

        # Buttons
        self.import_button = Button(self.control_frame, text="Import PDFs", command=self.import_pdfs, 
                                  bg="#444444", fg="white", font=("Arial", 12))
        self.import_button.pack(side="left", padx=10)

        self.export_button = Button(self.control_frame, text="Export to Excel", command=self.export_data,
                                  bg="#444444", fg="white", font=("Arial", 12))
        self.export_button.pack(side="left", padx=10)

        # Treeview
        self.tree_frame = ttk.Frame(root)
        self.tree_frame.pack(pady=10, fill="both", expand=True)

        self.scrollbar = Scrollbar(self.tree_frame)
        self.scrollbar.pack(side="right", fill="y")

        self.treeview = ttk.Treeview(
            self.tree_frame,
            columns=("Período", "Valor Líquido", "Subs. Refeição", "Kms", "Dias", "Descontos", "Delete"),
            show="headings",
            yscrollcommand=self.scrollbar.set,
            style="Custom.Treeview"
        )

        # Configure columns
        columns = {
            "Período": 150,
            "Valor Líquido": 150,
            "Subs. Refeição": 150,
            "Kms": 150,
            "Dias": 100,
            "Descontos": 150,
            "Delete": 50
        }

        for col, width in columns.items():
            self.treeview.heading(col, text=col)
            self.treeview.column(col, width=width)
            # Center all columns except Período
            if col != "Período":
                self.treeview.column(col, anchor="center")
            else:
                self.treeview.column(col, anchor="w")

        self.treeview.pack(side="left", fill="both", expand=True)
        self.scrollbar.config(command=self.treeview.yview)

         # Bind click event for delete button
        self.treeview.bind('<ButtonRelease-1>', self.handle_click)
    
     # Create database and load data at startup
        create_database()
        self.load_from_db()

    def save_to_db(self, data):
        conn = sqlite3.connect('payslips.db')
        c = conn.cursor()
        
        # Check if period already exists
        c.execute('SELECT id, periodo FROM payslips WHERE periodo = ?', (data['Período'],))
        existing_record = c.fetchone()
        
        if existing_record:
            if messagebox.askyesno("Duplicate Entry", 
                                f"Period {data['Período']} already exists.\nDo you want to replace it with the new entry?"):
                # Delete old entry
                c.execute('DELETE FROM payslips WHERE id = ?', (existing_record[0],))
                # Insert new entry
                c.execute('''INSERT INTO payslips 
                            (periodo, valor_liquido, subs_refeicao, kms, dias, descontos)
                            VALUES (?, ?, ?, ?, ?, ?)''',
                        (data['Período'], data['Valor Líquido'], data['Subs. Refeição'],
                        data['Kms'], data['Dias'], data['Descontos']))
                conn.commit()
                conn.close()
                return True
            else:
                conn.close()
                return False
                
        # If period doesn't exist, insert new record
        c.execute('''INSERT INTO payslips 
                    (periodo, valor_liquido, subs_refeicao, kms, dias, descontos)
                    VALUES (?, ?, ?, ?, ?, ?)''',
                (data['Período'], data['Valor Líquido'], data['Subs. Refeição'],
                data['Kms'], data['Dias'], data['Descontos']))
        conn.commit()
        conn.close()
        return True

    def load_from_db(self):
        self.treeview.delete(*self.treeview.get_children())
        conn = sqlite3.connect('payslips.db')
        c = conn.cursor()
        c.execute('SELECT * FROM payslips')
        rows = c.fetchall()

        # Convert period to sortable date and sort
        def period_to_sortable_date(row):
            month_year = row[1]  # periodo is in column 1
            month, year = month_year.split()
            months = {
                'Janeiro': 1, 'Fevereiro': 2, 'Março': 3,
                'Abril': 4, 'Maio': 5, 'Junho': 6,
                'Julho': 7, 'Agosto': 8, 'Setembro': 9,
                'Outubro': 10, 'Novembro': 11, 'Dezembro': 12
            }
            return int(year), months[month]

        # Sort rows by year and month (newest first)
        sorted_rows = sorted(rows, key=period_to_sortable_date, reverse=True)

        # Insert sorted rows into treeview
        for row in sorted_rows:
            self.treeview.insert("", END,
                values=(row[1], row[2], row[3], row[4], row[5], row[6], "✖"),
                tags=(str(row[0]),))
        
        conn.close()

    def delete_from_db(self, row_id):
        conn = sqlite3.connect('payslips.db')
        c = conn.cursor()
        c.execute('DELETE FROM payslips WHERE id = ?', (row_id,))
        conn.commit()
        conn.close()
        self.load_from_db()

    def handle_click(self, event):
        region = self.treeview.identify_region(event.x, event.y)
        if region == "cell":
            column = self.treeview.identify_column(event.x)
            if column == "#7":  # Delete column
                item = self.treeview.selection()[0]
                row_id = self.treeview.item(item, "tags")[0]
                if messagebox.askyesno("Delete", "Are you sure you want to delete this entry?"):
                    self.delete_from_db(row_id)

    def configure_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview",
            background="#333333",
            foreground="white",
            fieldbackground="#333333",
            rowheight=25
        )
        style.configure(
            "Custom.Treeview.Heading",
            background="#444444",
            foreground="white",
            font=("Arial", 10, "bold")
        )

    def calculate_kms(self, dias):
        try:
            km_value = float(self.km_multiplier.get().replace(',', '.'))
            dias_value = float(dias.replace(',', '.'))
            return f"{km_value * dias_value:.2f}".replace('.', ',')
        except ValueError:
            return "0,00"

    def import_pdfs(self):
        file_paths = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )

        if not file_paths:
            return

        for pdf_path in file_paths:
            text = extract_text_from_pdf(pdf_path)
            if not text:
                continue
                
            parsed_data = parse_pdf_data(text)
            kms = self.calculate_kms(parsed_data['Dias'])
            parsed_data['Kms'] = kms
            
            # Only save if period doesn't exist
            if self.save_to_db(parsed_data):
                messagebox.showinfo("Success", f"Data for {parsed_data['Período']} saved successfully!")
        
        # Clear temporary data
        self.data.clear()
        
        # Reload only from database
        self.load_from_db()


        def period_to_date(period):
            try:
                if period == "Not Found":
                    return datetime.min
                month, year = period.split('/')
                return datetime.strptime(f"20{year}-{month}", "%Y-%m")
            except:
                return datetime.min

        # Sort data by período (newest first)
        self.data.sort(key=lambda x: period_to_date(x['Período']), reverse=True)

        current_year = None
        for item in self.data:
            # Extract year from período (format: "Janeiro 2025")
            year = item['Período'].split()[-1]
            
            # If year changes, insert a separator
            if year != current_year:
                current_year = year
                self.treeview.insert(
                    "",
                    END,
                    values=(f"--- {year} ---", "", "", "", "", ""),
                    tags=('year_separator',)
                )
            
            # Insert regular item
            self.treeview.insert(
                "",
                END,
                values=(
                    item['Período'],
                    item['Valor Líquido'],
                    item['Subs. Refeição'],
                    item['Kms'],
                    item['Dias'],
                    item['Descontos']
                )
            )

        # Add to style configuration
        style = ttk.Style()
        style.configure(
            "Custom.Treeview",
            background="#333333",
            foreground="white",
            fieldbackground="#333333",
            rowheight=25
        )
        # Configure separator style
        self.treeview.tag_configure('year_separator', 
                                background='#444444',
                                font=('Arial', 10, 'bold'))


    def export_data(self):
        if not self.data:
            messagebox.showwarning("Warning", "No data available to export!")
            return

        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if file_path:
                df = pd.DataFrame(self.data)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Success", f"Data exported successfully to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not export data: {str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = PDFAnalyzerApp(root)
    root.mainloop()
