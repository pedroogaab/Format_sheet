import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

root = tk.Tk()
root.title("File Processor")
root.geometry("400x300")

training_file = None

def select_training_file():
    global training_file
    training_file = filedialog.askopenfilename(title="Select Training File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if training_file:
        training_label.config(text=f"Selected: {training_file.split('/')[-1]}")


def process_files():
    global training_file

    if not training_file:
        messagebox.showerror("Error", "Select the training file!")
        return

    try:
        df_training = pd.read_excel(training_file)
        df_email = pd.read_csv('view_base_funcionarios.csv', encoding='windows-1252', delimiter=';')

        df_training['Matricula_temp'] = df_training['Matricula'].astype(str).str.zfill(8)
        df_email['Matrícula do Funcionário'] = df_email['Matrícula do Funcionário'].astype(str).str.zfill(8)

        df_merged = pd.merge(df_training, df_email[['Matrícula do Funcionário', 'Email Gestor', 'E-mail do Funcionário']],
                             left_on='Matricula_temp', right_on='Matrícula do Funcionário', how='left')
        df_merged.drop(columns=['Matricula_temp', 'Matrícula do Funcionário'], inplace=True)

        # Format date columns
        date_columns = ['Data Início', 'Data Fim']
        for col in date_columns:
            if col in df_merged.columns:
                df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce').dt.strftime('%d/%m/%Y')

        output_filename = filename_entry.get().strip()
        if not output_filename:
            messagebox.showerror("Error", "O nome informado não é valido.")
            return

        with pd.ExcelWriter(f"{output_filename}.xlsx", engine='openpyxl') as writer:
            df_merged.to_excel(writer, index=False)
            worksheet = writer.sheets[writer.book.sheetnames[0]]

            for col_idx, column_cells in enumerate(worksheet.columns, start=1):
                max_length = 0
                col_letter = get_column_letter(col_idx)
                for cell in column_cells:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = max_length + 6
                worksheet.column_dimensions[col_letter].width = adjusted_width

            table = Table(displayName="Tabela1", ref=worksheet.dimensions)
            style = TableStyleInfo(name="TableStyleLight14", showFirstColumn=False, showLastColumn=False,
                                   showRowStripes=True, showColumnStripes=True)
            table.tableStyleInfo = style
            worksheet.add_table(table)

        messagebox.showinfo("Success", f"File saved as {output_filename}.xlsx!")
        sys.exit()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# UI Elements
btn_training = tk.Button(root, text="Select Training File", command=select_training_file)
btn_training.pack(pady=5)
training_label = tk.Label(root, text="No file selected")
training_label.pack(pady=5)

filename_entry = tk.Entry(root, width=30)
filename_entry.pack(pady=10)
filename_entry.insert(0,"")

btn_process = tk.Button(root, text="Process Files", command=process_files)
btn_process.pack(pady=10)

root.mainloop()
