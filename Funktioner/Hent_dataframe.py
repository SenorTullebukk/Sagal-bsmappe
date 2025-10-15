from Funktioner import *

def import_deltagerliste(title):
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    if file_path:
        df = pd.read_excel(file_path)
        return df
    else:
        return None

def import_postoversigt(title):
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    if file_path:
        df = pd.read_excel(file_path)
        return df
    else:
        return None