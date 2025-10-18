import pandas as pd

dataframe = pd.read_excel("/Users/hareide/Library/CloudStorage/OneDrive-Personligt/Dokumenter/Sagalob_Holdkort/Sagaløb2025_holdnavne.xlsx")
dataframe = dataframe.dropna(subset=['ID'])
dataframe = dataframe.astype(str)

dataframe['ID'] = dataframe['ID'].str.strip()
dataframe.loc[dataframe['ID'] == 'V01', 'Holdnavn'] = '6-8'

samlettekst = ""

for row in dataframe.itertuples():
    if row.ID[0] == "V":
        instructions = [
        r'\input{Preamble/VæbnerGraphics}',
        r'\vspace*{2cm}',
        r'\begin{center}',
        rf'{{\fontsize{{300}}{{60}}\selectfont\textbf{{\textcolor{{søblå}}{{{row.ID}}}}}\\}}',
        rf'{{\fontsize{{30}}{{50}}\selectfont\textbf{{\textcolor{{søblå}}{{{row.Holdnavn}}}}}\\}}',
        rf'{{\fontsize{{20}}{{50}}\selectfont\textbf{{FDF {row.Kreds}}}\\}}',
        r'\end{center}',
        r'\newpage'
        ]
    else:
        instructions = [
        r'\input{Preamble/SeniorGraphics}',
        r'\vspace*{2cm}',
        r'\begin{center}',
        rf'{{\fontsize{{300}}{{60}}\selectfont\textbf{{\textcolor{{flammefarvet}}{{{row.ID}}}}}\\}}',
        rf'{{\fontsize{{30}}{{50}}\selectfont\textbf{{\textcolor{{flammefarvet}}{{{row.Holdnavn}}}}}\\}}',
        rf'{{\fontsize{{20}}{{50}}\selectfont\textbf{{FDF {row.Kreds}}}\\}}',
        r'\end{center}',
        r'\newpage'
        ]
    samlettekst += "\n".join(instructions) + "\n"
    


with open("Sagalob_Foto_holdskilt/Holdskilt.tex", "w", encoding="utf-8") as f:
    f.write(samlettekst)
    f.write("\n")
    