import pandas as pd

dataframe = pd.read_excel("Sagaløb2025_holdnavne.xlsx")
dataframe = dataframe.dropna(subset=['ID'])
dataframe = dataframe.astype(str)

dataframe['ID'] = dataframe['ID'].str.strip()
dataframe.loc[dataframe['ID'] == 'V01', 'Holdnavn'] = '6-7'


samlettekst = ""

RuteA = r'''
\begin{tabular}{|>{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|}
\hline
\cellcolor{efterårsrød}\textbf{\textcolor{white}{\rule{0pt}{3cm}Rute A}} &
\cellcolor{søblå}\textbf{\textcolor{white}{Rute B}} &
\cellcolor{korngul}\textbf{\textcolor{white}{Rute C}} &
\cellcolor{græsgrøn}\textbf{\textcolor{white}{Rute D}} \\
\hline
\end{tabular}'''

RuteB = r'''
\begin{tabular}{|>{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|}
\hline
\cellcolor{søblå}\textbf{\textcolor{white}{\rule{0pt}{3cm}Rute B}} &
\cellcolor{korngul}\textbf{\textcolor{white}{Rute C}} &
\cellcolor{græsgrøn}\textbf{\textcolor{white}{Rute D}} &
\cellcolor{efterårsrød}\textbf{\textcolor{white}{Rute A}} \\
\hline
\end{tabular}'''

RuteC = r'''
\begin{tabular}{|>{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|}
\hline
\cellcolor{korngul}\textbf{\textcolor{white}{\rule{0pt}{3cm}Rute C}} &
\cellcolor{græsgrøn}\textbf{\textcolor{white}{Rute D}} &
\cellcolor{efterårsrød}\textbf{\textcolor{white}{Rute A}} &
\cellcolor{søblå}\textbf{\textcolor{white}{Rute B}} \\
\hline
\end{tabular}'''

RuteD = r'''
\begin{tabular}{|>{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|
                >{\centering\arraybackslash}p{3cm}|}
\hline
\cellcolor{græsgrøn}\textbf{\textcolor{white}{\rule{0pt}{3cm}Rute D}} &
\cellcolor{efterårsrød}\textbf{\textcolor{white}{Rute A}} &
\cellcolor{søblå}\textbf{\textcolor{white}{Rute B}} &
\cellcolor{korngul}\textbf{\textcolor{white}{Rute C}} \\
\hline
\end{tabular}'''

for row in dataframe.itertuples():
    if row.ID[0] == "V":
        farven = "søblå"
    else:
        farven = "flammefarvet"
    if row.Rute == "A-B-C-D":
        rute_var = RuteA
    elif row.Rute == "B-C-D-A":
        rute_var = RuteB
    elif row.Rute == "C-D-A-B":
        rute_var = RuteC
    elif row.Rute == "D-A-B-C":
        rute_var = RuteD
    else:
        rute_var = ""
    instructions = [
    r'\title{Tilmeld jeres telefon til Sagaløbet}\\',
    rf'{{\fontsize{{15}}{{36}}\selectfont',
    rf'Gå ind på \textbf{{www.25hold.sagalobet.dk}} og log ind med jeres hold nummer \textbf{{{row.ID}}}.\\',
    r'I skal bruge web-appen til at sende alle løsninger på døde poster,\\',
    r'koder I finder på løbet og andet, der kan give point.\\',
    r'\textbf{\textcolor{efterårsrød}{Nødnummer: 32 67 94 99}}\\',
    r'}',
    r'\begin{center}',
    rf'{{\fontsize{{140}}{{60}}\selectfont\textbf{{Hold \textcolor{{{farven}}}{{{row.ID}}}}}\\}}',
    rf'{{\fontsize{{30}}{{50}}\selectfont\textbf{{\textcolor{{{farven}}}{{{row.Holdnavn}}}}}\\}}',
    rf'{{\fontsize{{20}}{{50}}\selectfont\textbf{{FDF {row.Kreds}}}\\}}',
    rf'{{\fontsize{{20}}{{40}}\selectfont {row._6[0]} deltagere\\}}',
    r'{\vspace{0,5cm}}',
    rute_var + '\\' + '\\',
    r'\end{center}',
    r'\newpage'
    r'\vspace{8cm}',
    r'\begin{center}',
    rf'{{\fontsize{{250}}{{60}}\selectfont\textbf{{\textcolor{{{farven}}}{{{row.ID}}}}}\\}}',
    rf'{{\fontsize{{50}}{{50}}\selectfont\textbf{{\textcolor{{{farven}}}{{{row.Holdnavn}}}}}\\}}',
    rf'{{\fontsize{{40}}{{50}}\selectfont\textbf{{FDF {row.Kreds}}}\\}}',
    r'\end{center}',
    r'\newpage'
    ]
    samlettekst += "\n".join(instructions) + "\n"
    


with open("Holdkort.tex", "w", encoding="utf-8") as f:
    f.write(samlettekst)
    f.write("\n")
    f.write(dataframe.to_latex(index=False))

