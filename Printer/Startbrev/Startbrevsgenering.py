import pandas as pd

dataframe = pd.read_excel("Sagaløb2025_holdnavne.xlsx")
dataframe = dataframe.dropna(subset=['ID'])
dataframe = dataframe.astype(str)

dataframe['ID'] = dataframe['ID'].str.strip()
dataframe.loc[dataframe['ID'] == 'V01', 'Holdnavn'] = '6-7'


samlettekst = ""


for row in dataframe.itertuples():
    if row.ID[0] == "V":
        farven = "søblå"
        instructions = [
        r'\title{VELKOMMEN!}\\',
        r'\newline',
        rf'\title{{\textcolor{{{farven}}}{{{row.ID} - {row.Holdnavn} }}}}\\',
        r'\newline',
        r'Velkommen på Sagaløbet 2025!\\',
        r'Vi håber, at I har glædet jer - for det har vi!\\',
        r'Her får I de informationer, I skal bruge for at kunne starte løbet.\\',
        r'Sæt jer ned og læs brevet sammen. Det er vigtigt, at I alle sammen forstår, hvad der står i brevet, så I kommer godt igennem løbet.',
        r'\section{Hvad er der i kuverten?}',
        r'\begin{itemize}',
        r'    \item Brevet her (Startbrev)',
        r'    \item Jeres hold ID-kort',
        r'    \item Jeres første kort til jeres første rute',
        r'    \item Signaturforklaring til alle kort',
        r'    \item Sagaløbets regler',
        r'    \item Jeres første venteopgaver til til jeres første rute',
        r'\end{itemize}',
        r'\section{Løbets opbygning}',
        r'Sagaløbet 2025 er opbygget af fire mindre ruter, som I skal gennemføre. Alle ruter starter og slutter her ved basen. I skal gennemføre en hel rute, før I kan starte det næste. Alle hold skal gennemføre alle fire ruter i en bestemt rækkefølge. I kan se rækkefølgen på jeres hold ID-kort.\\'
        r'\newline',
        r'I bestemmer selv, om I vil gennemføre alle fire ruter, men der er kun mærker til dem, som har gennemført alle fire ruter! Når I skal starte en ny rute, er det \textbf{vigtigt}, at I snakker sammen først og finder ud af, om hele holdet kan gennemføre ruten. Er der nogen, som ikke tror, at de kan gennemføre den, er det bedre at give op ved basen end at give op midt på ruten. I får nemlig ekstra 100 point for hver rute, I gennemfører, hvor ingen af jer giver op undervejs.\\',        
        r'\section{Kortet}',
        f'Det kort, I har fået, viser posterne på \\textbf{{RUTE {row.Rute[0]}}}. Det er den første rute, I skal gennemføre. På ruten er de levende poster markeret med tal, og de døde poster er markeret med et \\textbf{{X}}. Alle posterne starter med bogstavet \\textbf{{{row.Rute[0]}}}. Husk at posterne SKAL gennemføres i den rigtige rækkefølge. Man kan altså ikke tage post A1 og så post A4. Ønsker I at springe en post over, skal I stadig tjekke ind og tjekke ud ved posten. Springer I en post over, kan I IKKE gå tilbage og gennemføre den senere. Bemærk at der kan være hemmelige poster på ruten, disse er nødvendigvis ikke markeret på kortet.',
        r'\section{App}',
        r'På \textbf{www.hold25.sagalobet.dk} skal I undervejs på løbet indtaste svarene på jeres døde poster, venteopgaver og andre spændende ting, I måske finder. Det er også på den side, I skal give posterne point til postpokalen.\\',
        f'For at logge ind skal I bruge jeres hold ID - altså \\textbf{{{row.ID}}}.',
        r'\newpage',
        r'\section{De levende poster}',
        r'Når I kommer til en levende post på jeres rute, skal I:',
        r'\begin{itemize}',
        r'    \item Tjekke ind ved postmandskabet og aflevere jeres ID-kort',
        r'    \item Læse postbeskrivelsen til posten',
        r'    \item Vælge om I vil forsøge at få de 25 bonuspoint på posten',
        r'    \item Løse posten',
        r'    \item Give posten mellem 1 og 10 point på \textbf{www.hold25.sagalobet.dk} - Bedøm posten ud fra kriterierne:',
        r'    \begin{itemize}',
        r'        \item Var posten sjov?',
        r'        \item Var posten udfordrende?',
        r'        \item Oplevede I noget nyt?',
        r'        \item Var posten gennemført?',
        r'        \end{itemize}',
        r'    \item Tjekke ud ved postmandskabet og få jeres ID-kort tilbage',
        r'\end{itemize}',
        r'\section{Når løbet begynder}',
        r'Når klokken bliver 16:00 begynder Sagaløbet 2025. Vi begynder med en fælles startpost, hvor I som hold kan vinde op til 100 point - Lige efter legen, åbner alle posterne, og så er det ellers med at komme af sted.',
        r'\section{Skolegårdsposter}',    
        r'I og omkring målområdet kan I finde 4 døde poster, som er nummereret G1-G4.\\',
        r'De første 3 poster skal hver især løses ved at sende en besvarelse ind på appen.\\',    
        r'I skal sørge for at indsende koden \textbf{samlet}, og ikke enkeltvise bogstaver - indsendes de enkeltvis, gives der ikke point.\\',
        r'\newline',
        r'Indsend jeres svar på \textbf{www.hold25.sagalobet.dk}\\',
        r'\newline',
        r'Den sidste post - G4 - er et QR mysterie, som løses ved at skanne den korrekte QR kode.\\',
        r'Her skal I være opmærksomme på, at alle QR koder kan scannes. Hvis I scanner den forkerte QR kode får I -50 point, men scannes den rigtige får I 200 point!\\',
        r'Omkring post G4 vil der være ledetråde til at hjælpe jer med at finde den rigtige QR kode.',
        r'\section{Hvis I bliver pressede}',
        r'Undervejs på løbet, bliver I helt sikkert trætte og eller pressede. Hvis I bliver så trætte, at I slet ikke kan overskue at gennemføre ruten, kan I altid give op på en levende post. Her ringer postmandskabet til basen og så bliver I hentet og kørt tilbage til basen. \textbf{I må IKKE ringe hjem og blive hentet}. I kan først komme helt hjem når Sagaløbet er færdig søndag morgen. Hvis en eller flere af jer vælger at give op, kan resten af holdet sagtens fortsætte. Man kan ikke komme tilbage på løbet, når først man har givet op.\\',
        r'\newline',
        r'Til sidst vil vi bare ønske jer rigtig god tur!\\',
        r'Husk at passe på hinanden, få så mange point, som I kan og hold øjne og ører åbne.\\',
        r'\newline',
        r'Rigtig godt løb!\\',
        r'\newline',
        r'\textcolor{søblå}{De bedste hilsner}\\',
        r'\textcolor{natblå}{\textbf{Alberte, Augusta, Charlotte, Dan, Jens Peter, Jonas, Lasse, Lucas, Magnus \& Thor}}\\',
        r'\textcolor{natblå}{\textbf{Sagaløbsudvalget}}\\',
        r'\newpage',
        ]
    
    else:
        farven = "flammefarvet"
        ekstra = ' \newline  \newline',

        instructions = [
        r'\title{VELKOMMEN!}\\',
        r'\newline',
        rf'\title{{\textcolor{{{farven}}}{{{row.ID} - {row.Holdnavn} }}}}\\',
        r'\newline',
        r'Velkommen på Sagaløbet 2025!\\',
        r'Vi håber, at I har glædet jer - for det har vi!\\',
        r'Her får I de informationer, I skal bruge for at kunne starte løbet.\\',
        r'Sæt jer ned og læs brevet sammen. Det er vigtigt, at I alle sammen forstår, hvad der står i brevet, så I kommer godt igennem løbet.',
        r'\section{Hvad er der i kuverten?}',
        r'\begin{itemize}',
        r'    \item Brevet her (Startbrev)',
        r'    \item Jeres hold ID-kort',
        r'    \item Jeres første kort til jeres første rute',
        r'    \item Signaturforklaring til alle kort',
        r'    \item Sagaløbets regler',
        r'    \item Jeres første venteopgaver til til jeres første rute',
        r'\end{itemize}',
        r'\section{Løbets opbygning}',
        r'Sagaløbet 2025 er opbygget af fire mindre ruter, som I skal gennemføre. Alle ruter starter og slutter her ved basen. I skal gennemføre en hel rute, før I kan starte det næste. Alle hold skal gennemføre alle fire ruter i en bestemt rækkefølge. I kan se rækkefølgen på jeres hold ID-kort.\\'
        r'\newline',
        r'I bestemmer selv, om I vil gennemføre alle fire ruter, men der er kun mærker til dem, som har gennemført alle fire ruter! Når I skal starte en ny rute, er det \textbf{vigtigt}, at I snakker sammen først og finder ud af, om hele holdet kan gennemføre ruten. Er der nogen, som ikke tror, at de kan gennemføre den, er det bedre at give op ved basen end at give op midt på ruten. I får nemlig ekstra 100 point for hver rute, I gennemfører, hvor ingen af jer giver op undervejs.\\',
        r'\newline',
        r'Undervejs er der seniorrelateret indhold, som \textbf{alle seniorhold skal} besøge. Hvis ikke seniorholdene besøger seniorposterne får holdet minus 100 point pr. post som ikke besøges. Posterne er levende poster, som har bestemte åbnings- og lukketider. Seniorposterne er markeret med orange \textbf{S} på kortene.\\',
        r'\newline',
        r'\section{Kortet}',
        f'Det kort, I har fået, viser posterne på \\textbf{{RUTE {row.Rute[0]}}}. Det er den første rute, I skal gennemføre. På ruten er de levende poster markeret med tal, og de døde poster er markeret med et \\textbf{{X}}. Alle posterne starter med bogstavet \\textbf{{{row.Rute[0]}}}. Husk at posterne SKAL gennemføres i den rigtige rækkefølge. Man kan altså ikke tage post A1 og så post A4. Ønsker I at springe en post over, skal I stadig tjekke ind og tjekke ud ved posten. Springer I en post over, kan I IKKE gå tilbage og gennemføre den senere. Bemærk at der kan være hemmelige poster på ruten, disse er nødvendigvis ikke markeret på kortet.',
        r'\section{App}',
        r'På \textbf{www.hold25.sagalobet.dk} skal I undervejs på løbet indtaste svarene på jeres døde poster, venteopgaver og andre spændende ting, I måske finder. Det er også på den side, I skal give posterne point til postpokalen.\\',
        f'For at logge ind skal I bruge jeres hold ID - altså \\textbf{{{row.ID}}}.',
        r'\newpage',
        r'\section{De levende poster}',
        r'Når I kommer til en levende post på jeres rute, skal I:',
        r'\begin{itemize}',
        r'    \item Tjekke ind ved postmandskabet og aflevere jeres ID-kort',
        r'    \item Læse postbeskrivelsen til posten',
        r'    \item Vælge om I vil forsøge at få de 25 bonuspoint på posten',
        r'    \item Løse posten',
        r'    \item Give posten mellem 1 og 10 point på \textbf{www.hold25.sagalobet.dk} - Bedøm posten ud fra kriterierne:',
        r'    \begin{itemize}',
        r'        \item Var posten sjov?',
        r'        \item Var posten udfordrende?',
        r'        \item Oplevede I noget nyt?',
        r'        \item Var posten gennemført?',
        r'        \end{itemize}',
        r'    \item Tjekke ud ved postmandskabet og få jeres ID-kort tilbage',
        r'\end{itemize}',
        r'\section{Når løbet begynder}',
        r'Når klokken bliver 16:00 begynder Sagaløbet 2025. Vi begynder med en fælles startpost, hvor I som hold kan vinde op til 100 point - Lige efter legen, åbner alle posterne, og så er det ellers med at komme af sted.',
        r'\section{Skolegårdsposter}',    
        r'I og omkring målområdet kan I finde 4 døde poster, som er nummereret G1-G4.\\',
        r'De første 3 poster skal hver især løses ved at sende en besvarelse ind på appen.\\',    
        r'I skal sørge for at indsende koden \textbf{samlet}, og ikke enkeltvise bogstaver - indsendes de enkeltvis, gives der ikke point.\\',
        r'\newline',
        r'Indsend jeres svar på \textbf{www.hold25.sagalobet.dk}\\',
        r'\newline',
        r'Den sidste post - G4 - er et QR mysterie, som løses ved at skanne den korrekte QR kode.\\',
        r'Her skal I være opmærksomme på, at alle QR koder kan scannes. Hvis I scanner den forkerte QR kode får I -50 point, men scannes den rigtige får I 200 point!\\',
        r'Omkring post G4 vil der være ledetråde til at hjælpe jer med at finde den rigtige QR kode.',
        r'\section{Hvis I bliver pressede}',
        r'Undervejs på løbet, bliver I helt sikkert trætte og eller pressede. Hvis I bliver så trætte, at I slet ikke kan overskue at gennemføre ruten, kan I altid give op på en levende post. Her ringer postmandskabet til basen og så bliver I hentet og kørt tilbage til basen. \textbf{I må IKKE ringe hjem og blive hentet}. I kan først komme helt hjem når Sagaløbet er færdig søndag morgen. Hvis en eller flere af jer vælger at give op, kan resten af holdet sagtens fortsætte. Man kan ikke komme tilbage på løbet, når først man har givet op.\\',
        r'\newline',
        r'Til sidst vil vi bare ønske jer rigtig god tur!\\',
        r'Husk at passe på hinanden, få så mange point, som I kan og hold øjne og ører åbne.\\',
        r'\newline',
        r'Rigtig godt løb!\\',
        r'\newline',
        r'\textcolor{søblå}{De bedste hilsner}\\',
        r'\textcolor{natblå}{\textbf{Alberte, Augusta, Charlotte, Dan, Jens Peter, Jonas, Lasse, Lucas, Magnus \& Thor}}\\',
        r'\textcolor{natblå}{\textbf{Sagaløbsudvalget}}\\',
        r'\newpage',
        ]

    
    samlettekst += "\n".join(instructions) + "\n"
    


with open("Startbrev.tex", "w", encoding="utf-8") as f:
    f.write(samlettekst)
    f.write("\n")
    