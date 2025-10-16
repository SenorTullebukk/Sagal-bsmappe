from Funktioner import *

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    deltagerliste = import_deltagerliste("Hent deltagerliste")
    postoversigt = import_postoversigt("Hent postoversigt")
    if deltagerliste is not None or postoversigt is not None:
        lav_pdf(deltagerliste, postoversigt)
    else:
        print("Fejl i excel import")
        
if __name__ == "__main__":
    main()

'''
Todo:
-Lave preamble for hver type side
-Lave en funktion som laver deltagerlisten om til en dataframe
-Laver Postlisten om til en dataframe
-Lave mapper til hver af de ting som skal laves pdf'er af
-Lave en funktion som samler pdf'erne til holdbrev og osv.
-Så er der også lidt øvrige som skal laves.'''