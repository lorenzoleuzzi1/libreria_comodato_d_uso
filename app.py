import tkinter as tk
from tkinter import ttk
from tkinter import messagebox 
import warnings
import pandas as pd
from PIL import ImageTk, Image
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import re

# Suppress FutureWarnings globally
warnings.simplefilter(action='ignore', category=FutureWarning)

class App(tk.Tk):

    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):

        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        
        w = 750 # width for the Tk root
        h = 550 # height for the Tk root

        # get screen width and height
        ws = self.winfo_screenwidth() # width of the screen
        hs = self.winfo_screenheight() # height of the screen

        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.title("Comodato d'uso")
        #self.configure(bg='red')
        # creating a container
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        # Set window color
        container.configure(bg='green')

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # initializing frames to an empty array
        self.frames = {}

        # iterating through a tuple consisting
        # of the different page layouts
        for F in (
            StartPage,
            AggiungiLibro,
            EliminaLibro,
            PrestaLibro,
            RestituisciLibro,
            AggiungiStudente,
            EliminaStudente,
            InviaEmail
        ):

            frame = F(container, self)
            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartPage)

    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


# first window frame startpage
class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        s = ttk.Style()
        s.configure(".", font=('Helvetica', 15))
        search_label = ttk.Label(self, text="RICERCA:")
        search_label.grid(row=1, column=1, padx=10, pady=10)
        search_entry = ttk.Entry(self, width=18)
        search_entry.grid(row=1, column=2, padx=10, pady=10)

        filtro = tk.IntVar()
        radio_nome = ttk.Checkbutton(
            self, text="Nome", variable=filtro, onvalue=1, offvalue=0,
        )
        radio_nome.grid(row=0, column=1, padx=10, pady=10)
        radio_titolo = ttk.Checkbutton(
            self, text="Titolo", variable=filtro, onvalue=2, offvalue=0
        )
        radio_titolo.grid(row=0, column=2, padx=10, pady=10)
        radio_codice = ttk.Checkbutton(
            self, text="Codice", variable=filtro, onvalue=3, offvalue=0
        )
        radio_codice.grid(row=0, column=3, padx=10, pady=10)

        da_restituire = tk.BooleanVar()
        radio_restituire = ttk.Checkbutton(
            self, text="Da resituire", variable=da_restituire, onvalue=1, offvalue=0
        )
        radio_restituire.grid(row=2, column=3, padx=10, pady=10)
        button_cerca = ttk.Button(
            self,
            text="CERCA",
            command=lambda: self.risultati(
                "cerca",
                search_entry.get(),
                filtro.get(),
                da_restituire.get(),
            ),
        )

        button_as = ttk.Button(
            self,
            text="AGGIUNGI STUDENTE",
            width=18,
            command=lambda: controller.show_frame(AggiungiStudente),
        )
        button_es = ttk.Button(
            self,
            text="ELIMINA STUDENTE",
            width=18,
            command=lambda: controller.show_frame(EliminaStudente),
        )
        button_ts = ttk.Button(
            self,
            text="TUTTI STUDENTI",
            width=18,
            command=lambda: self.risultati("tutti_studenti"),
        )
        button_al = ttk.Button(
            self,
            text="AGGIUNGI LIBRO",
            width=18,
            command=lambda: controller.show_frame(AggiungiLibro),
        )
        button_el = ttk.Button(
            self,
            text="ELIMINA LIBRO",
            width=18,
            command=lambda: controller.show_frame(EliminaLibro),
        )
        button_tl = ttk.Button(
            self,
            text="TUTTI LIBRI",
            width=18,
            command=lambda: self.risultati("tutti_libri"),
        )
        button_pl = ttk.Button(
            self,
            text="PRESTA LIBRO",
            width=18,
            command=lambda: controller.show_frame(PrestaLibro),
        )
        button_rl = ttk.Button(
            self,
            text="RESTITUISCI LIBRO",
            width=18,
            command=lambda: controller.show_frame(RestituisciLibro),
        )
        
        button_email = ttk.Button(
            self,
            text="INVIA EMAIL",
            width=18,
            command=lambda: controller.show_frame(InviaEmail),
        )
        
        button_cerca.grid(row=1, column=3, padx=20, pady=15)
        button_as.grid(row=3, column=1, padx=20, pady=15)
        button_es.grid(row=3, column=2, padx=20, pady=15)
        button_ts.grid(row=3, column=3, padx=20, pady=15)
        button_al.grid(row=4, column=1, padx=20, pady=15)
        button_el.grid(row=4, column=2, padx=20, pady=15)
        button_tl.grid(row=4, column=3, padx=20, pady=15)
        button_pl.grid(row=5, column=1, padx=20, pady=15)
        button_rl.grid(row=5, column=3, padx=20, pady=15)
        button_email.grid(row=6, column=2, padx=20, pady=15)
        
        # Create an object of tkinter ImageTk
        self.img = ImageTk.PhotoImage(Image.open("logo_elmas.png"))

        # Create a Label Widget to display the text or Image
        label_img = ttk.Label(self, image = self.img)
        label_img.grid(row=5, column=2, padx=15, pady=15)

    def risultati(
        self, valore, string=None, filtro=None, da_restituire=None
    ):
        # read from excel file
        if valore == "tutti_libri":
            df = pd.read_excel("libri.xlsx", sheet_name="Libri")
            df["ID"] = pd.to_numeric(df["ID"], errors='ignore')

        if valore == "tutti_studenti":
            df = pd.read_excel("studenti.xlsx", sheet_name="Studenti")
            # Convert 'ID' column to int (ignoring errors)
            df["ID"] = pd.to_numeric(df["ID"], errors='ignore')

            # Convert 'Telefono' column to int (ignoring errors)
            df["Telefono"] = pd.to_numeric(df["Telefono"], errors='ignore')
            df = df.sort_values(by='Nome')

        if valore == "cerca":
            df = pd.read_excel("libri.xlsx", sheet_name="Libri")
            df["ID"] = pd.to_numeric(df["ID"], errors='ignore')
            string = string.upper()
            
            if da_restituire:
                df = df[df["Data_prestito"].notnull()]
            
            if filtro == 1:
                df = df[df["Nome"].notnull()]
                df = df[df["Nome"].str.contains(string)]
            if filtro == 2:
                df = df[df["Titolo"].str.contains(string)]
            if filtro == 3:
                df = df[df["Codice"].str.contains(string)]
            
        
        popup = tk.Toplevel(self)
        popup.wm_title(f"Risultati N: {df.shape[0]}") 
        popup.geometry("1000x500")
        tree = ttk.Treeview(popup, columns=list(df.columns), show="headings")
        for column in df.columns:
            tree.heading(column, text=column)
            tree.column(column, width=100)  # Adjust the width as needed

        for index, row in df.iterrows():
            items = []
            for item in row.tolist():
                if str(item) == "nan":
                    item = "-"
                # if it's a float make it an int
                if isinstance(item, float):
                    item = str(int(item))
                items.append(item)
            tree.insert("", "end", values=items)

        # Create a style to draw borders between rows
        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=('TkDeafaultFont', 12))
        style.map("Treeview", background=[('selected', 'blue')])

        scrollbar = ttk.Scrollbar(popup, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")

        tree.pack(fill="both", expand=True)
        
class AggiungiStudente(tk.Frame):
    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)
        label = ttk.Label(self, text="AGGIUNGI STUDENTE")
        label.grid(row=1, column=0, padx=10, pady=10)
        
        cognome_label = ttk.Label(self, text="COGNOME:")
        cognome_label.grid(row=2, column=1, padx=10, pady=10)
        cognome_entry = ttk.Entry(self)
        cognome_entry.grid(row=2, column=2, padx=10, pady=10)

        nome_label = ttk.Label(self, text="NOME:")
        nome_label.grid(row=3, column=1, padx=10, pady=10)
        nome_entry = ttk.Entry(self)
        nome_entry.grid(row=3, column=2, padx=10, pady=10)
        
        telefono_label = ttk.Label(self, text="TELEFONO:")
        telefono_label.grid(row=4, column=1, padx=10, pady=10)
        telefono_entry = ttk.Entry(self)
        telefono_entry.grid(row=4, column=2, padx=10, pady=10)
        
        email_label = ttk.Label(self, text="EMAIL:")
        email_label.grid(row=5, column=1, padx=10, pady=10)
        email_entry = ttk.Entry(self)
        email_entry.grid(row=5, column=2, padx=10, pady=10)
        
        button_indietro = ttk.Button(
            self,
            text="<",
            command=lambda: controller.show_frame(StartPage),
        )
        button_indietro.grid(row=0, column=0, padx=10, pady=10)
        button_inserisci = ttk.Button(
            self,
            text="INSERISCI",
            command=lambda: self.inserisci(controller, cognome_entry.get(), nome_entry.get(), telefono_entry.get(), email_entry.get())
        )
        button_inserisci.grid(row=6, column=2, padx=10, pady=10)
        
    def inserisci(self, controller, cognome, nome, telefono, email):
        nome = nome.upper()
        cognome = cognome.upper()
        if len(nome) < 2 or len(cognome) < 2:
            messagebox.showinfo("Info", "Ricontrolla il nome/cognome")
            return
        
        df = pd.read_excel("studenti.xlsx", sheet_name="Studenti")
        # select the last row
        last_row = df.iloc[-1]
        id = int(last_row["ID"])
        if not telefono == "":
            telefono = int(telefono)
        
        if not self.is_valid_email(email) and not email == "":
            messagebox.showinfo("Info", "Ricontrolla email")
            return
        
        # make row with information 
        nome_completo = f"{cognome.replace(" ", "")} {nome.replace(" ", "")}"
        row = [str(id+1), nome_completo, str(telefono), email]
        # add row to dataframe
        with warnings.catch_warnings():
            df = pd.concat([df, pd.DataFrame([row], columns=df.columns)], ignore_index=True)
            warnings.simplefilter("ignore")
        
        #save to excel
        df.to_excel("studenti.xlsx", sheet_name="Studenti", index=False)
        messagebox.showinfo("Info", f"{row[0]} {row[1]} inserito correttamente")
        controller.show_frame(StartPage)

    def is_valid_email(self, email):
        # Regular expression pattern for a valid email address
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            
        # Use the re.match function to check if the email matches the pattern
        if re.match(pattern, email):
            return True
        else:
            return False
        
class EliminaStudente(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = ttk.Label(self, text="ELIMINA STUDENTE")
        label.grid(row=1, column=0, padx=10, pady=10)
        
        id_label = ttk.Label(self, text="ID STUDENTE:")
        id_label.grid(row=2, column=1, padx=10, pady=10)
        id_entry = ttk.Entry(self)
        id_entry.grid(row=2, column=2, padx=10, pady=10)
        
        button_indietro = ttk.Button(
            self,
            text="<",
            command=lambda: controller.show_frame(StartPage),
        )
        button_indietro.grid(row=0, column=0, padx=10, pady=10)
        button_elimina = ttk.Button(
            self,
            text="ELIMINA",
            command=lambda: self.elimina(controller, id_entry.get())
        )
        button_elimina.grid(row=3, column=2, padx=10, pady=10)
    
    def elimina(self, controller, id):
        
        df = pd.read_excel("studenti.xlsx", sheet_name="Studenti")
        #check if id in the df
        if not int(id) in df["ID"].tolist():
            messagebox.showinfo("Info", "ID non trovato")
            return
        # select the row where the id is id
        eminato = df[df["ID"] == int(id)]
        df = df[df["ID"] != int(id)]
        #delete that row
        df.to_excel("studenti.xlsx", sheet_name="Studenti", index=False)
        messagebox.showinfo("Info", f"{eminato.iloc[0]['ID']} {eminato.iloc[0]['Nome']} eliminato correttamente")
        controller.show_frame(StartPage)

class AggiungiLibro(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = ttk.Label(self, text="AGGIUNGI LIBRO")
        label.grid(row=1, column=0, padx=10, pady=10)
        
        titolo_label = ttk.Label(self, text="TITOLO:")
        titolo_label.grid(row=2, column=1, padx=10, pady=10)
        titolo = tk.StringVar()
        titolo_entry = ttk.Combobox(self, width = 27, textvariable = titolo)
        titolo_entry.grid(row=2, column=2, padx=10, pady=10)
        
        # read txt file
        with open("nomi_libri.txt", "r") as f:
            titoli = f.readlines()
        titoli.sort()
        titolo_entry['values'] = titoli
        
        codice_label = ttk.Label(self, text="CODICE:")
        codice_label.grid(row=3, column=1, padx=10, pady=10)
        codice_entry = ttk.Entry(self)
        codice_entry.grid(row=3, column=2, padx=10, pady=10)
        
        quantita_label = ttk.Label(self, text="QUANTITA':")
        quantita_label.grid(row=4, column=1, padx=10, pady=10)
        quantita_entry = ttk.Spinbox(self, from_ = 1, to = 30, width = 5)
        quantita_entry.grid(row=4, column=2, padx=10, pady=10)
        
        button_indietro = ttk.Button(
            self,
            text="<",
            command=lambda: controller.show_frame(StartPage),
        )
        button_indietro.grid(row=0, column=0, padx=10, pady=10)
        button_aggiungi = ttk.Button(
            self,
            text="AGGIUNGI",
            command=lambda: self.aggiungi(controller, titolo_entry.get(), codice_entry.get(), quantita_entry.get())
        )
        button_aggiungi.grid(row=5, column=2, padx=10, pady=10)
    
    def aggiungi(self, controller, titolo, codice, quantita):
        codice = codice.upper()
        n_progressivo = int(codice[-3:])
        quantita = int(quantita)

        if len(codice) < 5 or quantita < 1:
            messagebox.showinfo("Info", "Ricontrolla")
            return
        
        df = pd.read_excel("libri.xlsx", sheet_name="Libri")
        if str(codice) in df["Codice"].tolist():
            messagebox.showinfo("Info", "Libro già presente")
            return
        
        for i in range(quantita):
            nuovo_codice = codice[:-3] + str(n_progressivo)
            if str(nuovo_codice) in df["Codice"].tolist():
                messagebox.showinfo("Info", "Libro già presente")
                return
            # make row with information
            row = [nuovo_codice, titolo, None, None, None]
            n_progressivo += 1
            # add row to dataframe
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.concat([df, pd.DataFrame([row], columns=df.columns)], ignore_index=True)
        
        df.to_excel("libri.xlsx", sheet_name="Libri", index=False)
        messagebox.showinfo("Info", f"{quantita} libri {titolo} aggiunti correttamente")
        controller.show_frame(StartPage)

class EliminaLibro(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        
        label = ttk.Label(self, text="ELIMINA LIBRO")
        label.grid(row=1, column=0, padx=10, pady=10)
        
        id_label = ttk.Label(self, text="CODICE:")
        id_label.grid(row=2, column=2, padx=10, pady=10)
        id_entry = ttk.Entry(self)
        id_entry.grid(row=2, column=3, padx=10, pady=10)
        
        button_indietro = ttk.Button(
            self,
            text="<",
            command=lambda: controller.show_frame(StartPage),
        )
        button_indietro.grid(row=0, column=0, padx=10, pady=10)
        button_elimina = ttk.Button(
            self,
            text="ELIMINA",
            command=lambda: self.elimina(controller, id_entry.get())
        )
        button_elimina.grid(row=3, column=3, padx=10, pady=10)
    
    def elimina(self, controller, codice):
        codice = codice.upper()
        
        df = pd.read_excel("libri.xlsx", sheet_name="Libri")
        #check if id in the df
        if not str(codice) in df["Codice"].tolist():
            messagebox.showinfo("Info", "Codice non trovato")
            return
        # select the row where the id is id
        eminato = df[df["Codice"] == str(codice)]
        df = df[df["Codice"] != str(codice)]
        #delete that row
        df.to_excel("libri.xlsx", sheet_name="Libri", index=False)
        messagebox.showinfo("Info", f"{eminato.iloc[0]['Codice']} {eminato.iloc[0]['Titolo']} eliminato correttamente")
        controller.show_frame(StartPage)

class PrestaLibro(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        
        n_codici = 5

        label = ttk.Label(self, text="PRESTA")
        # print(self.studenti)
        label.grid(row=1, column=0, padx=10, pady=10)
        
        student_label = ttk.Label(self, text="STUDENTE:")
        student_label.grid(row=2, column=1, padx=10, pady=10)
        self.student_entry = ttk.Combobox(self)
        self.student_entry.bind('<KeyRelease>', self.check_input)
        self.student_entry.grid(row=2, column=2, padx=10, pady=10)
        
        codice_label = ttk.Label(self, text="CODICI:")
        codice_label.grid(row=3, column=1, padx=10, pady=10)
        codici_entries = []
        for i in range(n_codici):
            codice_entry1 = ttk.Entry(self, width = 8)
            codice_entry2 = ttk.Entry(self, width = 8)
            codice_entry1.grid(row=i+3, column=2, padx=10, pady=10)
            codice_entry2.grid(row=i+3, column=3, padx=10, pady=10)
            codici_entries.append(codice_entry1)
            codici_entries.append(codice_entry2)
        
        button_indietro = ttk.Button(
            self,
            text="<",
            command=lambda: controller.show_frame(StartPage)
        )
        button_indietro.grid(row=0, column=0, padx=10, pady=10)
        button_elimina = ttk.Button(
            self,
            text="PRESTA",
            command=lambda: self.presta(controller, [c.get() for c in codici_entries])
        )
        button_elimina.grid(row=8, column=2, padx=10, pady=10)
        
    def presta(self, controller, codici):
       
        for codice in codici:
            df = pd.read_excel("libri.xlsx", sheet_name="Libri")
            codice = codice.upper()
            
            # check if codice is not an empty string
            if codice == "":
                # go to next iteration
                continue
            
            pattern = r'^[A-Z]+\s[A-Z]+\s\([0-9]+\)$'
            if not re.match(pattern, self.student_entry.get()):
                messagebox.showinfo("Info", f"Studente non trovato.")
                return
            
            #check if id in the df
            if not str(codice) in df["Codice"].tolist():
                messagebox.showinfo("Info", f"{codice} Codice non trovato")
                continue
            
            # select the row where the codice is codice
            libro = df[df["Codice"] == str(codice)]
            
            # check if libro is already prestato
            if str(libro.iloc[0]["Data_prestito"]) != "nan":
                messagebox.showinfo("Info", f"{codice} Libro già prestato")
                continue
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                libro["Data_prestito"] = str(pd.Timestamp.today())[:-10]
                libro["Nome"] = self.student_entry.get().split("(")[0][:-1]
                libro["ID"] = str(self.student_entry.get().split("(")[1][:-1])
                # libro["ID"] = libro["ID"].astype("int", errors='ignore')
                df[df["Codice"] == str(codice)] = libro
            df.to_excel("libri.xlsx", sheet_name="Libri", index=False)
            messagebox.showinfo("Info", f"{libro.iloc[0]['Codice']} {libro.iloc[0]['Titolo']} prestato correttamente a {libro.iloc[0]['Nome']}")
        controller.show_frame(StartPage)
        
    def check_input(self, event):
        df = pd.read_excel("studenti.xlsx", sheet_name="Studenti")
        self.studenti = [f"{nome} ({id})" for nome, id in zip(df["Nome"].tolist(), df["ID"].tolist())]
        value = event.widget.get()

        if value == '':
            self.student_entry['values'] = self.studenti
        else:
            data = []
            for item in self.studenti:
                if value.lower() in item.lower():
                    data.append(item)

            self.student_entry['values'] = data


class RestituisciLibro(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        n_codici = 5
        # label of frame Layout 2
        label = ttk.Label(self, text="RESTITUISCI")

        # putting the grid in its place by using
        # grid
        label.grid(row=1, column=0, padx=10, pady=10)
        
        
        codice_label = ttk.Label(self, text="CODICI:")
        codice_label.grid(row=2, column=1, padx=10, pady=10)
        codici_entries = []
        for i in range(n_codici):
            codice_entry1 = ttk.Entry(self, width = 8)
            codice_entry2 = ttk.Entry(self, width = 8)
            codice_entry3 = ttk.Entry(self, width = 8)
            codice_entry1.grid(row=i+2, column=2, padx=10, pady=10)
            codice_entry2.grid(row=i+2, column=3, padx=10, pady=10)
            codice_entry3.grid(row=i+2, column=4, padx=10, pady=10)
            codici_entries.append(codice_entry1)
            codici_entries.append(codice_entry2)
            codici_entries.append(codice_entry3)
        
        button_indietro = ttk.Button(
            self,
            text="<",
            command=lambda: controller.show_frame(StartPage),
        )
        button_indietro.grid(row=0, column=0, padx=10, pady=10)
        button_restituisci = ttk.Button(
            self,
            text="RESTITUISCI",
            command=lambda: self.restituisci(controller, [c.get() for c in codici_entries])
        )
        button_restituisci.grid(row=8, column=3, padx=10, pady=10)
        
    def restituisci(self, controller, codici):
        for codice in codici:
            codice = codice.upper()
            
            df = pd.read_excel("libri.xlsx", sheet_name="Libri")
            libro = df[df["Codice"] == str(codice)]
            
            if codice == "":
                continue
            
            #check if id in the df
            if not str(codice) in df["Codice"].tolist():
                messagebox.showinfo("Info", f"{codice} Codice non trovato")
                continue
            
            # check if libro is already prestato
            if str(libro.iloc[0]["Data_prestito"]) == "nan":
                messagebox.showinfo("Info", f"{codice} Libro non in prestito")
                continue
            
            # select the row where the codice is codice
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                nome_restituito = libro.iloc[0]["Nome"]
                libro["Data_prestito"] = None
                libro["Nome"] = None
                libro["ID"] = None
                df[df["Codice"] == str(codice)] = libro
            df.to_excel("libri.xlsx", sheet_name="Libri", index=False)
            messagebox.showinfo("Info", f"{libro.iloc[0]['Codice']} {libro.iloc[0]['Titolo']} restituito correttamente da {nome_restituito}")
        controller.show_frame(StartPage)
        
class InviaEmail(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = ttk.Label(self, text="INVIA EMAIL")
        # print(self.studenti)
        label.grid(row=1, column=0, padx=10, pady=10)
        
        student_label = ttk.Label(self, text="STUDENTE:")
        student_label.grid(row=2, column=1, padx=10, pady=10)
        email_label = ttk.Label(self, text="EMAIL:")
        email_label.grid(row=3, column=1, padx=10, pady=10)
        email_info_label = ttk.Label(self, text="Se lasciato vuoto verrà inviata all'email registrata")
        email_info_label.grid(row=4, column=2, padx=10, pady=10)
        note_label = ttk.Label(self, text="NOTE:")
        note_label.grid(row=5, column=1, padx=10, pady=10)
        self.student_entry = ttk.Combobox(self)
        self.student_entry.bind('<KeyRelease>', self.check_input)
        self.student_entry.grid(row=2, column=2, padx=10, pady=10)
        
        email_entry = ttk.Entry(self)
        email_entry.grid(row=3, column=2, padx=10, pady=10)

        note_entry = tk.Text(self, height = 4, width = 50)
        note_entry.grid(row=5, column=2, padx=10, pady=10)

        self.warn = tk.BooleanVar()
        radio_warning = ttk.Checkbutton(
            self, text="Avviso", variable=self.warn
        )
        radio_warning.grid(row=6, column=2, padx=10, pady=10)
        
        button_indietro = ttk.Button(
            self,
            text="<",
            command=lambda: controller.show_frame(StartPage)
        )
        button_indietro.grid(row=0, column=0, padx=10, pady=10)
        button_invia = ttk.Button(
            self,
            text="INVIA",
            command=lambda: self.invia_email(controller, email_entry.get(), note_entry.get("1.0", "end-1c"))
        )
        button_invia.grid(row=9, column=2, padx=10, pady=10)
    
    def is_valid_email(self, email):
        # Regular expression pattern for a valid email address
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        
        # Use the re.match function to check if the email matches the pattern
        if re.match(pattern, email):
            return True
        else:
            return False
     
    def invia_email(self, controller, email, note):
        df = pd.read_excel("libri.xlsx", sheet_name="Libri")
        df_studenti = pd.read_excel("studenti.xlsx")
        
        pattern = r'^[A-Z]+\s[A-Z]+\s\([0-9]+\)$'
        if not re.match(pattern, self.student_entry.get()):
            messagebox.showinfo("Info", f"Studente non valido. Rincontrolla")
            return
        # df = df[df["Nome"].str.contains(self.student_entry)]
        id_studente = int(self.student_entry.get().split("(")[1][:-1])
        nome_studente = self.student_entry.get().split("(")[0][:-1]
        # take only the rows where the id is id_studente
        df = df[df["ID"] == id_studente]
        df_studenti = df_studenti[df_studenti["ID"] == id_studente]

        if email == "":
            email = df_studenti.iloc[0]["Email"]
        
        # check if email is not none
        if str(email) == "nan":
            messagebox.showinfo("Info", "Email non presente.")
            return

        if not self.is_valid_email(email):
            messagebox.showinfo("Info", f"Email non valida. Rincontrolla")
            return
        
        codici = df["Codice"].tolist()
        titoli = df["Titolo"].tolist()
        libri_body = ""
        for titolo, codice in zip(titoli, codici):
            libri_body = libri_body + f"{titolo} con codice {codice}\n"
        if self.warn.get():
            note += "\nSe i testi di cui sopra non saranno prontamente restituiti, ai sensi del codice civile, articolo 1803 e successivi, l'istituto addebiterà allo studente e alla sua famiglia (a titolo di risarcimento) una quota pari al 50% del prezzo sostenuto al momento dell'acquisto."
        else:
            note += ""
        email_body = f"Salve, \n \nL'alunno {nome_studente} ha attualmente in prestito i seguenti libri:\n" + libri_body + f"\n{note}\n" + "\nCordiali saluti, \nProfessori Marcella Marras e Aureliano Congiu\nComodato d'uso Agrario Elmas\n\nEmail generata automaticamente"
        # send email with body email_body
        # Email configuration
        sender_email = "libri.comodato@agrarioelmas.it"
        
        file = open("passkey.txt", "r")
        sender_password = file.read()
        file.close()
        # sender_password = "rbzdpuxroeyeteiv" #MaCa2023? 
        recipient_email = email
        subject = "Libri in comodato d'uso Agrario Elmas"
        
        # Create the MIMEText object
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient_email
        message['Subject'] = subject
        message.attach(MIMEText(email_body, 'plain'))
        
        # Establish a secure SMTP connection to the email server
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()  # Start TLS encryption
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())
        
        messagebox.showinfo("Info", f"Email inviata correttamente a {self.student_entry.get()}")
        controller.show_frame(StartPage)
        
    def check_input(self, event):
        df = pd.read_excel("studenti.xlsx", sheet_name="Studenti")
        self.studenti = [f"{nome} ({id})" for nome, id in zip(df["Nome"].tolist(), df["ID"].tolist())]

        value = event.widget.get()

        if value == '':
            self.student_entry['values'] = self.studenti
        else:
            data = []
            for item in self.studenti:
                if value.lower() in item.lower():
                    data.append(item)

            self.student_entry['values'] = data

# Driver Code
app = App()
app.mainloop()
