import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import math
from datetime import datetime
from openpyxl import Workbook

# Classes pour la gestion des utilisateurs et factures
class Personne:
    def __init__(self, name: str, missing_day: int, type: bool):
        self.name = name
        self.missing_day = missing_day
        self.type = type
        self.fact = 0

    def get_name(self):
        return self.name

    def get_fact(self):
        return self.fact

    def get_missing_day(self):
        return self.missing_day

class Facture:
    def __init__(self, fact1: int, fact2: int, fact3: int, personnes: list):
        self.fact1 = fact1
        self.fact2 = fact2
        self.fact3 = fact3
        self.personnes = personnes
        
    def get_total(self):
        sum = 0
        for p in self.personnes:
            sum += p.fact

        return sum
    def get_facture(self):
        max = 35000
        moy_jiro = math.ceil((self.fact1 + self.fact2 + self.fact3) / 3)  # Moyenne des 3 derniers mois
        fact_normal = math.ceil(moy_jiro / len(self.personnes))  # Part normale

        types = self.get_type()
        normal_type = types['normal']
        other_type = types['others']

        rest = self.fact3 - moy_jiro
        
        # Part pour ceux qui utilisent des équipements électriques
        if len(other_type) >= 1:
            if moy_jiro > 35000: 
                moy_jiro = max
                fact_normal = math.ceil(moy_jiro / len(self.personnes))
            fact_others = fact_normal + math.ceil(rest / len(other_type)) 
        else: 
            fact_others = fact_normal

        # Facture initiale pour tous les types
        for p in self.personnes:
            if not p.type:
                p.fact = fact_normal
            else:
                p.fact = fact_others

        differents = self.get_personne_of_day_different()

        # S'il y a des exceptions
        if len(differents) > 0:
            rest_normal = 0
            rest_others = 0
            cpt_normal = len(normal_type)
            cpt_others = len(other_type)

            fact_a_day_normal = math.ceil(fact_normal / 30)  # Facture par jour pour type normal
            fact_a_day_others = math.ceil(fact_others / 30)  # Facture par jour pour type autres

            for p in differents:
                if p.type:
                    p.fact = fact_others - (fact_a_day_others * p.missing_day)
                    rest_others += (fact_others - p.fact)
                    cpt_others -= 1
                else:
                    p.fact = fact_normal - (fact_a_day_normal * p.missing_day)
                    rest_normal += (fact_normal - p.fact)
                    cpt_normal -= 1

            for p in self.personnes:
                if p not in differents:
                    if p.type and cpt_others > 0:
                        p.fact += math.ceil(rest_others / cpt_others)
                    elif not p.type and cpt_normal > 0:
                        p.fact += math.ceil(rest_normal / cpt_normal)

    def get_type(self):
        types = {'normal': [], 'others': []}
        for p in self.personnes:
            if not p.type:
                types['normal'].append(p)
            else:
                types['others'].append(p)
        return types

    def get_personne_of_day_different(self):
        differents = []
        for p in self.personnes:
            if p.get_missing_day() > 0:
                differents.append(p)
        return differents

# Interface Tkinter
entries = []
comboboxes = []
spinboxes = []

# Créer et éditer les utilisateurs
def editUser():
    newWindow = tk.Toplevel(window)
    newWindow.title('Ajouter les utilisateurs')
    newWindow.geometry('700x700')
    
    # Créer un canvas et une scrollbar
    canvas = tk.Canvas(newWindow)
    scrollbar = ttk.Scrollbar(newWindow, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    # Configurer le canvas et la scrollbar
    canvas.configure(yscrollcommand=scrollbar.set)

    # Pack les éléments dans la fenêtre
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    k = 0
    for i in range(0, int(number_user_entry.get())):
        ttk.Label(scrollable_frame, text='Somme à payer').grid(row=k, column=3)
        topay = ttk.Label(scrollable_frame, text='non défini')
        topay.grid(row=k+1, column=3)

        ttk.Label(scrollable_frame, text=f"Prénom user (*) {k+1}").grid(row=k, column=0)
        entry = ttk.Entry(scrollable_frame)
        entry.grid(row=k+1, column=0)
        entries.append(entry)

        ttk.Label(scrollable_frame, text='Type d\'utilisateur (*)').grid(row=k, column=1)
        combobox = ttk.Combobox(scrollable_frame, values=['Normale', 'Autre'])
        combobox.grid(row=k+1, column=1)
        comboboxes.append(combobox)

        ttk.Label(scrollable_frame, text='Nombre de jours d\'absence').grid(row=k, column=2)
        spinbox = ttk.Spinbox(scrollable_frame, from_=0, to='infinity')
        spinbox.grid(row=k+1, column=2)
        spinboxes.append(spinbox)
        k += 2

    btn_validate = ttk.Button(scrollable_frame, text='Sauvegarder', command=validateData)
    btn_validate.grid(row=k, column=1)

# Valider les données saisies
def validateData():
    personnes = []

    # Validation des données entrées
    for i in range(len(entries)):
        if entries[i].get() == '':
            messagebox.showerror(title='Données invalides', message='Le champ prénom est obligatoire')
            return
        name = entries[i].get()

        if comboboxes[i].get() == '':
            messagebox.showerror(title='Données invalides', message='Choisir entre utilisateur normal ou autre')
            return
        type_user = comboboxes[i].get()
        if type_user == 'Normale' : type_user = False
        else: type_user = True

        missing_day = int(spinboxes[i].get())

        # Création d'un objet Personne
        p = Personne(name, missing_day, type_user)
        personnes.append(p)
        print(p.type)

    # Simuler les valeurs des factures
    facture = Facture(66986.98, 58352.98, 63965.08, personnes)
    facture.get_facture()

    # Générer le fichier Excel
    current_time = datetime.now().strftime("%y-%m-%d") + str(int(datetime.now().timestamp()))
    filename = f"facture-{current_time}.xlsx"

    # Utilisation d'openpyxl pour créer un fichier Excel
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Factures"

    # Ajouter les en-têtes de colonnes
    sheet.append(['Prénom', 'Valeur réelle', 'Net à payer'])

    # Ajouter les données des utilisateurs
    for p in personnes:
        sheet.append([p.get_name(), f"{p.get_fact():,}".replace(',', ' ') + " ar", f"{math.ceil(p.get_fact() * 0.01) * 100:,}".replace(',', ' ') + " ar"])

    sheet.append(['', '', facture.get_total(), f"{math.ceil(facture.get_total() * 0.01) * 100:,}".replace(',', ' ') + " ar"])
    # Sauvegarder le fichier Excel
    workbook.save(filename)

    messagebox.showinfo(title='Facture générée', message=f"Facture générée dans le fichier : {filename}")

# Fenêtre principale Tkinter
window = tk.Tk()
window.title('Zaraoma')
window.geometry('400x400')

# Ajout d'utilisateur
number_user_label = ttk.Label(window, text='Entrer le nombre d\'utilisateurs')
number_user_label.pack(pady=50)

number_user_entry = ttk.Spinbox(window, from_=1, to='infinity')
number_user_entry.pack(pady=5)

btn = ttk.Button(window, text='Valider', command=editUser)
btn.pack()

# Exécution de l'application
window.mainloop()
