import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
import math
from datetime import datetime
from openpyxl import Workbook
import pandas as pd
import os


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
            rest = self.fact3 - moy_jiro 
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
def destroy(frame):
    frame.destroy()
    print('destroyed')

def valideFistData():
    if facture_entry1.get() == '' or facture_entry2.get() == '' or facture_entry3.get() == '' or number_user_entry.get() == '':
        messagebox.showerror(message='Aucun champs de doit etre vide')
    else:
        try:
            prod = float(facture_entry1.get()) * float(facture_entry2.get()) * float(facture_entry3.get())
            editUser()
        except Exception as e:
            messagebox.showerror(message='Entrer des nombres dans les champs')
            print(e)


def editUser():
    newWindow = tk.Toplevel(window)
    newWindow.title('Ajouter les utilisateurs')
    newWindow.geometry('700x680')
    
    label_scroll = ctk.CTkLabel(newWindow, text="Remplir tous les champs s'il vous plait", font=('Arial', 20, 'bold'))
    label_scroll.pack(pady=(30, 20))
    
    primary_frame = ctk.CTkFrame(newWindow, width=680, height=520, fg_color='#EBEBEB')
    primary_frame.pack()
    
    scrollable_frame = ctk.CTkScrollableFrame(primary_frame, width=660, height=500, fg_color='#EBEBEB')
    scrollable_frame.pack(pady=10)

    fname = 'users.xlsx'
    i = 0
    k = 0
    if os.path.exists(fname):
        df = pd.read_excel(fname)
        for i in range(df.shape[0]):
            user_frame = ctk.CTkFrame(scrollable_frame,
                                      corner_radius=10, 
                                      fg_color='#DBDBDB')
            user_frame.pack(pady=10, padx=10)
            
            ctk.CTkLabel(user_frame, text=f"Prénom user {i+1}").grid(row=k, column=0)
            entry = ctk.CTkEntry(user_frame, bg_color='green')
            entry.insert(0, df.iloc[i, 0])
            entry.grid(row=k+1, column=0, padx=10, pady=10)
            entries.append(entry)

            ctk.CTkLabel(user_frame, text='Type d\'utilisateur').grid(row=k, column=1)
            combobox = ctk.CTkComboBox(user_frame, values=['Normale', 'Autre'])
            if df.iloc[i, 1] == True: combobox.set('Autre')
            else: combobox.set('Normale')
            combobox.grid(row=k+1, column=1, padx=10, pady=10)
            comboboxes.append(combobox)

            ctk.CTkLabel(user_frame, text='Absence (jour)').grid(row=k, column=2, padx=10)
            spinbox = ctk.CTkEntry(user_frame)
            spinbox.grid(row=k+1, column=2, padx=10, pady=10)
            spinbox.insert(0, 0)
            spinboxes.append(spinbox)
            
            supprButton = ctk.CTkButton(user_frame, 
                                        text='Supprimer', 
                                        fg_color='red', 
                                        text_color='white',
                                        width=30,
                                        command=lambda f=user_frame: destroy(f))
            supprButton.grid(row=k+1, column=3, padx=10)
            k += 2
            
    for i in range(i, int(number_user_entry.get())+i):
        user_frame = ctk.CTkFrame(scrollable_frame, border_color='red', corner_radius=10, fg_color='#DBDBDB')
        
        user_frame.pack(pady=10, padx=10)
        
        ctk.CTkLabel(user_frame, text=f"Prénom user {i+1}").grid(row=k, column=0)
        entry = ctk.CTkEntry(user_frame, bg_color='green')
        entry.grid(row=k+1, column=0, padx=10, pady=10)
        entries.append(entry)

        ctk.CTkLabel(user_frame, text='Type d\'utilisateur').grid(row=k, column=1)
        combobox = ctk.CTkComboBox(user_frame, values=['Normale', 'Autre'])
        combobox.grid(row=k+1, column=1, padx=10, pady=10)
        comboboxes.append(combobox)

        ctk.CTkLabel(user_frame, text='Absence (jour)').grid(row=k, column=2, padx=10)
        spinbox = ctk.CTkEntry(user_frame)
        spinbox.grid(row=k+1, column=2, padx=10, pady=10)
        spinbox.insert(0, 0)
        spinboxes.append(spinbox)
        
        supprButton = ctk.CTkButton(user_frame, 
                                    text='Supprimer', 
                                    fg_color='red', 
                                    text_color='white',
                                    width=30,
                                    command=lambda f=user_frame: destroy(f))
        supprButton.grid(row=k+1, column=3, padx=10)
        k += 2

    btn_validate = ctk.CTkButton(newWindow, text='Sauvegarder', command=validateData)
    btn_validate.pack(pady=10)

# Valider les données saisies
def validateData():
    personnes = []

    # Validation des données entrées
    w = Workbook()
    sh = w.active
    sh.append(['Prenom', 'Type'])
    fname = 'users.xlsx'
    
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

        user_type = 'Normale'
        if type_user : user_type = 'Autre'
        sh.append([name, user_type])
        missing_day = int(spinboxes[i].get())

        # Création d'un objet Personne
        p = Personne(name, missing_day, type_user)
        personnes.append(p)
    w.save(fname)
    
    # Calcul de la part de chacun
    facture = Facture(float(facture_entry1.get()), float(facture_entry2.get()), float(facture_entry3.get()), personnes)
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

    sheet.append(['', facture.get_total(), f"{math.ceil(facture.get_total() * 0.01) * 100:,}".replace(',', ' ') + " ar"])
    # Sauvegarder le fichier Excel
    workbook.save(filename)

    messagebox.showinfo(title='Facture générée', message=f"Facture générée dans le fichier : {filename}")

# Fenêtre principale Tkinter
window = ctk.CTk()
window.title('Zaraoma')
window.geometry('400x400')

# Titre
label = ctk.CTkLabel(window, text='Veuillez remplir tous les champs', font=('Arial',23))
label.pack(pady=20)

# CTK Frame
mainFrame = ctk.CTkFrame(window)
mainFrame.pack(pady=20)

# Ajout 2 derniers facture
facture_label1 = ctk.CTkLabel(mainFrame, text='Facture du 2e mois recent:')
facture_label1.grid(row=0, column=0, padx=10, pady=10)
facture_entry1 = ctk.CTkEntry(mainFrame)
facture_entry1.grid(row=0, column=1, padx=10, pady=10)

facture_label2 = ctk.CTkLabel(mainFrame, text='Facture du 1er mois recent:')
facture_label2.grid(row=1, column=0, padx=10, pady=10)
facture_entry2 = ctk.CTkEntry(mainFrame)
facture_entry2.grid(row=1, column=1)

facture_label3 = ctk.CTkLabel(mainFrame, text='Facture a payer (ce mois):')
facture_label3.grid(row=2, column=0, padx=10, pady=10)
facture_entry3 = ctk.CTkEntry(mainFrame)
facture_entry3.grid(row=2, column=1, padx=10, pady=10)

# Ajout d'utilisateur
number_user_label = ctk.CTkLabel(mainFrame, text='Nombre d\'utilisateurs (nouveau):')
number_user_label.grid(row=3, column=0, padx=10, pady=10)

number_user_entry = ctk.CTkEntry(mainFrame)
number_user_entry.grid(row=3, column=1)

btn = ctk.CTkButton(mainFrame, text='Valider', command=valideFistData)
btn.grid(row=4, column=1, pady=10)

# Exécution de l'application
window.mainloop()
