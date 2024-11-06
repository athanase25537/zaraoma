import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
import math
from datetime import datetime
from openpyxl import Workbook
import pandas as pd
import os
import time
from PIL import Image, ImageTk
import cairosvg

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
            fact_normal = self.fact3/len(self.personnes)
            fact_others = fact_normal

        # Facture initiale pour tous les types
        for p in self.personnes:
            if not p.type:
                p.fact = fact_normal
            else:
                p.fact = fact_others

        differents = self.get_personne_of_day_different()

        # S'il y a des exceptions
        rest_normal = 0
        rest_others = 0
        cpt_normal = len(normal_type)
        cpt_others = len(other_type)

        fact_a_day_normal = math.ceil(fact_normal / 30)  # Facture par jour pour type normal
        fact_a_day_others = math.ceil(fact_others / 30)  # Facture par jour pour type autres

        if len(differents) > 1:
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
                        
        elif len(differents) > 0:
            for p in differents:
                if p.type:
                    p.fact = fact_others - (fact_a_day_others * p.missing_day)
                    rest = self.fact3 - (fact_normal * (len(self.personnes) - 1)) - p.fact
                else:
                    p.fact = fact_normal - (fact_a_day_normal * p.missing_day)
                    rest = fact_normal - p.fact
                
                
            for p in self.personnes:
                if p not in differents:
                    p.fact += rest / (len(self.personnes) - 1)

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
errors_name = []
errors_type = []
errors_miss = []
firstname = []
errors_main = []

def add_user(entry, scrollable_frame, btn_validate, k):
    if entry.get() == '' or not entry.get().isdigit(): entry.configure(border_color='#D9534F')
    else:
        for i in range(int(entry.get())):
            user_frame = ctk.CTkFrame(scrollable_frame, 
                                  corner_radius=10,
                                  fg_color='#DBDBDB',)
        
            user_frame.pack(pady=10, padx=10)
            
            ctk.CTkLabel(user_frame, text=f"Prénom user {i+1}").grid(row=k, column=0)
            entry = ctk.CTkEntry(user_frame, 
                                border_color='#F27438',)
            
            entry.grid(row=k+1, column=0, padx=10, pady=10)
            entries.append(entry)

            ctk.CTkLabel(user_frame, text='Type d\'utilisateur').grid(row=k, column=1)
            combobox = ctk.CTkComboBox(user_frame, 
                                    values=['Normale', 'Autre'],
                                    border_color='#F27438',
                                    button_color='#F27438')
            combobox.grid(row=k+1, column=1, padx=10, pady=10)
            
            comboboxes.append(combobox)

            ctk.CTkLabel(user_frame, text='Absence (jour)').grid(row=k, column=2, padx=10)
            spinbox = ctk.CTkEntry(user_frame,
                                border_color='#F27438',)
            spinbox.grid(row=k+1, column=2, padx=10, pady=10)
            spinbox.insert(0, 0)
            spinboxes.append(spinbox)
            
            supprButton = ctk.CTkButton(user_frame, 
                                        text='Supprimer', 
                                        text_color='white',
                                        fg_color='#D9534F',
                                        hover_color='#C9302C',
                                        width=30,
                                        command=lambda f=user_frame: destroy(scrollable_frame, f, btn_validate))
            supprButton.grid(row=k+1, column=3, padx=10)
            
            error_label_name = ctk.CTkLabel(user_frame, 
                                        text='',
                                        text_color='#D9534F')
            error_label_name.grid(row=k+2, column=0)
            errors_name.append(error_label_name)
            
            
            error_label_type = ctk.CTkLabel(user_frame, 
                                        text='',
                                        text_color='#D9534F')
            error_label_type.grid(row=k+2, column=1)
            errors_type.append(error_label_type)
            
            error_label_miss = ctk.CTkLabel(user_frame, 
                                        text='',
                                        text_color='#D9534F')
            error_label_miss.grid(row=k+2, column=2)
            errors_miss.append(error_label_miss)
            k += 3

def check_valider_activation():
    entree = [facture_entry1, facture_entry2, facture_entry3, number_user_entry]
    i = 0
    there_are_error = []
    for entry in entree:
        parent = window.nametowidget(entry.winfo_parent())
        # Create error container
        errCont = parent.nametowidget(errors_main[i])
        
        if entry.get() == '':
            btn.configure(state="disabled")
            errCont.configure(text='Ce champ ne doit pas être vide')
            there_are_error.append(False)
        elif not entry.get().isdigit():
            btn.configure(state="disabled")
            entry.configure(border_color='#D9534F')
            errCont.configure(text='Entrer un nombre')
            there_are_error.append(False)
        else: 
            entry.configure(border_color='#F27438')
            errCont.configure(text='')
            there_are_error.append(True)
        i += 1
    if not False in there_are_error: btn.configure(state="normal")

def on_entry_change(event=None):
    check_valider_activation()

# Fonction pour vérifier si le cadre est vide
def is_frame_empty(frame):
    return len(frame.winfo_children()) == 0

# Fonction pour écouter si le cadre est vide
def check_if_frame_empty(frame):
    if is_frame_empty(frame): return True
    else: return False

    # Vérifier à nouveau après un certain délai

def hide():
    for i in range(101):
        progressBar.set(i / 100)  # Met à jour la barre de progression
        splash_percentage
        splash_percentage.configure(text='Loading... '+str(int(i/100*100))+'%')
        splash_screen.update_idletasks()  # Met à jour l'interface pour afficher la progression
        time.sleep(0.03)
    splash_screen.destroy()
    
# Créer et éditer les utilisateurs
def destroyFrame(frame):
    entries.clear()
    comboboxes.clear()
    spinboxes.clear()
    errors_name.clear()
    errors_type.clear()
    errors_miss.clear()
    firstname.clear()
    errors_main.clear()
    frame.destroy()
    
    
def destroy(containerFrame, frame, btn_validate):
    frame.destroy()
    emptyFrame = check_if_frame_empty(containerFrame)
    if emptyFrame: btn_validate.configure(state='disabled')

def valideFistData():
    if facture_entry1.get() == '' or facture_entry2.get() == '' or facture_entry3.get() == '' or number_user_entry.get() == '':
        messagebox.showerror(message='Aucun champs de doit etre vide')
    else:
        try:
            prod = float(facture_entry1.get()) * float(facture_entry2.get()) * float(facture_entry3.get())
            editUser()
        except Exception as e:
            messagebox.showerror(message='Entrer des nombres dans les champs')

def editUser():
    newWindow = tk.Toplevel(window)
    newWindow.resizable(False, False)
    newWindow.title('Ajouter les utilisateurs')
    newWindow.geometry('700x680')
    
    label_scroll = ctk.CTkLabel(newWindow, text="Remplir tous les champs s'il vous plait", font=('Arial', 20, 'bold'))
    label_scroll.pack(pady=10)
    
    primary_frame = ctk.CTkFrame(newWindow, width=680, height=520, fg_color='#424340')
    primary_frame.pack()
    
    scrollable_frame = ctk.CTkScrollableFrame(primary_frame, 
                                              width=600, 
                                              height=450, 
                                              fg_color='#EBEBEB')
    scrollable_frame.pack(pady=10, padx=10)

    dirname = 'utilisateurs'
    fname = 'users.xlsx'
    i = 0
    k = 0
    if os.path.exists(os.path.join(dirname, fname)):
        df = pd.read_excel(os.path.join(dirname, fname))
        for i in range(df.shape[0]):
            user_frame = ctk.CTkFrame(scrollable_frame,
                                      corner_radius=10, 
                                      fg_color='#DBDBDB')
            user_frame.pack(pady=10, padx=5)
            
            ctk.CTkLabel(user_frame, text=f"Prénom user {i+1}").grid(row=k, column=0)
            entry = ctk.CTkEntry(user_frame,
                                 border_color='#F27438')
            entry.insert(0, df.iloc[i, 0].capitalize())
            firstname.append(df.iloc[i, 0])
            entry.grid(row=k+1, column=0, padx=10, pady=10)
            entries.append(entry)

            ctk.CTkLabel(user_frame, text='Type d\'utilisateur').grid(row=k, column=1)
            combobox = ctk.CTkComboBox(user_frame, values=['Normale', 'Autre'],
                                       border_color='#F27438',
                                       button_color='#F27438')
            if df.iloc[i, 1] == True: combobox.set('Autre')
            else: combobox.set('Normale')
            combobox.grid(row=k+1, column=1, padx=10, pady=10)
            comboboxes.append(combobox)

            ctk.CTkLabel(user_frame, text='Absence (jour)').grid(row=k, column=2, padx=10)
            spinbox = ctk.CTkEntry(user_frame,
                                   border_color='#F27438',)
            spinbox.grid(row=k+1, column=2, padx=10, pady=10)
            spinbox.insert(0, 0)
            spinboxes.append(spinbox)
            
            supprButton = ctk.CTkButton(user_frame, 
                                        text='Supprimer',
                                        text_color='white',
                                        fg_color='#D9534F',
                                        hover_color='#C9302C',
                                        width=30,
                                        command=lambda f=user_frame: destroy(scrollable_frame, f, btn_validate))
            supprButton.grid(row=k+1, column=3, padx=10)
            
            error_label_name = ctk.CTkLabel(user_frame, 
                                      text='',
                                      text_color='#D9534F')
            error_label_name.grid(row=k+2, column=0)
            errors_name.append(error_label_name)
            
            
            error_label_type = ctk.CTkLabel(user_frame, 
                                        text='',
                                        text_color='#D9534F')
            error_label_type.grid(row=k+2, column=1)
            errors_type.append(error_label_type)
            
            error_label_miss = ctk.CTkLabel(user_frame, 
                                        text='',
                                        text_color='#D9534F')
            error_label_miss.grid(row=k+2, column=2)
            errors_miss.append(error_label_miss)
            k += 3
            
    for i in range(i, int(number_user_entry.get())+i):
        user_frame = ctk.CTkFrame(scrollable_frame, 
                                  corner_radius=10,
                                  fg_color='#DBDBDB',)
        
        user_frame.pack(pady=10, padx=5)
        
        ctk.CTkLabel(user_frame, text=f"Prénom user {i+1}").grid(row=k, column=0)
        entry = ctk.CTkEntry(user_frame, 
                             border_color='#F27438',)
        
        entry.grid(row=k+1, column=0, padx=10, pady=10)
        entries.append(entry)

        ctk.CTkLabel(user_frame, text='Type d\'utilisateur').grid(row=k, column=1)
        combobox = ctk.CTkComboBox(user_frame, 
                                   values=['Normale', 'Autre'],
                                   border_color='#F27438',
                                   button_color='#F27438')
        combobox.grid(row=k+1, column=1, padx=10, pady=10)
        
        comboboxes.append(combobox)

        ctk.CTkLabel(user_frame, text='Absence (jour)').grid(row=k, column=2, padx=10)
        spinbox = ctk.CTkEntry(user_frame,
                               border_color='#F27438',)
        spinbox.grid(row=k+1, column=2, padx=10, pady=10)
        spinbox.insert(0, 0)
        spinboxes.append(spinbox)
        
        supprButton = ctk.CTkButton(user_frame, 
                                    text='Supprimer', 
                                    text_color='white',
                                    fg_color='#D9534F',
                                    hover_color='#C9302C',
                                    width=30,
                                    command=lambda f=user_frame: destroy(scrollable_frame, f, btn_validate))
        supprButton.grid(row=k+1, column=3, padx=10)
        
        error_label_name = ctk.CTkLabel(user_frame, 
                                      text='',
                                      text_color='#D9534F')
        error_label_name.grid(row=k+2, column=0)
        errors_name.append(error_label_name)
        
        
        error_label_type = ctk.CTkLabel(user_frame, 
                                      text='',
                                      text_color='#D9534F')
        error_label_type.grid(row=k+2, column=1)
        errors_type.append(error_label_type)
        
        error_label_miss = ctk.CTkLabel(user_frame, 
                                      text='',
                                      text_color='#D9534F')
        error_label_miss.grid(row=k+2, column=2)
        errors_miss.append(error_label_miss)
        k += 3
    
    
    # Add button
    user_frame = ctk.CTkFrame(primary_frame, 
                              corner_radius=10,
                              fg_color='#EBEBEB')
        
    user_frame.pack(pady=(0, 10))
    
    add_entry_frame = ctk.CTkFrame(user_frame,
                                   corner_radius=10,
                                   fg_color='#EBEBEB')
    add_entry_frame.pack(side=ctk.LEFT, pady=30, padx=(0,136))
    add_label = ctk.CTkLabel(add_entry_frame,
                             text='Entrer nommbre d\'utilisateur',)
    add_label.pack(side=ctk.LEFT, padx=(10, 12))
    add_entry = ctk.CTkEntry(add_entry_frame)
    add_entry.pack(side=ctk.RIGHT)
    
    btn_add = ctk.CTkButton(user_frame,
                            text='Ajouter',
                            command=lambda: add_user(add_entry, scrollable_frame, btn_validate, k))
    btn_add.pack(side=ctk.RIGHT, pady=30, padx=10)
    
    # Save and back button
    save_back_frame = ctk.CTkFrame(newWindow, width=680)
    save_back_frame.pack()
    btn_validate = ctk.CTkButton(save_back_frame, 
                                text='Sauvegarder', 
                                text_color='white',
                                fg_color='#F27438',
                                hover_color='#D86228',
                                state='normal',
                                command=lambda: validateData(newWindow))


    btn_validate.pack(side=ctk.RIGHT,
                      pady=10,
                      padx=68)
    btn_back = ctk.CTkButton(save_back_frame,
                             text='Retour',
                             width=80,
                             command=lambda: destroyFrame(newWindow))
    btn_back.pack(side=ctk.RIGHT)

# Valider les données saisies
def validateData(newWindow):
    personnes = []
    # Validation des données entrées
    w = Workbook()
    sh = w.active
    sh.append(['Prenom', 'Type'])
    dirname = 'utilisateurs'
    fname = 'users.xlsx'
    if not os.path.exists(os.path.join(dirname, fname)): os.makedirs(dirname, exist_ok=True)
    
    errorData = []
    for i in range(0, len(entries)):
        try:
            parent = newWindow.nametowidget(entries[i].winfo_parent())
            errorContainer = parent.nametowidget(errors_name[i])
            if entries[i].get() == '':
                errorContainer.configure(text='Le prénom ne doit pas être vide')
                errorData.append(False)
                # return
            elif i >= len(firstname) and (entries[i].get().lower() in firstname):
                errorContainer.configure(text=entries[i].get()+' existe déjà')
                errorData.append(False)
            else: 
                errorContainer.configure(text='')
                errorData.append(True)
            if not False in errorData: name = entries[i].get().lower()
        except:
            continue

        try:
            errorContainer = parent.nametowidget(errors_type[i])
            if comboboxes[i].get() != 'Normale' and comboboxes[i].get() != 'Autre':
                errorContainer.configure(text='Le type doit être "Normal" ou "Autre"')
                errorData.append(False)
            else: 
                errorContainer.configure(text='')
                errorData.append(True)
            if not False in errorData: type_user = comboboxes[i].get()
        except:
            continue
        try:
            errorContainer = parent.nametowidget(errors_miss[i])
            if spinboxes[i].get() == '' or int(spinboxes[i].get()) < 0:
                errorContainer.configure(text='Absence ne peut être vide ou négatif')
                errorData.append(False)
            else: 
                errorContainer.configure(text='')
                errorData.append(True)
            if not False in errorData: missing_day = int(spinboxes[i].get())
        except:
            continue
        
        if type_user == 'Normale' : type_user = False
        else: type_user = True

        user_type = 'Normale'
        if type_user : user_type = 'Autre'
        sh.append([name, user_type])

        # Création d'un objet Personne
        p = Personne(name, missing_day, type_user)
        personnes.append(p)
        
    if False in errorData: return
    w.save(os.path.join(dirname, fname))
    
    # Calcul de la part de chacun
    facture = Facture(float(facture_entry1.get()), float(facture_entry2.get()), float(facture_entry3.get()), personnes)
    facture.get_facture()

    # Générer le fichier Excel
    current_time = datetime.now().strftime("%y-%m-%d") + str(int(datetime.now().timestamp()))
    filename = f"facture-{current_time}.xlsx"
    folder_name = 'factures'
    os.makedirs(folder_name, exist_ok=True)
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
    workbook.save(os.path.join(folder_name, filename))

    messagebox.showinfo(title='Facture générée', message=f"Facture générée dans le fichier : {filename}")

# Fonction pour convertir un fichier SVG en image compatible
def svg_to_png(svg_path, output_path):
    cairosvg.svg2png(url=svg_path, write_to=output_path)
    

# Splash screen
splash_screen_width = 550
splash_screen_height = 300

# splash_screen = ctk.CTk(fg_color='#3030ff')
splash_screen = ctk.CTk()
splash_screen.overrideredirect(True)


# Obtenir la taille de l'écran
screen_width = splash_screen.winfo_screenwidth()
screen_height = splash_screen.winfo_screenheight()

# Calculer la position pour centrer la fenêtre
x = (screen_width // 2) - (splash_screen_width // 2)
y = (screen_height // 2) - (splash_screen_height // 2)

# Appliquer la taille et la position centrée
splash_screen.geometry(f"{splash_screen_width}x{splash_screen_height}+{x}+{y}")

'''
Convertir le SVG en PNG temporairement
svg_path = "src/img/zr-logo.svg"
png_path = "src/img/zr-logo.png" 

svg_to_png(svg_path, png_path)

Charger l'image PNG convertie avec Pillow
image = Image.open(png_path)
image = ImageTk.PhotoImage(image)

Insert Logo
logo = ctk.CTkLabel(splash_screen, 
                    image=image,
                    text='')
logo.pack(pady=(50,0))
'''

splash_text = ctk.CTkLabel(splash_screen,
                           text='ZARAOMA', 
                           font=('Arial', 30, 'normal'),
                           text_color='#333333')
splash_description = ctk.CTkLabel(splash_screen, 
                                  text='Partage équitable en Eau et Electricité', 
                                  font=('Arial', 18, 'italic'),
                                  text_color='#333333')
splash_text.pack(pady=(10, 0))
splash_description.pack(pady=10)
splash_percentage = ctk.CTkLabel(splash_screen, 
                                 text='Loading ... 0%',
                                 text_color='#333333')
splash_percentage.pack()

# Progress Bar
progressBar = ctk.CTkProgressBar(splash_screen, 
                                 orientation='horizontal', 
                                 mode='determinate', 
                                 width=400,
                                 height=10,
                                 progress_color='#F27438',
                                 fg_color='#F0E2D0')
progressBar.set(0)
progressBar.pack(pady=5)

splash_screen.after(5000, hide)
splash_screen.mainloop()

# Fenêtre principale Tkinter
window = ctk.CTk()
window.resizable(False, False)
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
facture_label1.grid(row=0, column=0, padx=10)
facture_entry1 = ctk.CTkEntry(mainFrame,
                              border_color='#F27438',)
facture_entry1.bind("<KeyRelease>", on_entry_change)
facture_entry1.grid(row=0, column=1, padx=10, pady=(10, 0))
error_fact1 = ctk.CTkLabel(mainFrame, 
                           text='', 
                           text_color='#D9534F',)
error_fact1.grid(row=1, column=1, pady=1)
errors_main.append(error_fact1)

facture_label2 = ctk.CTkLabel(mainFrame, text='Facture du 1er mois recent:')
facture_label2.grid(row=2, column=0, padx=10)
facture_entry2 = ctk.CTkEntry(mainFrame,
                              border_color='#F27438',)
facture_entry2.bind("<KeyRelease>", on_entry_change)
facture_entry2.grid(row=2, column=1)
error_fact2 = ctk.CTkLabel(mainFrame, text='', text_color='#D9534F')
error_fact2.grid(row=3, column=1, pady=1)
errors_main.append(error_fact2)

facture_label3 = ctk.CTkLabel(mainFrame, text='Facture a payer (ce mois):')
facture_label3.grid(row=4, column=0, padx=10)
facture_entry3 = ctk.CTkEntry(mainFrame,
                              border_color='#F27438',)
facture_entry3.bind("<KeyRelease>", on_entry_change)
facture_entry3.grid(row=4, column=1, padx=10)
error_fact3 = ctk.CTkLabel(mainFrame, text='', text_color='#D9534F')
error_fact3.grid(row=5, column=1, pady=1)
errors_main.append(error_fact3)

# Ajout d'utilisateur
number_user_label = ctk.CTkLabel(mainFrame, text='Nombre d\'utilisateurs (nouveau):')
number_user_label.grid(row=6, column=0, padx=10)

number_user_entry = ctk.CTkEntry(mainFrame,
                                 border_color='#F27438',)
number_user_entry.bind("<KeyRelease>", on_entry_change)
number_user_entry.grid(row=6, column=1)
error_number_user = ctk.CTkLabel(mainFrame, 
                                 text='', text_color='#D9534F')
error_number_user.grid(row=7, column=1, pady=1)
errors_main.append(error_number_user)

btn = ctk.CTkButton(mainFrame, 
                    text='Valider', 
                    text_color='white',
                    fg_color='#F27438',
                    hover_color='#D86228',
                    state='disabled',
                    command=valideFistData)
btn.grid(row=8, column=1, pady=10)

# Exécution de l'application
window.mainloop()
