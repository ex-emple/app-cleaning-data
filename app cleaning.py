import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from fpdf import FPDF

# Fonction pour supprimer les doublons
def supprimer_doublons():
    global df
    df.drop_duplicates(inplace=True)

# Fonction pour supprimer les lignes incomplètes
def supprimer_lignes_incompletes():
    global df
    df.dropna(how='any', inplace=True)

# Fonction pour supprimer les lignes vides
def supprimer_lignes_vides():
    global df
    df.dropna(how='all', inplace=True)

# Fonction pour fusionner plusieurs tableaux
def fusionner_tableaux():
    global df
    files = filedialog.askopenfilenames(title="Sélectionner les fichiers à fusionner", filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx"), ("Text", "*.txt")])
    dfs = []

    for file in files:
        if file.endswith('.csv'):
            dfs.append(pd.read_csv(file))
        elif file.endswith('.xlsx'):
            dfs.append(pd.read_excel(file))
        elif file.endswith('.txt'):
            dfs.append(pd.read_csv(file, delimiter='\t'))

    df = pd.concat(dfs, ignore_index=True)
    messagebox.showinfo("Fusion terminée", f"{len(dfs)} tableaux ont été fusionnés avec succès.")

# ✅ Fonction corrigée pour nettoyer les espaces (sans .str)
def nettoyer_espaces():
    global df
    for col in df.select_dtypes(include=['object', 'string']).columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# ✅ Fonction corrigée pour mettre la première lettre en majuscule (sans .str)
def majuscule_premiere_lettre():
    global df
    for col in df.select_dtypes(include=['object', 'string']).columns:
        df[col] = df[col].apply(lambda x: x.capitalize() if isinstance(x, str) else x)

# Fonction pour supprimer les lignes avec un nombre précis de valeurs manquantes
def supprimer_lignes_manquantes_exactement(n):
    global df
    df = df[df.isnull().sum(axis=1) != n]

# Fonction pour sauvegarder le fichier en Excel
def sauvegarder_excel():
    global df
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Sauvegarde", "Le fichier a été sauvegardé avec succès.")

# Fonction pour générer un rapport PDF
def generer_rapport_pdf():
    global df
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
    if file_path:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12, style='B')
        pdf.cell(200, 10, txt="Rapport de nettoyage des données", ln=True, align='C')

        pdf.ln(10)
        pdf.set_font("Arial", size=10)
        pdf.cell(200, 10, txt="Résumé des données après nettoyage:", ln=True)
        pdf.multi_cell(200, 10, txt=df.describe(include='all').to_string())

        pdf.ln(10)
        pdf.cell(200, 10, txt="Aperçu des données après nettoyage:", ln=True)
        preview = df.head(10).to_string()
        pdf.multi_cell(200, 10, txt=preview)

        pdf.output(file_path)
        messagebox.showinfo("Génération PDF", "Le rapport PDF a été généré avec succès.")

# Fonction pour sélectionner et appliquer les tâches
def appliquer_taches():
    for task, func in taches.items():
        if task_var[task].get():
            func()

# Fonction pour charger un fichier Excel, CSV ou TXT
def charger_fichier():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Microsoft Excel Worksheet", "*.xlsx"), ("CSV", "*.csv"), ("Text", "*.txt")])
    if file_path:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        elif file_path.endswith('.txt'):
            df = pd.read_csv(file_path, delimiter='\t')
        messagebox.showinfo("Fichier chargé", f"Le fichier {file_path} a été chargé avec succès.")

# Interface Tkinter
root = tk.Tk()
root.title("Outils de nettoyage des données")
root.geometry("700x600")
root.configure(bg="#f7f7f7")

font_large = ("Helvetica", 14, "bold")
font_medium = ("Helvetica", 12)
font_small = ("Helvetica", 10)

frame_buttons = ttk.Frame(root, padding="20")
frame_buttons.grid(row=0, column=0, sticky="nsew", padx=20, pady=10)

frame_tasks = ttk.Frame(root, padding="20")
frame_tasks.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)

# Liste des tâches
taches = {
    "Supprimer les doublons": supprimer_doublons,
    "Supprimer les lignes incomplètes": supprimer_lignes_incompletes,
    "Supprimer les lignes vides": supprimer_lignes_vides,
    "Fusionner les tableaux": fusionner_tableaux,
    "Nettoyer les espaces inutiles": nettoyer_espaces,
    "Première lettre majuscule": majuscule_premiere_lettre,
    "Supprimer lignes avec 1 valeur manquante": lambda: supprimer_lignes_manquantes_exactement(1),
    "Supprimer lignes avec 2 valeurs manquantes": lambda: supprimer_lignes_manquantes_exactement(2),
    "Supprimer lignes avec 3 valeurs manquantes": lambda: supprimer_lignes_manquantes_exactement(3),
    "Supprimer lignes avec 4 valeurs manquantes": lambda: supprimer_lignes_manquantes_exactement(4),
}

task_var = {}
for task in taches:
    task_var[task] = tk.BooleanVar()
    ttk.Checkbutton(frame_tasks, text=task, variable=task_var[task], style="TCheckbutton").pack(anchor="w", padx=10, pady=5)

# Style des boutons
style = ttk.Style()
style.configure("TButton",
                padding=6,
                relief="flat",
                background="#4CAF50",
                foreground="white",
                font=font_medium)
style.map("TButton", background=[("active", "#45a049")])

ttk.Button(frame_buttons, text="Charger un fichier", command=charger_fichier, style="TButton").pack(fill='x', padx=10, pady=10)
ttk.Button(frame_buttons, text="Appliquer les tâches", command=appliquer_taches, style="TButton").pack(fill='x', padx=10, pady=10)
ttk.Button(frame_buttons, text="Sauvegarder en Excel", command=sauvegarder_excel, style="TButton").pack(fill='x', padx=10, pady=10)
ttk.Button(frame_buttons, text="Générer un rapport PDF", command=generer_rapport_pdf, style="TButton").pack(fill='x', padx=10, pady=10)

style.configure("TCheckbutton",
                font=font_medium,
                background="#f7f7f7",
                foreground="black")

# Démarrage de l’interface
root.mainloop()
