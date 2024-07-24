import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from pulp import *
import pandas as pd 

# Données

# Lecture du fichier Excel
donnees_excel = pd.read_excel("C:\\Users\\RAEDJLASSSI\\OneDrive\\Bureau\\donnesRo.xlsx")

# Sélection de la ligne 11 et des colonnes C à H
d1 = donnees_excel.iloc[9][2:8]  # 10 pour la ligne 11 (index 0-based), 2:8 pour les colonnes C à H
d2= donnees_excel.iloc[10][2:8]  # 10 pour la ligne 11 (index 0-based), 2:8 pour les colonnes C à H

# Création du problème
prob = LpProblem("planification de production de pro deco", LpMinimize)

# Déclaration des variables de décision
x1t = LpVariable.matrix("x1t", range(6), lowBound=0, cat=LpInteger)
x2t = LpVariable.matrix("x2t", range(6), lowBound=0, cat=LpInteger)
s1t = LpVariable.matrix("s1t", range(7), lowBound=0, cat=LpInteger)
s2t = LpVariable.matrix("s2t", range(7), lowBound=0, cat=LpInteger)
r1t = LpVariable.matrix("r1t", range(7), lowBound=0, cat=LpInteger)
r2t = LpVariable.matrix("r2t", range(7), lowBound=0, cat=LpInteger)
y = LpVariable.matrix("y", range(6), lowBound=0, cat=LpInteger)
l = LpVariable.matrix("l", range(6), lowBound=0, cat=LpInteger)
soust1 = LpVariable.matrix("soust1", range(6), lowBound=0, cat=LpInteger)
soust2 = LpVariable.matrix("soust2", range(6), lowBound=0, cat=LpInteger)
et = LpVariable.matrix("et", range(7), lowBound=0, cat=LpInteger)

# Fonction objectif
prob += lpSum([(x1t[i] + x2t[i]) * 45 + (s1t[i] + s2t[i]) * 8 + (r1t[i] + r2t[i]) * 40 + 18 * 20 * 8 * et[i] + soust1[i] * 125 + soust2[i] * 105 for i in range(6)]) + y * 800 + l * 1600, "objective"

# Initialisation des variables
prob += et[0] == 65
prob += s1t[0] == 0
prob += s2t[0] == 0
prob += r1t[0] == 0
prob += r2t[0] == 0

# Contrainte de capacité de production
for i in range(6):
    prob += 3.5 * x1t[i] + 2.5 * x2t[i] <= 160 * et[i], f"contrainte_{i}"

# Contraintes de gestion des stocks
for i in range(0, 6):
    prob += (s1t[i + 1] - r1t[i + 1]) + (s1t[i] - r1t[i]) == (x1t[i] + s1t[i] - d1[i])
    prob += (s2t[i + 1] - r2t[i + 1]) + (s2t[i] - r2t[i]) == (x2t[i] + s2t[i] - d2[i])

# Contraintes de sous-traitance
prob += lpSum(d1[i] for i in range(6)) <= lpSum(x1t[i] + soust1[i] for i in range(6))
prob += lpSum(d2[i] for i in range(6)) <= lpSum(x2t[i] + soust2[i] for i in range(6))

# Contraintes de l'effectif
for i in range(6):
    prob += et[i + 1] == et[i] - l[i] + y[i]

# Fonction pour résoudre le problème et afficher les résultats dans l'interface graphique
def resoudre_et_afficher_resultats():
    # Résolution du problème
    prob.solve()

    # Affichage des résultats dans l'interface graphique
    status_label.config(text="Statut: " + LpStatus[prob.status])
    objective_result_label.config(text="Résultat de la fonction objectif: " + str(value(prob.objective)))
    results_text.delete(1.0, tk.END)  # Efface le texte précédent
    results_text.insert(tk.END, "Variables:\n")
    for v in prob.variables():
        results_text.insert(tk.END, v.name + " = " + str(v.varValue) + "\n")

# Fonction pour afficher le graphique dequantite de  p1
def afficher_graphique_x1t():
    valeurs_x1t = [x1t[i].varValue for i in range(6)]
    plt.plot(range(1, 7), valeurs_x1t, marker='o', label='Quantité produite')
    valeurs_soust1 = [soust1[i].varValue for i in range(6)]
    plt.plot(range(1, 7), valeurs_soust1, marker='o', color='red', label='Quantité sous-traitée')
    valeurs_s1t = [s1t[i].varValue for i in range(6)]
    plt.plot(range(1, 7), valeurs_s1t, marker='o', color='green', label='Stock')
    plt.xlabel('Temps')
    plt.ylabel('Valeur')
    plt.title('Graphique de quantité de p1')
    plt.legend()
    plt.grid(True)
    plt.show()

# Fonction pour afficher le graphique de quantite  p2
def afficher_graphique_x2t():
    valeurs_x2t = [x2t[i].varValue for i in range(6)]
    plt.plot(range(1, 7), valeurs_x2t, marker='o', label='Quantité produite')
    valeurs_soust2 = [soust2[i].varValue for i in range(6)]
    plt.plot(range(1, 7), valeurs_soust2, marker='o', color='red', label='Quantité sous-traitée ')
    valeurs_s2t = [s2t[i].varValue for i in range(6)]
    plt.plot(range(1, 7), valeurs_s2t, marker='o', color='green', label='Stock ')
    plt.xlabel('Temps')
    plt.ylabel('Valeur')
    plt.title('Graphique quantité de p2')
    plt.legend()
    plt.grid(True)
    plt.show()

# Fonction pour afficher le graphique de et, l, y
def afficher_graphique_effe():
    valeurs_et = [et[i].varValue for i in range(7)]
    plt.plot(range(0, 7), valeurs_et, marker='o', label='Nombre d\'ouvriers')
    valeurs_l = [l[i].varValue for i in range(6)]
    plt.plot(range(0, 6), valeurs_l, marker='o', color='red', label='Nombre de licenciements')
    valeurs_y = [y[i].varValue for i in range(6)]
    plt.plot(range(0, 6), valeurs_y, marker='o', color='green', label='Nombre d\'embauches')

    plt.xlabel('Temps')
    plt.ylabel('Valeur')
    plt.title('Graphique de main d\'œuvre')
    plt.legend()
    plt.grid(True)
    plt.show()

# Création de la fenêtre principale
root = tk.Tk()
root.title("Planification de production")
root.configure(bg="#1e1e1e")

# Cadre pour afficher le statut
status_frame = ttk.Frame(root)
status_frame.pack(pady=10)
status_label = ttk.Label(status_frame, text="")
status_label.pack()

# Cadre pour afficher le résultat de la fonction objectif
objective_result_label = ttk.Label(root, text="")
objective_result_label.pack()

# Cadre pour afficher les résultats
results_frame = ttk.Frame(root)
results_frame.pack(pady=10)
results_label = ttk.Label(results_frame, text="Résultats:")
results_label.pack()
results_text = tk.Text(results_frame, height=10, width=50, bg='black', fg='white')
results_text.pack()

# Boutons pour résoudre le problème et afficher les résultats
solve_button = tk.Button(root, text="Lancer les calculs", bg="#0074D9", fg="white", command=resoudre_et_afficher_resultats)
solve_button.pack(pady=10)

# Boutons pour afficher les graphiques de x1t et x2t et main-d’œuvre
graph_x1t_button = tk.Button(root, text="Afficher le graphique de quantité de p1", bg="#0074D9", fg="white", command=afficher_graphique_x1t)
graph_x1t_button.pack(pady=5)

graph_x2t_button = tk.Button(root, text="Afficher le graphique de quantité de p2", bg="#0074D9", fg="white", command=afficher_graphique_x2t)
graph_x2t_button.pack(pady=5)

graph_et_button = tk.Button(root, text="Afficher le graphique de main d\'œuvre ", bg="#0074D9", fg="white", command=afficher_graphique_effe)
graph_et_button.pack(pady=5)

root.mainloop()
