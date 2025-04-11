# -*- coding: utf-8 -*-
"""
Created on Tue Mar 25 13:16:07 2025

@author: aeriksen
"""

# Prosjektoppgave
# Importerer nødvendige biblioteker
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Del a
# Leser inn Excel-filen
df = pd.read_excel('support_uke_24.xlsx')

# Lager arrayer for hver kolonne
u_dag = np.array(df.iloc[:, 0])           # Ukedag
kl_slett = np.array(df.iloc[:, 1])        # Klokkeslett
varighet_str = df.iloc[:, 2]              # Samtalens varighet
score = np.array(df.iloc[:, 3])           # Tilfredshet

# Konverterer varighet fra tid-format til sekunder
varighet_timedelta = pd.to_timedelta(varighet_str, errors='coerce')
df["Varighet"] = varighet_timedelta.dt.total_seconds()

# Hent ut varighet
varighet = df["Varighet"].dropna().values

# Del b - Antall henvendelser per ukedag
ukedager = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag']
antall_henvendelser = [np.sum(u_dag == dag) for dag in ukedager]

# Søylediagram
plt.figure(figsize=(10, 6))
plt.bar(ukedager, antall_henvendelser, color='red')
plt.title('Antall henvendelser per ukedag')
plt.xlabel('Ukedag')
plt.ylabel('Antall henvendelser')
plt.grid(axis='y', linestyle='--', alpha=0.7)
for i, v in enumerate(antall_henvendelser):
    plt.text(i, v + 0.5, str(v), ha='center')
plt.show()

# Del c og d - Samtaletid
if len(varighet) > 0:
    min_varighet = np.nanmin(varighet) / 60     # konverterer til minutter
    max_varighet = np.nanmax(varighet) / 60
    gjennomsnitt_varighet = np.nanmean(varighet) / 60
else:
    min_varighet = max_varighet = gjennomsnitt_varighet = None

# Del e - Antall henvendelser per tidsbolk
tidsbolker = ['08-10', '10-12', '12-14', '14-16']
antall_per_tidsbolk = []

for bolk in tidsbolker:
    start, slutt = bolk.split('-')
    antall = np.sum((kl_slett >= f'{start}:00') & (kl_slett < f'{slutt}:00'))
    antall_per_tidsbolk.append(antall)

# Kakediagram for tidsbolker
plt.figure(figsize=(8, 8))
plt.pie(antall_per_tidsbolk, labels=tidsbolker, autopct='%1.1f%%', startangle=90)
plt.title('Henvendelser per tidsbolk')
plt.axis('equal')
plt.show()

# Skriver ut resultater i konsollen
print("\nAntall henvendelser per dag:")
for dag, antall in zip(ukedager, antall_henvendelser):
    print(f"{dag}: {antall} henvendelser")

print("\nOversikt over samtaler i uke 24:")
if min_varighet is not None:
    print(f"Korteste samtalen: {min_varighet:.2f} minutter.")
    print(f"Lengste samtalen: {max_varighet:.2f} minutter.")
    print(f"Gjennomsnittlig samtaletid: {gjennomsnitt_varighet:.2f} minutter.")
else:
    print("Ingen gyldige samtalevarigheter ble funnet i datasettet.")

print("\nAntall henvendelser per tidsbolk:")
for bolk, antall in zip(tidsbolker, antall_per_tidsbolk):
    print(f"Tidsbolk {bolk}: {antall} henvendelser")
