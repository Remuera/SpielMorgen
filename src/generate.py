import csv

# 10 deutsche Vornamen × 12 deutsche Nachnamen = 120 eindeutige Kombinationen
first_names = [
    "Lukas","Jonas","Leon","Elias","Noah",
    "Paul","Finn","Ben","Luis","Felix"
]
last_names = [
    "Müller","Schmidt","Schneider","Fischer","Weber","Meyer",
    "Wagner","Becker","Schulz","Hoffmann","Schäfer","Koch"
]

# Basis-Reihenfolge der 7 Aktivitäten
prefs = [
    "Sackgumpen","Jassen","Volleyball",
    "Puzzle","Fussball","Schach","Schwimmen"
]

with open("kinder_praeferenzen.csv", "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    # Header entspricht Excel-Zeile 1
    writer.writerow(["F", "G", "H"])

    idx = 0
    for fn in first_names:
        for ln in last_names:
            # Rotation um idx mod 7, so dass jede Zeile eine gültige Permutation (ohne Doppelten) erhält
            shift = idx % len(prefs)
            order = prefs[shift:] + prefs[:shift]
            writer.writerow([";".join(order), fn, ln])
            idx += 1

print("Datei 'kinder_praeferenzen.csv' mit 120 Einträgen erstellt.")
