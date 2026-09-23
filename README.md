# Belogas szamitas automatizalasa

Ez a projekt tavvezetek sodronyak belogasat szamolja ki harom egymasra epulo Python modul segitsegevel:

1. Aramerosseg szamitas - aramerosseg.py
2. Szigma szamitas - szigma_szamitas.py
3. Belogas szamitas - belogas_szamitas.py

A rendszer Excel alapjan vegzi a szamitasokat, es az eredmenyeket visszairja a megfelelo Excel fajlokba.

---

## Kovetelmenyek

A projekt Python 3.10+ verzioval mukodik.

Szukseges Python csomagok:
pandas
openpyxl
xlrd
numpy

## Szükséges bemeneti fajlok

A repository nem tartalmazza sem a  szamitast vegzo sem az adatokat biztosito excel fajlokat.


A fajlnevek tetszolegesek, a program futas elott kell oket megadni.

---

## A rendszer mukodese

A szamitas harom lepesbol all, es mindegyik lepes kulon Python fajlban talalhato.

---

### 1. Aramerosseg szamitas - aramerosseg.py

Feladatok:

- meresi Excel fajl beolvasasa
- idopillanatokhoz tartozo aramerosseg ertekek kiszamitasa
- eredmenyek visszairasa a cel Excel megfelelo oszlopaiba


### 2. Szigma szamitas - szigma_szamitas.py

Feladatok:

- aramerosseg alapjan mechanikai feszultseg (sigma) meghatarozasa
- eredmenyek beirasa a belogas szamito Excel megfelelo cellaiba

---

### 3. Belogas szamitas - belogas_szamitas.py

Feladatok:

- sigma ertekek es geometriai adatok alapjan belogas meghatarozasa
- eredmenyek beirasa a vegleges Excel fajlba









