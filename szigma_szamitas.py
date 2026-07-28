import pandas as pd
import xlwings as xw
import re
import os
import sys

# 1. FÁJLOK MEGLÉTÉNEK ELLENŐRZÉSE
szukseges_fajlok = {
    "oszlopkozok": "Oszlopközök(earth)2.xlsx",
    "balmaz": "_Balmazújváros - Nádudvar 20-111. távvezeték bejárási jegyzőkönyv(1-92).xlsx",
    "iee": "IEEE738.xlsb",
    "belogas": "Vezetéksodrony_belógás_KZ_végleges.xlsm"
}

minden_megvan = True
print("--- Fájlok ellenőrzése ---")
for kulcs, fajlnev in szukseges_fajlok.items():
    if os.path.exists(fajlnev):
        print(f"✅ Megvan: {fajlnev}")
    else:
        print(f"❌ HIÁNYZIK: {fajlnev}")
        minden_megvan = False

if not minden_megvan:
    print("\nHiba: A hiányzó fájl(ok) miatt a kód leállt.")
    sys.exit()

print("\nMinden fájl megvan. Adatok betöltése...")

# 2. BEJÁRÁSI JEGYZŐKÖNYV ADATAINAK BEOLVASÁSA (Pandas)
try:
    df_balmaz = pd.read_excel(szukseges_fajlok["balmaz"])
    print(f"-> Bejárási jegyzőkönyv sikeresen beolvasva ({len(df_balmaz)} sor).")
except Exception as e:
    print(f"❌ Hiba a bejárási jegyzőkönyv beolvasásakor: {e}")
    sys.exit()

balmaz_adatok = {}
for idx, row in df_balmaz.iterrows():
    try:
        if len(row) > 5 and pd.notna(row.iloc[5]):
            oszlop_id = str(row.iloc[5]).split('.')[0].strip()  
            funkcio_szoveg = str(row.iloc[14]).lower() if len(row) > 14 and pd.notna(row.iloc[14]) else ""
            kulso_hom = row.iloc[10] if len(row) > 10 and pd.notna(row.iloc[10]) else 20.0
            mert_magassag = row.iloc[9] if len(row) > 9 and pd.notna(row.iloc[9]) else None
            
            funkcio = "feszítő" if "feszítő" in funkcio_szoveg else "tartó"
            
            if oszlop_id and oszlop_id != 'nan':
                balmaz_adatok[oszlop_id] = {
                    'funkcio': funkcio,
                    'kulso_hom': kulso_hom,
                    'mert_magassag': mert_magassag
                }
    except Exception as e:
        print(f"Figyelmeztetés a bejárási jegyzőkönyv {idx}. soránál: {e}")

# 3. INTERAKCIÓ AZ EXCEL MODELLEKKEL (xlwings)
print("\nExcel alkalmazás indítása a háttérben...")
app = xw.App(visible=False)
app.display_alerts = False 

try:
    print("Munkafüzetek megnyitása...")
    wb_oszlop = app.books.open(szukseges_fajlok["oszlopkozok"])
    
    # Dinamikus munkalap-keresés részleges név alapján
    ws_oszlop = None
    for sheet in wb_oszlop.sheets:
        if "balmaz" in sheet.name.lower():
            ws_oszlop = sheet
            break
            
    if ws_oszlop is None:
        ws_oszlop = wb_oszlop.sheets[1] if len(wb_oszlop.sheets) > 1 else wb_oszlop.sheets[0]
            
    print(f"-> Kiválasztott cél fül neve: '{ws_oszlop.name}'")
    
    wb_iee = app.books.open(szukseges_fajlok["iee"])
    ws_transient = wb_iee.sheets['T Transient']
    
    wb_belogas = app.books.open(szukseges_fajlok["belogas"])
    ws_szamitas = wb_belogas.sheets[0] 
    
    # Utolsó sor megkeresése az A oszlopban
    max_sor = ws_oszlop.range('A' + str(ws_oszlop.cells.last_cell.row)).end('up').row
    print(f"-> Feldolgozandó sorok száma az Oszlopközökben: {max_sor - 1}")
    
    # Oszloptípus (K) és Hőmérséklet (M) kitöltése
    feszito_pontok = []
    for sor in range(2, max_sor + 1):
        elso_oszlop = str(ws_oszlop.range(f'A{sor}').value).split('.')[0].strip()
            
        if elso_oszlop in balmaz_adatok:
            ws_oszlop.range(f'K{sor}').value = balmaz_adatok[elso_oszlop]['funkcio']
            ws_oszlop.range(f'M{sor}').value = balmaz_adatok[elso_oszlop]['kulso_hom']
            if balmaz_adatok[elso_oszlop]['funkcio'] == 'feszítő':
                feszito_pontok.append((sor, elso_oszlop))
        else:
            ws_oszlop.range(f'K{sor}').value = "tartó"
            
    print(f"-> Talált feszítő oszlopok: {[p[1] for p in feszito_pontok]}")

    # Feszítőközök (J) kiszámítása
    for sor in range(2, max_sor + 1):
        aktualis_id = str(ws_oszlop.range(f'A{sor}').value).split('.')[0].strip()
        try:
            akt_szam = int(re.sub("[^0-9]", "", aktualis_id))
            for i in range(len(feszito_pontok) - 1):
                start = int(re.sub("[^0-9]", "", feszito_pontok[i][1]))
                end = int(re.sub("[^0-9]", "", feszito_pontok[i+1][1]))
                if start <= akt_szam <= end:
                    ws_oszlop.range(f'J{sor}').value = f"{feszito_pontok[i][1]}-{feszito_pontok[i+1][1]}"
                    break
        except:
            pass

    print("-> Feszítőközök beírva. Szakaszok szigma számítása következik...")

    # Egyedi feszítőközök kigyűjtése
    feszitokoz_list = [ws_oszlop.range(f'J{s}').value for s in range(2, max_sor + 1)]
    egyedi_feszitokozok = sorted(list(set([f for f in feszitokoz_list if f])))
    
    for fkoz in egyedi_feszitokozok:
        print(f"   Szakasz feldolgozása: {fkoz}...")
        erintett_sorok = [s for s in range(2, max_sor + 1) if ws_oszlop.range(f'J{s}').value == fkoz]
        
        mert_magassag_talalt = False
        szigma = None
        
        for sor in erintett_sorok:
            elso_oszlop = str(ws_oszlop.range(f'A{sor}').value).split('.')[0].strip()
            
            if elso_oszlop in balmaz_adatok and balmaz_adatok[elso_oszlop]['mert_magassag'] is not None:
                mert_magassag = balmaz_adatok[elso_oszlop]['mert_magassag']
                kulso_hom = ws_oszlop.range(f'M{sor}').value
                aram_erosseg = ws_oszlop.range(f'L{sor}').value
                
                if kulso_hom is not None and aram_erosseg is not None:
                    # IEEE738 értékek bevitele
                    ws_transient.range('C4').value = kulso_hom
                    ws_transient.range('O4').value = aram_erosseg
                    sodrony_hom = ws_transient.range('K9').value
                    
                    if sodrony_hom is None:
                        print(f"   ⚠️ Figyelem: IEEE738 K9 cellája üres a(z) {elso_oszlop} oszlopnál!")
                        continue
                    
                    # Belógásszámító adatok átadása
                    ws_szamitas.range('E16').value = ws_oszlop.range(f'E{sor}').value  # Oszlopköz hossza
                    ws_szamitas.range('E18').value = ws_oszlop.range(f'B{sor}').value  # 1. oszlop magassága
                    ws_szamitas.range('E19').value = ws_oszlop.range(f'D{sor}').value  # 2. oszlop magassága
                    ws_szamitas.range('E20').value = sodrony_hom
                    ws_szamitas.range('E22').value = mert_magassag
                    
                    # 1. KÍSÉRLET: Makró futtatása a megadott pontos néven
                    try:
                        makro = app.macro("Vezetéksodrony_belógás_KZ_végleges.xlsm!ComputeSigmaFromMeasuredSag")
                        makro()
                        # Kiolvasás a fotó szerinti E23-as cellából
                        szigma = ws_szamitas.range('E23').value 
                    except Exception as macro_err:
                        print(f"   ⚠️ Makró futtatási zökkenő, megpróbáljuk a beépített Célértékkeresőt...")
                        szigma = None
                    
                    # 2. KÍSÉRLET (Tartalék): Ha a makró nem adott vissza értéket, a GoalSeek-kel kényszerítjük ki
                    if szigma is None or szigma == "":
                        try:
                            ws_szamitas.range('E23').value = ws_szamitas.range('E21').value  # Tervezett érték kezdésnek
                            cel_cella = ws_szamitas.range('E32').api      # Minimális föld feletti magasság cellája
                            valtozo_cella = ws_szamitas.range('E23').api  # Kiszámolt meglévő szigma cellája
                            
                            cel_cella.GoalSeek(Goal=mert_magassag, ChangingCell=valtozo_cella)
                            szigma = ws_szamitas.range('E23').value
                        except Exception as gs_err:
                            print(f"   ❌ A Célértékkereső is kudarcot vallott: {gs_err}")
                            szigma = "Számítási hiba"
                    
                    if szigma is not None and szigma != "Számítási hiba":
                        mert_magassag_talalt = True
                        print(f"   -> Sikeres számítás! Szigma: {szigma}")
                    break
        
        # Eredmény visszaírása a szakasz összes sorára az I oszlopba
        for sor in erintett_sorok:
            if mert_magassag_talalt:
                ws_oszlop.range(f'I{sor}').value = szigma
            else:
                ws_oszlop.range(f'I{sor}').value = "Nincs mért adat"

    wb_oszlop.save()
    print("\n🎉 A folyamat sikeresen lefutott! Az adatok a meglévő oszlopokban frissültek.")

except Exception as e:
    print(f"\n❌ Súlyos hiba történt a végrehajtás során: {e}")
    import traceback
    traceback.print_exc()

finally:
    print("Excel kapcsolatok lezárása...")
    try:
        wb_oszlop.close()
        wb_iee.close()
        wb_belogas.close()
    except:
        pass
    app.quit()