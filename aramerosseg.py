import pandas as pd
import xlwings as xw

# Fájlnevek és paraméterek
file_bejaras = "Nádudvar - Kaba 8-32 Távvezeték.xlsx"
file_meres = "MvAvr15m_260721_103014_1.xlsx"
file_cel = "Oszlopközök(earth)2.xlsx"
cel_lap_neve = "Nádudvar-Kaba 8-32 Távvezeték"


def tiszta_oszlop_nev(ertek):
    """Segédfüggvény az oszlopszámok egységes szöveges formátumra hozásához
    (hogy pl. a 10 és a 10.0 vagy '10 ' ugyanannak számítson)."""
    if pd.isna(ertek) or ertek is None:
        return ""
    szoveg = str(ertek).strip()
    if szoveg.endswith('.0'):
        szoveg = szoveg[:-2]
    return szoveg


def main():
    print("-" * 50)
    print("🚀 FELDOLGOZÁS INDÍTÁSA...")
    print("-" * 50)

    # --- 1. Lépés ---
    print("\n[1/5] Adatfájlok beolvasása (Ez eltarthat néhány másodpercig)...")
    print("  -> Bejárási jegyzőkönyv és mérési adatok memóriába töltése.")
    df_bejaras = pd.read_excel(file_bejaras)
    df_meres = pd.read_excel(file_meres)

    # --- 2. Lépés ---
    print("\n[2/5] Dátumformátumok ellenőrzése és átalakítása...")
    print("  -> 'Kezdés ideje' és 'Dátum' oszlopok időbélyeggé alakítása.")
    # Figyelem: A 'Dátum' oszlopnevet cseréld ki arra, ami a MvAvr15m... fájlban ténylegesen szerepel!
    df_bejaras['Kezdés ideje'] = pd.to_datetime(df_bejaras['Kezdés ideje'])
    df_meres['Dátum'] = pd.to_datetime(df_meres['Dátum'])

    df_bejaras = df_bejaras.sort_values('Kezdés ideje').reset_index(drop=True)
    df_meres = df_meres.sort_values('Dátum').reset_index(drop=True)
    print("  -> Dátumok sikeresen növekvő sorrendbe rendezve.")

    # --- 3. Lépés ---
    print("\n[3/5] Időpontok párosítása és Oszlopszám-Áramerősség szótár készítése...")
    print("  -> A jegyzőkönyv időpontjaihoz legközelebb eső mérési sorok megkeresése.")
    merged = pd.merge_asof(
        df_bejaras,
        df_meres,
        left_on='Kezdés ideje',
        right_on='Dátum',
        direction='nearest'
    )

    # Megtisztítjuk az oszlopszámokat, és készítünk egy szótárat (dictionary)
    # Ez a szótár rendeli hozzá az oszlopszámhoz az áramerősséget
    merged['Oszlopszám_tiszta'] = merged['Oszlopszám'].apply(tiszta_oszlop_nev)
    oszlop_aram_dict = dict(zip(merged['Oszlopszám_tiszta'], merged['Áramerősség']))
    print("  -> Áramerősségek hozzárendelve a bejárási oszlopszámokhoz.")

    # --- 4. Lépés ---
    print("\n[4/5] Cél Excel fájl megnyitása a háttérben (xlwings)...")
    app = xw.App(visible=False)
    try:
        wb = app.books.open(file_cel)
        print("  -> 'Oszlopközök(earth)2.xlsx' sikeresen megnyitva (láthatatlan módban).")

        if cel_lap_neve in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets[cel_lap_neve]
            print(f"  -> '{cel_lap_neve}' munkalap azonosítva és kiválasztva.")
        else:
            raise ValueError(f"Nem található a '{cel_lap_neve}' nevű munkalap a célfájlban.")

        # --- 5. Lépés ---
        print("\n[5/5] Adatok egyeztetése és beírása a megfelelő sorokba...")
        fejlecek = sheet.range('A1').expand('right').value

        # '1. oszlopszám' oszlop megkeresése
        if '1. oszlopszám' not in fejlecek:
            raise ValueError("Nem található '1. oszlopszám' oszlop a célfájlban!")
        oszlop1_idx = fejlecek.index('1. oszlopszám') + 1

        # 'Áramerősség (A)' oszlop megkeresése vagy létrehozása
        if 'Áramerősség (A)' in fejlecek:
            aram_idx = fejlecek.index('Áramerősség (A)') + 1
            print("  -> 'Áramerősség (A)' oszlop megtalálva.")
        else:
            aram_idx = len(fejlecek) + 1
            sheet.range((1, aram_idx)).value = 'Áramerősség (A)'
            print("  -> 'Áramerősség (A)' oszlop létrehozva.")

        print("  -> Sorok egyeztetése az '1. oszlopszám' alapján...")
        # Kikeressük, meddig tartanak az adatok a célfájlban
        utolso_sor = sheet.range((sheet.cells.last_cell.row, oszlop1_idx)).end('up').row

        if utolso_sor > 1:
            # Beolvassuk a célfájlban lévő '1. oszlopszám' adatokat
            cel_oszlop_ertekek = sheet.range((2, oszlop1_idx), (utolso_sor, oszlop1_idx)).value
            if not isinstance(cel_oszlop_ertekek, list):
                cel_oszlop_ertekek = [cel_oszlop_ertekek]

            aramerosseg_kiiras = []

            # Végigmegyünk a célfájl oszlopszámain, és kikeresjük a szótárból az értéket
            for ertek in cel_oszlop_ertekek:
                tiszta_ertek = tiszta_oszlop_nev(ertek)
                keresett_aram = oszlop_aram_dict.get(tiszta_ertek, None)
                aramerosseg_kiiras.append(keresett_aram)

            print("  -> A párosított értékek beírása.")
            sheet.range((2, aram_idx)).options(transpose=True).value = aramerosseg_kiiras
        else:
            print("  -> Nincsenek adatsorok az '1. oszlopszám' oszlopban.")

        wb.save()
        print("  -> Változtatások elmentve a fájlba.")

    except Exception as e:
        print(f"\n❌ HIBA TÖRTÉNT A FOLYAMAT SORÁN:\n{e}")

    finally:
        print("\n⏳ Excel alkalmazás biztonságos bezárása...")
        try:
            wb.close()
        except:
            pass
        app.quit()

    print("-" * 50)
    print("✅ KÉSZ! A program sikeresen lefutott.")
    print("-" * 50)


if __name__ == '__main__':
    main()