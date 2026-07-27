import xlwings as xw
import time

# Fájlnevek beállítása
target_file = "Oszlopközök(earth)2.xlsx"
calc_file = "Vezetéksodrony_belógás_KZ_végleges.xlsm"

print("Excel alkalmazás indítása a háttérben...")
app = xw.App(visible=False)

# Biztonsági és teljesítményoptimalizáló beállítások
app.display_alerts = False
app.screen_updating = False

try:
    # Fájlok megnyitása
    wb_target = app.books.open(target_file)
    wb_calc = app.books.open(calc_file)

    # A pontos munkalapnév megadása a számító fájlban
    sheet_calc = wb_calc.sheets["Belógás számítás"]

    # Végigmegyünk a célfájl ÖSSZES munkalapján
    for sheet_target in wb_target.sheets:
        if sheet_target.name == "Sheet1":
            continue

        print(f"\n--- Munkalap feldolgozása: {sheet_target.name} ---")

        # Dinamikus utolsó sor keresés
        max_row = sheet_target.cells.last_cell.row
        last_row = sheet_target.range(f"E{max_row}").end('up').row

        if last_row < 2:
            print("Üres munkalap vagy csak fejléc van rajta, átugrás...")
            continue

        print(f"Sorok száma ezen a lapon: {last_row - 1}")

        kihagyott_sorok = 0 # Számláló a már kiszámolt sorokhoz

        # Soronkénti feldolgozás az adott lapon
        for row in range(2, last_row + 1):

            # ÚJ JAVÍTÁS: Ellenőrizzük, hogy az F oszlop (fok_40) tartalmaz-e már eredményt.
            # Ha nem None és nem üres string, akkor már ki van számolva.
            existing_result = sheet_target.range(f"F{row}").value
            if existing_result not in [None, ""]:
                kihagyott_sorok += 1
                continue # Ugrik is a következő sorra, nem olvas, nem számol

            h1 = sheet_target.range(f"B{row}").value
            h2 = sheet_target.range(f"D{row}").value
            length = sheet_target.range(f"E{row}").value
            sigma = sheet_target.range(f"I{row}").value

            # Csak akkor számolunk, ha MINDEN adat megvan ÉS azok valós számok
            valid_data = all(isinstance(x, (int, float)) for x in [h1, h2, length, sigma])

            if not valid_data:
                continue

            # Adatok beírása a számító Excelbe
            sheet_calc.range("E16").value = length
            sheet_calc.range("E18").value = h1
            sheet_calc.range("E19").value = h2
            sheet_calc.range("E21").value = sigma

            # Kényszerített újraszámolás
            wb_calc.app.calculate()
            time.sleep(0.15)

            # Eredmények kiolvasása
            fok_40 = sheet_calc.range("D34").value
            fok_60 = sheet_calc.range("D35").value
            fok_80 = sheet_calc.range("D36").value

            # Visszaírás a célfájl aktuális munkalapjára
            sheet_target.range(f"F{row}").value = fok_40
            sheet_target.range(f"G{row}").value = fok_60
            sheet_target.range(f"H{row}").value = fok_80

            if (row - kihagyott_sorok) % 10 == 0:
                print(f"  Újonnan kiszámolva {row - kihagyott_sorok} sor...")

        if kihagyott_sorok > 0:
            print(f"  {kihagyott_sorok} sor átugorva (már ki volt számolva).")

    # Mentés a végén
    wb_target.save()
    print("\nMinden munkalap sikeresen feldolgozva és elmentve!")

except Exception as e:
    print(f"\nHiba történt a futás során: {e}")

finally:
    # Az Excel folyamat biztonságos lezárása
    app.quit()
    print("Excel sikeresen bezárva.")