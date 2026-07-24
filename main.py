import xlwings as xw
import time

# Fájlnevek beállítása
target_file = "Oszlopközök(earth)2 másolata.xlsx"
calc_file = "Vezetéksodrony_belógás_KZ_végleges másolata.xlsm"

print("Excel alkalmazás indítása a háttérben...")
app = xw.App(visible=False)

try:
    # Fájlok megnyitása
    wb_target = app.books.open(target_file)
    wb_calc = app.books.open(calc_file)
    
    # JAVÍTÁS: A pontos munkalapnév megadása
    sheet_calc = wb_calc.sheets["Belógás számítás"] 
    
    # Végigmegyünk a célfájl ÖSSZES munkalapján
    for sheet_target in wb_target.sheets:
        # Kihagyjuk a Sheet1-et, ha az csak egy üres alapértelmezett lap
        if sheet_target.name == "Sheet1":
            continue
            
        print(f"\n--- Munkalap feldolgozása: {sheet_target.name} ---")
        
        # JAVÍTÁS: Stabilabb utolsó sor keresés az E oszlopban
        last_row = sheet_target.range("E10000").end('up').row
        
        # Ha a lap üres vagy csak fejléc van benne
        if last_row < 2:
            print("Üres munkalap vagy csak fejléc van rajta, átugrás...")
            continue
            
        print(f"Sorok száma ezen a lapon: {last_row - 1}")
        
        # Soronkénti feldolgozás az adott lapon
        for row in range(2, last_row + 1):
            h1 = sheet_target.range(f"B{row}").value       # 1. oszlop magassága
            h2 = sheet_target.range(f"D{row}").value       # 2. oszlop magassága
            length = sheet_target.range(f"E{row}").value   # Oszlopköz hossza
            sigma = sheet_target.range(f"I{row}").value    # Szigma
            
            # JAVÍTÁS: Csak akkor számolunk, ha MINDEN adat megvan ÉS azok valós számok (int vagy float)
            valid_data = all(isinstance(x, (int, float)) for x in [h1, h2, length, sigma])
            
            if not valid_data:
                # Ha hiányzik adat vagy szöveg van benne (pl. "nincs adat"), átugorjuk
                continue
                
            # Adatok beírása a számító Excelbe
            sheet_calc.range("E16").value = length
            sheet_calc.range("E18").value = h1
            sheet_calc.range("E19").value = h2
            sheet_calc.range("E21").value = sigma
            
            # Kényszerített újraszámolás
            wb_calc.app.calculate()
            time.sleep(0.15) # Kicsit megemelt szünet a komplex makrók/képletek miatt
            
            # Eredmények kiolvasása
            fok_40 = sheet_calc.range("D34").value
            fok_60 = sheet_calc.range("D35").value
            fok_80 = sheet_calc.range("D36").value
            
            # Visszaírás a célfájl aktuális munkalapjára (F, G, H oszlopok)
            sheet_target.range(f"F{row}").value = fok_40
            sheet_target.range(f"G{row}").value = fok_60
            sheet_target.range(f"H{row}").value = fok_80
            
            if row % 10 == 0:
                print(f"  {row} sor kész...")

    # Mentés a végén
    wb_target.save()
    print("\nMinden munkalap sikeresen feldolgozva és elmentve!")

except Exception as e:
    print(f"\nHiba történt a futás során: {e}")

finally:
    # Az Excel folyamat lezárása
    app.quit()
    print("Excel sikeresen bezárva.")