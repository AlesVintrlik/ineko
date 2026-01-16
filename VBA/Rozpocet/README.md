# Rozpočet - Aplikace pro správu rozpočtu

## Přehled

Rozpočet je VBA aplikace v Microsoft Excel pro správu a vizualizaci rozpočtových dat v IN-EKO. Aplikace využívá pivotové tabulky, grafy a poskytuje přehlednou navigaci mezi různými pohledy na data.

## Klíčové funkce

### 📊 Pivotové tabulky
- Kontingentní tabulka s dynamickým rozbalováním/sbalováním skupin
- Automatické obnovování dat z databáze
- Filtrování podle různých kritérií

### 📈 Grafy
- Vizualizace rozpočtových dat podle kategorií
- Kumulativní grafy
- Vlastní barvy grafů (konfigurovatelné)
- Export grafů jako obrázky

### 🧭 Navigace
- Intuitivní navigace mezi listy pomocí ikonek
- Listy: Aplikace, Kumulace, Kontingentní tabulka, Grafy
- Vizuální indikace aktivního listu (změna barvy ikonek)

### 💾 Export
- Kopírování dat do nového sešitu (bez vzorců)
- Export listů "Aplikace" a "Kumulace"
- Automatický název souboru s datem
- Převod vzorců na hodnoty

### 📝 Poznámky
- Možnost přidávat poznámky k jednotlivým položkám
- Ukládání poznámek do databáze

## Požadavky

- **Microsoft Excel** 2010 nebo novější
- **SQL Server** s IN-EKO ERP databází
- **Přístupová práva**: SELECT na view pro rozpočtová data
- **ADODB** (ActiveX Data Objects) knihovna

## Instalace

1. Otevřete soubor **Rozpočet.xlsm** (771 KB)
2. Povolte makra
3. Při prvním spuštění se zobrazí přihlašovací formulář

## Použití

### První spuštění

1. Přihlaste se (stejně jako v Gantt/Plánování)
2. Aplikace se otevře na listu "Aplikace"
3. Data se načtou z databáze

### Navigace

**Ikony navigace** (ve skupině "navigace"):
- 🏠 **ico_aplikace** - Přepne na hlavní pohled
- 📊 **ico_load** - Načte data z dotazů
- 📋 **ico_kontingencni** - Otevře kontingentní tabulku
- 📊 **ico_grafy** - Otevře grafy
- 💾 **ico_copy** - Exportuje data do nového sešitu

**Barvy ikonek:**
- Šedá (RGB 134, 134, 134) - Aktivní/dostupná
- Světle šedá (RGB 250, 250, 250) - Neaktivní

### Práce s daty

#### Načtení dat
1. Klikněte na ikonu "Load" (načíst)
2. Data se obnoví z databázových dotazů
3. Pivotové tabulky se automaticky refreshnou

#### Zobrazení kontingentní tabulky
1. Klikněte na ikonu "Kontingentní tabulka"
2. Zobrazí se pivotová tabulka "Rozpočet"
3. Použijte tlačítka **Plus** a **Minus** pro rozbalení/sbalení skupin

#### Práce s grafy
1. Klikněte na ikonu "Grafy"
2. Zobrazí se list s grafy
3. Grafy se automaticky aktualizují podle dat

**Změna barev grafů:**
```vba
' Spustit z VBA Editoru
Call VyberBarvyGrafu
```
- Postupně vyberte hlavní a doplňkovou barvu
- Barvy se uloží do listu "Konfigurace"

**Export grafů:**
```vba
Call SaveChartAsImage
```
- Grafy se uloží jako PNG obrázky
- Uživatel vybere umístění

### Export dat

**Zkratka:** `Ctrl+K`

1. Klikněte na ikonu "Copy" nebo stiskněte `Ctrl+K`
2. Aplikace vytvoří nový sešit s listy "Aplikace" a "Kumulace"
3. Všechny vzorce se převedou na hodnoty
4. Zkopírují se šířky sloupců
5. Vyberte umístění a název souboru
6. Výchozí název: `Kopie rozpočtu ze dne YYYYMMDD.xlsx`

**Co se exportuje:**
- List "Aplikace": Rozsah B4:AQ{poslední řádek}
- List "Kumulace": Rozsah B4:AQ{poslední řádek}
- Šířky sloupců B:AQ

**Co se NEexportuje:**
- Vzorce (převedeny na hodnoty)
- Skryté listy
- Makra

## Struktura projektu

```
VBA/Rozpocet/
├── Rozpočet.xlsm           # Hlavní Excel soubor (771 KB)
├── Modules/
│   ├── Connection.bas      # Správa připojení (stejné jako Gantt)
│   ├── Data.bas            # Načítání dat
│   ├── Navigace.bas        # Navigace mezi listy, ikony
│   ├── Copy.bas            # Export do nového sešitu
│   ├── Grafy.bas           # Správa grafů a barev
│   ├── Poznamky.bas        # Správa poznámek
│   ├── Rutiny.bas          # Utility funkce
│   ├── TestConnection.bas  # Testování připojení
│   ├── OldVersion.bas      # Starší verze kódu
│   ├── Temp.bas            # Dočasné funkce
│   └── ExportToGit.bas     # Export VBA do Git
├── Forms/
│   ├── frmLogin.frm        # Přihlašovací formulář
│   ├── frmProgress.frm     # Progress bar
│   ├── frmChangelog.frm    # Changelog verzí
│   └── frmGraf.frm         # Formulář pro grafy
└── ExcelObjects/
    ├── ThisWorkbook.cls    # Události workbooku
    └── List*.cls           # Třídy jednotlivých listů (9 listů)
```

## Klávesové zkratky

| Zkratka | Funkce | Popis |
|---------|--------|-------|
| `Ctrl+K` | Export dat | Vytvoří nový sešit s daty (bez vzorců) |
| `Ctrl+L` | Toggle Ribbon | Zobrazí/skryje Ribbon, panel vzorců a stavový řádek |

## Excel listy

### List "Aplikace"
- Hlavní pohled na rozpočtová data
- Rozsah dat: B4:AQ{N}
- Obsahuje vzorce a výpočty

### List "Kumulace"
- Kumulativní pohled na data
- Stejný rozsah jako "Aplikace"

### List "Kontingentní tabulka"
- Pivotová tabulka "Rozpočet"
- Dynamické rozbalování podle skupin
- Možnost refreshu dat

### List "Grafy"
- **GrafKategorie** - Graf podle kategorií
- **GrafKategorieKumulativni** - Kumulativní graf
- Vlastní barvy (konfigurovatelné)

### List "Konfigurace"
- Databázové připojení (serverName, databaseName)
- Barvy grafů:
  - C7: Hlavní barva (default: RGB 35, 176, 160)
  - C8: Doplňková barva (default: RGB 209, 209, 209)

### List "Rozpočet"
- Měsíc: B2 (1-12)
- Ovládá filtrování dat

## Databázové závislosti

### Views/Queries
- Aplikace používá databázové dotazy pro načítání dat
- Přesný název view není v poskytnutém kódu
- Data se načítají pomocí `RefreshAll` nebo specifických query objektů

## Bezpečnost

### Autentizace
- Stejný systém jako v Gantt (XOR šifrování)
- Podpora NT autentizace

### Ochrana listů
- Listy jsou chráněny pomocí `LockSpecificSheets`
- Ochrana parametry:
  - `UserInterfaceOnly:=True`
  - `DrawingObjects:=True`
  - `Contents:=True`
  - `AllowUsingPivotTables:=True`
  - `AllowFormattingColumns:=True`
- Pro odemknutí: `Call UnlockAllSheets`

**Poznámka:** Ochrana není heslem chráněna (na rozdíl od Plánování).

## Pokročilé funkce

### Dynamické přiřazování maker

Aplikace automaticky přiřazuje makra ikonkám na základě:
- Názvu ikony (začíná "ico_")
- Aktivního listu
- Dostupnosti funkce

**Příklad:**
```vba
' Ikona "ico_load" na listu "Aplikace"
→ OnAction = "LoadDataFromQueries"

' Ikona "ico_load" na jiném listu
→ OnAction = "TotoMakroNicNedela"
→ Barva = ICO_DISABLE_COLOR
```

### Docasná změna buňky

Pro vynucení refreshu po importu dat:
```vba
Call DocasnaZmenaBunky
' Změní B2 na 13 (neexistující měsíc) a vrátí zpět
```

### Changelog

Kliknutím na buňku B5 se zobrazí changelog verzí:
```vba
Sub FollowHyperlink(ByVal Target As Range)
    If Target.Address = "$B$5" Then
        frmChangelog.Show
    End If
End Sub
```

## Známé limity

1. **Pevně kódované rozsahy**: B4:AQ je hardcoded
2. **Žádná validace exportu**: Nekontroluje, zda jsou data platná
3. **Absence error handlingu**: V některých procedurách chybí
4. **Legacy kód**: Obsahuje staré verze v OldVersion.bas a Temp.bas

## Řešení problémů

### Data se nenačítají
- Zkontrolujte databázové připojení
- Ověřte, že existují potřebné dotazy (Queries)
- Zkuste ruční refresh: Data → Aktualizovat vše

### Ikony nereagují
- Zkontrolujte, zda existuje skupina "navigace"
- Ověřte názvy ikonek (musí začínat "ico_")
- Zkontrolujte, zda jste na správném listu

### Grafy nemají správné barvy
1. Otevřete VBA Editor (Alt+F11)
2. Spusťte: `Call InicializujBarvyGrafu`
3. Pak: `Call NastavBarvyGrafu`

### Export selhal
- Zkontrolujte, zda máte práva k zápisu do cílové složky
- Ověřte, že listy "Aplikace" a "Kumulace" existují
- Zkontrolujte volné místo na disku

### Pivotová tabulka je prázdná
1. Otevřete list "Kontingentní tabulka"
2. Pravý click na pivotovou tabulku → Aktualizovat
3. Nebo spusťte: `Call KontingencniTabulka`

## Best Practices

### Doporučený workflow
1. **Otevření**: Přihlášení a automatické načtení dat
2. **Prohlížení**: Navigace mezi listy pomocí ikonek
3. **Analýza**: Použití pivotových tabulek a grafů
4. **Export**: Před odesláním dat mimo aplikaci

### Tipy
- Pravidelně aktualizujte data (ikona "Load")
- Pro rychlou analýzu použijte kontingentní tabulku
- Před exportem zkontrolujte aktuálnost dat
- Vlastní barvy grafů uložte do konfigurace

## Podpora

Pro technickou podporu kontaktujte tým vývoje IN-EKO ERP.

## Autor

IN-EKO VBA Development Team

## Verze

**Export:** 2026-01-16
**Velikost:** 771 KB
**Řádků kódu:** ~2,430
**Excel soubor:** Rozpočet.xlsm
