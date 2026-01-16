# Plánování - Systém pro plánování výroby

## Přehled

Plánování je pokročilá VBA aplikace v Microsoft Excel pro komplexní plánování výroby v IN-EKO s integrací do Helios ERP systému. Aplikace umožňuje správu termínů zakázek, výpočet pracnosti, sledování obsazenosti výrobních úseků a obousměrnou synchronizaci dat s Helios.

## Klíčové funkce

### 📋 Správa zakázek
- Automatické načítání zakázek z databáze
- Zobrazení termínů výroby pro jednotlivé úseky (Příprava, Pila, Svařování, Montáž, Elektro, Balení)
- Řazení zakázek podle termínu expedice
- Filtrování a vyhledávání zakázek

### ⏱️ Pracnost zakázek
- Evidence plánovaných hodin celkem
- Rozdělení hodin podle skupin pracovníků (1-5)
- Evidence kooperací
- Zobrazení skutečných a plánovaných hodin

### 📊 Obsazenost výroby
- Sledování obsazenosti výrobních úseků
- Sledování obsazenosti jednotlivých pracovišť
- Grafická vizualizace obsazenosti
- Kontrola překročení kapacit

### 🔄 Integrace s Helios
- Obousměrná synchronizace dat
- Volání stored procedures pro výpočet termínů
- Aktualizace termínů v Helios systému
- Načítání operací a normovaných časů

### ⚡ Pokročilé funkce
- Klávesové zkratky pro rychlou práci (Ctrl+N, Ctrl+D pro +/- týden)
- Automatické formátování a aplikace optimálního fontu
- Ochrana listů s možností filtrování
- Changelog verzí
- Progress bar pro dlouhé operace

## Požadavky

- **Microsoft Excel** 2013 nebo novější (preferováno 2016+)
- **SQL Server** s IN-EKO ERP databází
- **Přístupová práva**:
  - SELECT na view: `hvw_TerminyZakazekProPlanovani`
  - SELECT, UPDATE na tabulky: `TabZakazka`, `TabZakazka_EXT`
  - EXECUTE na procedury: `EP_PlanTerminyVyrobyDoZakazek`, `ep_DoplnTerminOperacePodleUseku`
- **ADODB** (ActiveX Data Objects) knihovna
- **Scripting Runtime** (Dictionary objekt)

## Instalace

1. Otevřete soubor **Plánování.xlsm**
2. Povolte makra (Soubor → Možnosti → Centrum zabezpečení → Nastavení centra zabezpečení → Makra → Povolit všechna makra)
3. Při prvním spuštění se zobrazí přihlašovací formulář

## První spuštění

### Přihlášení

1. Při otevření se zobrazí přihlašovací formulář
2. Vyplňte údaje:
   - **Server**: Název SQL serveru (např. `SERVER01`)
   - **Databáze**: Název databáze (např. `ERP_IN_EKO`)
   - **Uživatel**: Uživatelské jméno (nebo prázdné pro NT autentizaci)
   - **Heslo**: Heslo (nebo prázdné pro NT autentizaci)
3. Klikněte na **Přihlásit**

### Načtení dat

Po úspěšném přihlášení se automaticky:
1. Načtou zakázky z databáze
2. Aktualizuje se seznam zakázek v plánu
3. Nastaví se optimální font a formátování
4. Zobrazí se hlavní list "Plan"

## Použití

### Základní workflow

1. **Otevření aplikace** → Automatické načtení aktuálních zakázek
2. **Prohlížení plánu** → List "Plan" zobrazuje všechny zakázky s termíny
3. **Úprava termínů** → Klikněte do sloupce s termínem a upravte datum
4. **Použití klávesových zkratek**:
   - `Ctrl+N` - Přidat týden k termínu
   - `Ctrl+D` - Odebrat týden od termínu
5. **Uložení do Helios** → Po úpravách můžete synchronizovat zpět do Helios

### Práce se zakázkami

#### Zobrazení detailu zakázky
1. Klikněte na řádek se zakázkou
2. Stiskněte pravé tlačítko myši nebo použijte tlačítko v ribbonu
3. Otevře se formulář s detailem zakázky

#### Výpočet termínů výroby
1. Otevřete formulář zakázky
2. Zadejte nebo upravte **datum ukončení** (expedice)
3. Klikněte na **Vypočítat termíny**
4. Aplikace zavolá stored proceduru `EP_PlanTerminyVyrobyDoZakazek`
5. Termíny se automaticky přepočítají zpětným plánováním

#### Evidence pracnosti
1. Klikněte na řádek se zakázkou
2. Otevřete formulář **Hodiny**
3. Vyplňte:
   - Hodiny celkem
   - Hodiny podle skupin pracovníků (1-5)
   - Hodiny kooperací
4. Klikněte na **Uložit**

### Obsazenost výroby

#### Zobrazení obsazenosti úseků
1. Použijte tlačítko **Obsazenost úseků** v ribbonu
2. Zobrazí se list s přehledem obsazenosti všech úseků
3. Zelená = volná kapacita, Červená = překročení

#### Zobrazení obsazenosti pracovišť
1. Použijte tlačítko **Obsazenost pracovišť** v ribbonu
2. Zobrazí se detailní přehled jednotlivých pracovišť

### Synchronizace s Helios

#### Aktualizace dat z Helios
- Data se načítají automaticky při otevření
- Pro ruční aktualizaci použijte tlačítko **Aktualizovat data**

#### Odeslání termínů do Helios
1. Upravte termíny v listu "Plan"
2. Použijte funkci **Naplnit kontrolu** (vytvoří se list s přehledem změn)
3. Zkontrolujte změny
4. Klikněte na **Aktualizovat data v Helios**
5. Aplikace:
   - Uloží termíny do `TabZakazka_EXT`
   - Zavolá `ep_DoplnTerminOperacePodleUseku` pro přepočet operací
   - Synchronizuje změny zpět do Helios

## Struktura projektu

```
VBA/Planovani/
├── Plánování.xlsm          # Hlavní Excel soubor (1.9 MB)
├── Modules/
│   ├── Connection.bas      # Správa připojení a autentizace
│   ├── Data.bas            # Načítání dat ze serveru
│   ├── Main.bas            # Hlavní funkce aplikace
│   ├── VypocetPlanu.bas    # Výpočet termínů výroby
│   ├── PracnostZakazek.bas # Správa pracnosti (hodiny)
│   ├── Helios.bas          # Integrace s Helios ERP
│   └── ExportToGit.bas     # Export VBA do Git
├── Forms/
│   ├── frmLogin.frm        # Přihlašovací formulář
│   ├── frmProgress.frm     # Progress bar
│   ├── frmZakazka.frm      # Detail zakázky
│   ├── frmHodiny.frm       # Evidence hodin
│   ├── frmChangelog.frm    # Changelog verzí
│   └── frmReady.frm        # Oznámení o dokončení načítání
└── ExcelObjects/
    ├── ThisWorkbook.cls    # Události workbooku
    └── List*.cls           # Třídy jednotlivých listů (11 listů)
```

## Klávesové zkratky

| Zkratka | Funkce | Popis |
|---------|--------|-------|
| `Ctrl+N` | Přidat týden | Přidá 7 dní k termínu v aktivní buňce |
| `Ctrl+D` | Odebrat týden | Odebere 7 dní od termínu v aktivní buňce |
| `Ctrl+L` | Toggle Ribbon | Zobrazí/skryje Ribbon, panel vzorců a stavový řádek |

**Poznámka:** Zkratky `Ctrl+N` a `Ctrl+D` fungují pouze ve sloupcích s termíny (M, P, S, V, Y, AB).

## Výrobní úseky

Aplikace pracuje s následujícími výrobními úseky:

| Úsek | ID | Sloupec původní | Sloupec plánovaný | Popis |
|------|----|----|----|----|
| Příprava | 1 | L | M | Přípravné práce |
| Pila | 4 | O | P | Řezání materiálu |
| Svařování | 2 | R | S | Svářečské práce |
| Montáž | 3 | U | V | Montážní práce |
| Elektro | 5 | X | Y | Elektroinstalace |
| Balení | 8 | AA | AB | Balení a expedice |

**Sloupce:**
- **Původní termín**: Původní plánovaný termín z Helios
- **Plánovaný termín**: Upravený termín (editovatelný)

## Databázové závislosti

### View
- `hvw_TerminyZakazekProPlanovani` - Seznam zakázek s termíny

### Tabulky
- `TabZakazka` - Hlavní tabulka zakázek (READ)
- `TabZakazka_EXT` - Rozšíření zakázek (READ/WRITE)
  - Sloupce termínů: `_U1Start`, `_U1Konec`, `_U2Start`, `_U2Konec`, ...
  - Sloupce hodin: `_HodCelkem`, `_HodSkPrac1-5`, `_HodKoop`

### Stored Procedures
- `EP_PlanTerminyVyrobyDoZakazek` - Výpočet termínů výroby
  - Parametry: `@ID` (Long), `@DatumUkonceni` (Date/NULL)
  - Funkce: Zpětné plánování od data expedice

- `ep_DoplnTerminOperacePodleUseku` - Doplnění termínů operací
  - Parametry: `@ID` (Long)
  - Funkce: Přepočítá termíny jednotlivých operací podle úseků

## Zabezpečení

### Autentizace
- Stejný systém jako v Gantt aplikaci
- XOR šifrování přihlašovacích údajů
- Podpora NT autentizace (doporučeno)

### Ochrana listů
- List "Plan" je chráněn heslem: `MrkevNeniOvoce123`
- Ochrana umožňuje:
  - Filtrování dat
  - Úpravu určitých buněk (termíny)
  - Spouštění maker
- Ochrana zabraňuje:
  - Smazání vzorců
  - Změně struktury listu
  - Přesunutí řádků/sloupců

## Pokročilé funkce

### Automatická detekce fontu
Aplikace automaticky vybere nejlepší dostupný font v tomto pořadí:
1. Segoe UI Semilight
2. Segoe UI Light
3. Calibri Light
4. Arial Narrow
5. Arial (fallback)

Font musí:
- Být nainstalován v systému
- Podporovat českou diakritiku

### Formátování
- Automatické nastavení formátu datumu: `dd.mm.yy`
- Automatické přizpůsobení šířky sloupců
- Konzistentní formátování čísel
- Ochrana vzorců

### Logování
Aplikace obsahuje logování pro debugging:
```vba
WriteLog "zpráva"
```
Logy jsou zapisovány pro klíčové operace jako:
- Načítání zakázek
- Volání stored procedures
- Chyby při zpracování

## Známé limity

1. **Výkon**: S 500+ zakázkami může aktualizace trvat 30-60 sekund
2. **Font detection**: Vyžaduje Win32 API, funguje pouze na Windows
3. **Ochrana**: Heslo pro odemknutí listu je v kódu (bezpečnostní riziko)
4. **Časové prodlevy**: Některé operace mají hardcoded delay 1 sekunda (např. v Helios.bas)
5. **Concurrent access**: Aplikace nekontroluje, zda někdo jiný upravuje stejnou zakázku

## Řešení problémů

### Nepodařilo se připojit k databázi
- Zkontrolujte název serveru a databáze
- Ověřte přístupová práva
- Zkuste NT autentizaci (prázdné uživatelské jméno a heslo)

### Data se nenačítají
- Zkontrolujte, zda existuje view `hvw_TerminyZakazekProPlanovani`
- Ověřte SELECT oprávnění
- Zkontrolujte připojení k síti

### Chyba při výpočtu termínů
- Ověřte, že existuje stored procedura `EP_PlanTerminyVyrobyDoZakazek`
- Zkontrolujte EXECUTE oprávnění
- Ověřte, že zakázka existuje v `TabZakazka`

### Zkratky Ctrl+N, Ctrl+D nefungují
- Zkontrolujte, že jste ve sloupcích M, P, S, V, Y nebo AB
- Zkontrolujte, že jste na řádku >= 13
- Ověřte, že levá buňka obsahuje datum (původní termín)

### Formulář "Ready" bliká
- To je záměr, indikuje dokončení načítání
- Klikněte na "OK" pro zavření

### List je zamčený
- List "Plan" je chráněn heslem pro ochranu vzorců
- Pro odemknutí: Revize → Odemknout list → Heslo: `MrkevNeniOvoce123`
- **⚠️ Pozor:** Odemknutí může vést k nechtěnému smazání vzorců

## Changelog

Aplikace obsahuje changelog verzí v `frmChangelog`. Pro zobrazení:
1. Otevřete VBA Editor (Alt+F11)
2. Najděte `frmChangelog` v Project Explorer
3. Spusťte formulář (F5)

## Best Practices

### Doporučený workflow
1. **Ráno**: Otevřít aplikaci → Aktualizace dat z Helios
2. **Úprava**: Upravit termíny podle aktuální situace
3. **Kontrola**: Zkontrolovat obsazenost úseků
4. **Export**: Naplnit kontrolu a zkontrolovat změny
5. **Synchronizace**: Aktualizovat data v Helios
6. **Zavření**: Zavřít aplikaci (data se neukládají lokálně)

### Tipy pro efektivní práci
- Používejte filtry pro zobrazení relevantních zakázek
- Využívejte klávesové zkratky pro rychlé úpravy
- Pravidelně kontrolujte obsazenost úseků
- Před synchronizací do Helios vždy zkontrolujte změny v listu "Kontrola"

## Podpora

Pro technickou podporu nebo hlášení chyb kontaktujte:
- **Email**: vba-team@in-eko.cz
- **Internal**: #vba-planning channel

## Autor

IN-EKO VBA Development Team

## Licence

Internal use only - IN-EKO s.r.o.

## Verze

**Export:** 2026-01-16
**Velikost:** 1.9 MB
**Řádků kódu:** ~2,193
**Excel soubor:** Plánování.xlsm
