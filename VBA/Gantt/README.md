# Gantt - Vizualizace plánování výroby

## Přehled

Gantt je VBA aplikace v Microsoft Excel pro vizualizaci plánování výroby v IN-EKO. Aplikace zobrazuje časovou osu zakázek, kontroluje kapacity výrobních středisek (Příprava, Svařování, Montáž, Elektro) a upozorňuje na překročení kapacit pomocí barevného kódování.

## Funkce

- **Autentizace**: Zabezpečené připojení k SQL Server databázi s podporou NT autentizace i SQL autentizace
- **Načítání zakázek**: Automatické načtení zakázek z databázového view `hvw_TerminyZakazekProPlanovani`
- **Gantt diagram**: Vizualizace termínů jednotlivých fází výroby pro každou zakázku
- **Kontrola kapacit**: Automatické počítání obsazenosti výrobních středisek po dnech
- **Barevné kódování**: Zelená (volno), oranžová (plná kapacita), červená (překročení)
- **Export dat**: Podpora exportu VBA kódu do Git

## Požadavky

- Microsoft Excel 2010 nebo novější
- Připojení k SQL Server databázi IN-EKO ERP
- Oprávnění pro čtení z view `hvw_TerminyZakazekProPlanovani`
- ADODB (ActiveX Data Objects) knihovna

## Instalace

1. Otevřete Excel soubor s makry (`.xlsm`)
2. Ujistěte se, že máte povolená makra (Soubor → Možnosti → Centrum zabezpečení)
3. Nakonfigurujte připojení na listu "Konfigurace":
   - `serverName`: Název SQL serveru
   - `databaseName`: Název databáze
   - `login`: Uživatelské jméno (volitelné, pro NT autentizaci ponechte prázdné)

## Použití

### První spuštění

1. Při otevření souboru se zobrazí přihlašovací formulář
2. Vyplňte přihlašovací údaje:
   - **Server**: Název SQL serveru
   - **Databáze**: Název databáze
   - **Uživatel**: Uživatelské jméno (nebo prázdné pro NT autentizaci)
   - **Heslo**: Heslo (nebo prázdné pro NT autentizaci)
3. Klikněte na "Přihlásit"

### Načtení dat

**Ruční načtení:**
- Spusťte makro `LoadOrUpdateData` pro načtení dat ze serveru do listu "Zakazky"
- Spusťte makro `AktualizovatSeznamZakazek` pro aktualizaci Gantt diagramu

**Aktualizace kapacit:**
- Spusťte makro `SumarizaceBoduProVsechnySloupce` pro přepočet obsazenosti středisek

### Interpretace barev

Na konci Gantt diagramu se zobrazují řádky s kontrolou kapacit:

| Barva | Význam |
|-------|--------|
| 🟢 Zelená | Kapacita je volná (méně než maximum) |
| 🟠 Oranžová | Kapacita je na 100% |
| 🔴 Červená | Kapacita je překročena |
| ⚪ Bílá | Víkend nebo svátek / žádné zakázky |

### Výrobní střediska

Aplikace sleduje 4 výrobní střediska:

1. **Příprava** - 1 pracovník
2. **Svařování** - 2 pracovníci
3. **Montáž** - 2 pracovníci
4. **Elektro** - 1 pracovník

## Struktura projektu

```
VBA/Gantt/
├── Modules/
│   ├── Connection.bas      # Správa databázového připojení a autentizace
│   ├── Data.bas            # Načítání a aktualizace dat ze serveru
│   ├── Advanced.bas        # Pokročilé funkce (kontrola kapacit, formátování)
│   └── ExportToGit.bas     # Export VBA kódu do Git
├── Forms/
│   ├── frmLogin.frm        # Přihlašovací formulář
│   └── frmProgress.frm     # Ukazatel průběhu operací
└── ExcelObjects/
    ├── ThisWorkbook.cls    # Události workbooku (Workbook_Open, BeforeClose)
    └── List*.cls           # Třídy jednotlivých listů
```

## Databázová závislost

Aplikace vyžaduje přístup k následujícím databázovým objektům:

- **View**: `hvw_TerminyZakazekProPlanovani`
  - Obsahuje seznam zakázek s termíny jednotlivých fází výroby
  - Sloupce: Zakázka, DatumZahajeni, DatumUkonceni, a termíny pro jednotlivé fáze

## Bezpečnost

- Přihlašovací údaje jsou šifrovány pomocí XOR šifrování s klíčem `ENCRYPTION_KEY`
- Údaje jsou uloženy pouze v paměti po dobu běhu aplikace
- Aplikace podporuje Windows NT autentizaci (doporučeno)

## Známé limity

- Aplikace předpokládá konkrétní strukturu listů (Gantt, Zakazky, Konfigurace, Svátky)
- Kapacity výrobních středisek jsou napevno definovány v kódu (Advanced.bas:66-69)
- Minimalistický režim (bez mřížky a záhlaví) se aktivuje automaticky při otevření

## Řešení problémů

### Nepodařilo se připojit k databázi
- Zkontrolujte název serveru a databáze v přihlašovacím formuláři
- Ověřte síťové připojení k SQL serveru
- Ujistěte se, že máte přístupová práva k databázi

### Data se nenačítají
- Zkontrolujte, zda existuje view `hvw_TerminyZakazekProPlanovani`
- Ověřte, že máte oprávnění SELECT na tento view
- Zkontrolujte, zda je list "Zakazky" přítomen v sešitu

### Soubor je již otevřen jiným uživatelem
- Aplikace kontroluje, zda není soubor otevřen jiným uživatelem
- Pokud je soubor otevřen vámi na jiném počítači, zavřete jej tam

## Podpora

Pro technickou podporu nebo hlášení chyb kontaktujte tým vývoje IN-EKO ERP.

## Autor

IN-EKO VBA Development Team

## Verze

Export: 2026-01-16
