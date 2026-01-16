# Rozpocet – specifikace pro AppSheet (CZ)

Tento dokument převádí aktuální VBA aplikaci „Rozpocet“ do návrhu AppSheet řešení. Uvádí: analýzu současných funkcí, návrh datového modelu, bezpečnostní filtry, workflow a návrh UI/UX. Návrh respektuje AppSheet best practices: row‑based model, stabilní primární klíče, auditní sloupce a multi‑user přístup.

---

## 1) Shrnutí současné VBA aplikace

### 1.1 Primární zdroje dat

- **SQL Server view** `hvw_ReportRozpocetPlneni` → načítá se do listu **Rozpočet**. VBA používá `SELECT *` a řadí podle `Obdobi, Skupina, Ucet`.【F:VBA/Rozpocet/Modules/Data.bas†L20-L96】
- **SQL Server view** `hvw_ReportRozpocetPlneniDetail` → detailní drill‑down podle `SkupinaUctu, Rok, Mesic` a sloupců `Datum, Firma, Zamestnanec, Ucet, Nazev, CastkaMD, CastkaDAL, Popis`.【F:VBA/Rozpocet/Modules/Data.bas†L153-L246】

### 1.2 Pracovní listy a jejich role

- **Aplikace**: hlavní rozpočtová tabulka; sloupce B:AQ obsahují měsíční hodnoty v mřížce, řádky 4/5 nesou rok/měsíc; poznámky ve sloupci AS se ukládají do `data.ini`.【F:VBA/Rozpocet/Modules/Data.bas†L520-L534】【F:VBA/Rozpocet/Modules/Poznamky.bas†L14-L90】
- **Kumulace**: kumulace plán/skutečnost/rozdíl dle checkboxů; data se přepočítávají makrem a navazují na strukturu Aplikace. Rozsahy plán/skutečnost/rozdíl jsou fixní (F–Q, S–AD, AF–AQ).【F:VBA/Rozpocet/Modules/Data.bas†L285-L395】
- **DataProGraf**: pomocný list pro grafy, kopíruje data z Aplikace/Kumulace pro zobrazení ve formu grafů.【F:VBA/Rozpocet/Modules/Data.bas†L510-L534】
- **Grafy**: obsahuje grafy `GrafKategorie` a `GrafKategorieKumulativni`; grafy se exportují jako JPG a zobrazují ve formuláři `frmGraf`.【F:VBA/Rozpocet/Modules/Grafy.bas†L80-L139】
- **Kontingenční tabulka**: pivot „Rozpočet“ s možností drill‑down. 【F:VBA/Rozpocet/Modules/Navigace.bas†L27-L53】
- **Konfigurace**: uložení serveru a DB + barev grafů (C7/C8).【F:VBA/Rozpocet/Modules/Connection.bas†L80-L112】【F:VBA/Rozpocet/Modules/Grafy.bas†L5-L76】

### 1.3 Typický workflow v Excel/VBA

1. **Otevření sešitu** → skrytí Ribbon/toolbar, zobrazení `frmLogin`.【F:VBA/Rozpocet/ExcelObjects/ThisWorkbook.cls†L29-L69】
2. **Login** → ověření přes SQL Server (SQLOLEDB), uložení přihlašovacích údajů v paměti, načtení dat a konfigurace. 【F:VBA/Rozpocet/Forms/frmLogin.frm†L17-L67】【F:VBA/Rozpocet/Modules/Connection.bas†L22-L129】
3. **Načtení dat** → `LoadDataFromQueries` načte `hvw_ReportRozpocetPlneni` do listu Rozpočet, zobrazí progress form. 【F:VBA/Rozpocet/Modules/Data.bas†L20-L139】
4. **Kumulace** → checkboxy v Kumulaci určují, které řádky se kumulují; single‑select checkbox spouští grafy. 【F:VBA/Rozpocet/Modules/Data.bas†L256-L507】【F:VBA/Rozpocet/ExcelObjects/List3.cls†L12-L33】
5. **Detail** → klik do sloupců S–AD spustí detailní dotaz na `hvw_ReportRozpocetPlneniDetail` a vytvoří list Detail. 【F:VBA/Rozpocet/Modules/Data.bas†L153-L246】
6. **Export** → kopie do nového XLSX se sloupci B:AQ pro Aplikaci i Kumulaci (převod na hodnoty). 【F:VBA/Rozpocet/Modules/Copy.bas†L3-L99】

### 1.4 Současné omezení multi‑user

Sešit je fakticky **single‑user**: při otevření se pokouší přepnout na `xlReadWrite` a pokud je soubor otevřen jiným uživatelem, aplikace se ukončí. 【F:VBA/Rozpocet/ExcelObjects/ThisWorkbook.cls†L42-L61】

---

## 2) Identifikace klíčových entit a objektů

### 2.1 Entity vhodné pro samostatné tabulky

- **BudgetPeriods** (období rozpočtu)
- **CostCenters** (střediska / oddělení)
- **Categories** (kategorie rozpočtu)
- **BudgetLines** (řádkový rozpočet)
- **Actuals / Transactions** (skutečné transakce)
- **Users**
- **Roles** + **UserRoles**
- **UserCostCenters** (oprávnění k nákladovým střediskům)

### 2.2 Klíčové vstupy/výstupy

- **Input**: editace BudgetLines (plán), aktualizace master dat (Categories, CostCenters), správa BudgetPeriods.
- **Output**: přehledy plán vs skutečnost, kumulace, dashboardy a drill‑down detailů (Actuals).

---

## 3) Nový datový model (AppSheet backend)

### 3.1 Základní zásady

- Jeden objekt = jedna tabulka (row‑based data).
- Stabilní primární klíče (UUID/Text).
- Auditní sloupce (`CreatedAt`, `CreatedBy`, `UpdatedAt`, `UpdatedBy`).
- Multi‑user přístup s bezpečnostními filtry dle `UserEmail()`.

### 3.2 Přehled tabulek

#### 1) **Users**
**Purpose:** uživatelé aplikace.

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| UserId | Text | ✅ |  | UUID, unique |
| Email | Email |  |  | Unique, required |
| FullName | Text |  |  |  |
| IsActive | Yes/No |  |  | default TRUE |
| DefaultCostCenterId | Ref |  | CostCenters | optional |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

#### 2) **Roles**
**Purpose:** role‑based access.

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| RoleId | Text | ✅ |  | UUID |
| RoleName | Text |  |  | Unique (BudgetEditor/Approver/Viewer) |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

#### 3) **UserRoles**
**Purpose:** M:N vazba Users ↔ Roles.

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| UserRoleId | Text | ✅ |  | UUID |
| UserId | Ref |  | Users | required |
| RoleId | Ref |  | Roles | required |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

#### 4) **CostCenters**
**Purpose:** nákladová střediska / oddělení.

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| CostCenterId | Text | ✅ |  | UUID |
| CostCenterCode | Text |  |  | Unique |
| CostCenterName | Text |  |  | required |
| ManagerUserId | Ref |  | Users | optional |
| IsActive | Yes/No |  |  | default TRUE |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

#### 5) **Categories**
**Purpose:** rozpočtové kategorie (hierarchie volitelně).

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| CategoryId | Text | ✅ |  | UUID |
| CategoryCode | Text |  |  | Unique |
| CategoryName | Text |  |  | required |
| ParentCategoryId | Ref |  | Categories | optional |
| IsActive | Yes/No |  |  | default TRUE |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

#### 6) **BudgetPeriods**
**Purpose:** období rozpočtu (měsíce/kvartály/roky).

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| BudgetPeriodId | Text | ✅ |  | UUID |
| PeriodName | Text |  |  | e.g. 2025-01 |
| StartDate | Date |  |  | required |
| EndDate | Date |  |  | required |
| Status | Enum |  |  | Draft/Submitted/Approved/Locked |
| SourcePeriodId | Ref |  | BudgetPeriods | optional |
| AdjustmentPercent | Number |  |  | default 0 |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

#### 7) **BudgetLines**
**Purpose:** hlavní rozpočtový řádek (plán).

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| BudgetLineId | Text | ✅ |  | UUID |
| BudgetPeriodId | Ref |  | BudgetPeriods | required |
| CostCenterId | Ref |  | CostCenters | required |
| CategoryId | Ref |  | Categories | required |
| AccountCode | Text |  |  | optional |
| AmountPlanned | Number |  |  | required |
| Notes | LongText |  |  | optional |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

**Uniqueness constraint (AppSheet Valid_If):**
```
COUNT(
  SELECT(BudgetLines[BudgetLineId],
    AND(
      [BudgetPeriodId]=[_THISROW].[BudgetPeriodId],
      [CostCenterId]=[_THISROW].[CostCenterId],
      [CategoryId]=[_THISROW].[CategoryId],
      [AccountCode]=[_THISROW].[AccountCode]
    )
  )
)=1
```

#### 8) **Actuals**
**Purpose:** skutečné transakce (detail), ekvivalent VBA detailu.

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| ActualId | Text | ✅ |  | UUID |
| BudgetPeriodId | Ref |  | BudgetPeriods | required |
| CostCenterId | Ref |  | CostCenters | required |
| CategoryId | Ref |  | Categories | optional |
| AccountCode | Text |  |  | optional |
| TransactionDate | Date |  |  | required |
| Amount | Number |  |  | required |
| Company | Text |  |  | optional |
| Employee | Text |  |  | optional |
| Description | LongText |  |  | optional |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

#### 9) **UserCostCenters**
**Purpose:** oprávnění k datům dle střediska.

| Column | Type | Key | Ref | Constraints |
|---|---|---|---|---|
| UserCostCenterId | Text | ✅ |  | UUID |
| UserId | Ref |  | Users | required |
| CostCenterId | Ref |  | CostCenters | required |
| PermissionLevel | Enum |  |  | Viewer/Editor/Approver |
| CreatedAt | DateTime |  |  | audit |
| CreatedBy | Email |  |  | audit |
| UpdatedAt | DateTime |  |  | audit |
| UpdatedBy | Email |  |  | audit |

---

## 4) AppSheet backend konfigurace

### 4.1 Security Filters (hlavní tabulky)

**BudgetLines**:
```
IN(
  [CostCenterId],
  SELECT(UserCostCenters[CostCenterId],
    [UserId] = LOOKUP(USEREMAIL(), Users, Email, UserId)
  )
)
```

**Actuals**:
```
IN(
  [CostCenterId],
  SELECT(UserCostCenters[CostCenterId],
    [UserId] = LOOKUP(USEREMAIL(), Users, Email, UserId)
  )
)
```

**BudgetPeriods**:
```
IN(
  [BudgetPeriodId],
  SELECT(BudgetLines[BudgetPeriodId],
    IN([CostCenterId],
      SELECT(UserCostCenters[CostCenterId],
        [UserId] = LOOKUP(USEREMAIL(), Users, Email, UserId)
      )
    )
  )
)
```

### 4.2 Slices podle rolí a stavů

- **MyCostCenters_BudgetLines**: BudgetLines filtrované dle UserCostCenters.
- **DraftPeriods**: `BudgetPeriods[Status] = "Draft"`
- **SubmittedPeriods**: `BudgetPeriods[Status] = "Submitted"`
- **ApprovedPeriods**: `BudgetPeriods[Status] = "Approved"`
- **ToApprove** (Approver): SubmittedPeriods pouze pro Approver roli.

### 4.3 Agregace (virtuální sloupce)

**BudgetLines[AmountActual]** (virtuální):
```
SUM(
  SELECT(Actuals[Amount],
    AND(
      [BudgetPeriodId]=[_THISROW].[BudgetPeriodId],
      [CostCenterId]=[_THISROW].[CostCenterId],
      [CategoryId]=[_THISROW].[CategoryId]
    )
  )
)
```

**BudgetLines[Variance]**:
```
[AmountPlanned] - [AmountActual]
```

**Dashboard metrics** (např. per period):
```
SUM(SELECT(BudgetLines[AmountPlanned], [BudgetPeriodId]=[_THISROW].[BudgetPeriodId]))
```

### 4.4 Copy/Roll‑forward období (Action + Bot)

**Action: CreateNewPeriodFromPrevious** (na BudgetPeriods)
- Action type: `Data: add a new row to another table using values from this row`
- Target table: BudgetPeriods
- Values:
  - `PeriodName = CONCATENATE(TEXT([StartDate], "YYYY"), "-", TEXT([StartDate], "MM"))`
  - `StartDate = EDATE([StartDate], 1)`
  - `EndDate = EOMONTH([StartDate], 0)`
  - `Status = "Draft"`
  - `SourcePeriodId = [BudgetPeriodId]`
  - `AdjustmentPercent = 0`

**Bot: CopyBudgetLinesOnPeriodCreate**
- Event: `AddsOnly` na BudgetPeriods
- Process step: `Data: execute an action on a set of rows`
- Referenced rows:
  ```
  SELECT(BudgetLines[BudgetLineId], [BudgetPeriodId]=[_THISROW].[SourcePeriodId])
  ```
- Action on rows: **CopyBudgetLine**
  - Add row to BudgetLines with:
    - `BudgetPeriodId = [_THISROW-1].[BudgetPeriodId]` (nové období)
    - `CostCenterId = [CostCenterId]`
    - `CategoryId = [CategoryId]`
    - `AccountCode = [AccountCode]`
    - `AmountPlanned = [AmountPlanned] * (1 + [_THISROW-1].[AdjustmentPercent] / 100)`
    - `Notes = [Notes]`

### 4.5 Uzamčení období

**Editable_If** na BudgetLines:
```
[BudgetPeriodId].[Status] = "Draft"
```

**Approve action**:
```
SET [Status] = "Approved"
```

---

## 5) AppSheet UI/UX návrh

### 5.1 View list (typ + data source)

1) **HomeDashboard** (Dashboard)
   - Views: `MyBudgetLines` (deck), `PeriodKPI` (chart), `Approvals` (table)
   - Filter: aktuální období + cost center uživatele

2) **BudgetLines** (Table/Deck)
   - Slice: `MyCostCenters_BudgetLines`
   - Quick filters: PeriodName, CostCenter, Category

3) **BudgetLine Form** (Form)
   - Vstup pro `AmountPlanned`, `Notes`
   - `Editable_If` dle stavu period

4) **BudgetPeriods** (Table)
   - Admin/Approver only
   - Actions: CreateNewPeriodFromPrevious, Submit, Approve, Reject

5) **Actuals** (Table)
   - Slice podle uživatele
   - Drill‑down do detailu

6) **MasterData**
   - CostCenters (Table), Categories (Table) – Admin only

### 5.2 Role a přístup k view

- **BudgetEditor**: HomeDashboard, BudgetLines (Draft), Actuals.
- **Approver**: ToApprove + Approve/Reject akce, přístup k BudgetPeriods.
- **Viewer**: pouze ApprovedPeriods a souhrnné dashboardy.

---

## 6) Migrace VBA logiky do AppSheet

| VBA blok | Funkce | AppSheet implementace |
|---|---|---|
| `LoadDataFromQueries` | Načtení SQL view | Přímá tabulka AppSheet na SQL view (read‑only) + Sync | 
| `LoadAccountDetails` | Detailní drill‑down | Actuals tabulka + detail view | 
| `KumulujVysledovkuPodleCheckboxu` | Kumulace sloupců | Virtual columns + SUM přes BudgetLines (row‑based) | 
| `CopyDataToDataProGraf` + grafy | Grafy ve formuláři | Chart view v AppSheet dashboardu | 
| `SaveToINI` poznámky | Lokální poznámky | Sloupec Notes v BudgetLines | 

---

## 7) Doporučený postup migrace

1. **Vytvořit backend tabulky** v SQL/Google Sheets/AppSheet DB dle návrhu.
2. **Importovat master data** (CostCenters, Categories, Users).
3. **Migrovat rozpočet** z Excelu do BudgetLines (row‑based) – rozpad mřížky B:AQ na řádky.
4. **Napojit Actuals** z ERP (SQL view nebo ETL do tabulky Actuals).
5. **Nastavit security filters** a role.
6. **Vytvořit UI/UX** dle návrhu a otestovat workflow: Draft → Submitted → Approved.

---

## 8) Poznámky k rozdílům oproti VBA

- V AppSheet není nutné exportovat grafy do JPG – dashboardy mají nativní grafy.
- Row‑based data odstraní limity mřížkové struktury a zjednoduší agregace.
- Multi‑user přístup bude řešen přes bezpečnostní filtry a role (bez omezení na single‑user sešit).
