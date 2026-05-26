// ============================================================
// FILE: reconciliation.m
// PROJECT: Semi-Automated Treasury Reconciliation
// AUTHOR: Wanlaya Sirirojanan
// LAST UPDATED: May 2026
//
// PURPOSE:
// Power Query (M Code) steps for parsing CAMT.053 XML,
// joining with GL data, reconciling transactions, and
// preparing the consolidated table for Power BI.
//
// HOW TO USE:
// Open Excel → Power Query Editor → Advanced Editor
// Copy and paste each section into the relevant query step.
// ============================================================


// ------------------------------------------------------------
// SECTION 1: FOLDER PARAMETERS
// Set folder paths before running. Update these paths to match
// your local folder structure.
// ------------------------------------------------------------

// BankStatementFolder parameter
"C:\TreasuryRecon\BankStatement"

// GLStatementFolder parameter
"C:\TreasuryRecon\GLStatement"


// ------------------------------------------------------------
// SECTION 2: LOAD & PARSE CAMT.053 XML (Bank Statement)
// Load the latest XML file from the BankStatement folder,
// expand all relevant nodes, and clean the raw data.
// ------------------------------------------------------------

let
    // Load all files from BankStatement folder
    Source = Folder.Files(BankStatementFolder),

    // Keep only .xml files
    FilterXML = Table.SelectRows(Source, each [Extension] = ".xml"),

    // Take the most recently modified file
    LatestFile = Table.FirstN(
        Table.Sort(FilterXML, {{"Date modified", Order.Descending}}), 1
    ),

    // Load XML content
    FilePath = LatestFile{0}[Folder Path] & LatestFile{0}[Name],
    XmlContent = Xml.Tables(File.Contents(FilePath)),

    // Navigate to transaction entries (Ntry nodes)
    BkToCstmrStmt = XmlContent{0}[BkToCstmrStmt],
    Stmt = BkToCstmrStmt{0}[Stmt],
    Ntry = Stmt{0}[Ntry],

    // Expand core transaction fields
    ExpandNtry = Table.ExpandTableColumn(Ntry, "NtryDtls",
        {"TxDtls"}, {"TxDtls"}),
    ExpandTxDtls = Table.ExpandTableColumn(ExpandNtry, "TxDtls",
        {"Refs", "Amt", "CdtDbtInd", "RltdPties", "RmtInf"},
        {"Refs", "TxAmt", "TxCdtDbtInd", "RltdPties", "RmtInf"}),

    // Extract EndToEndId (primary reference key)
    ExpandRefs = Table.ExpandTableColumn(ExpandTxDtls, "Refs",
        {"EndToEndId"}, {"EndToEndId"}),

    // Extract remittance information (description)
    ExpandRmtInf = Table.ExpandTableColumn(ExpandRefs, "RmtInf",
        {"Ustrd"}, {"Description"}),

    // Extract bank transaction code (Cd and SubFmlyCd)
    ExpandBkTxCd = Table.ExpandTableColumn(ExpandRmtInf, "BkTxCd",
        {"Domn"}, {"Domn"}),
    ExpandDomn = Table.ExpandTableColumn(ExpandBkTxCd, "Domn",
        {"Cd", "Fmly"}, {"Cd", "Fmly"}),
    ExpandFmly = Table.ExpandTableColumn(ExpandDomn, "Fmly",
        {"SubFmlyCd"}, {"SubFmlyCd"}),

    // Rename and select required columns
    RenameColumns = Table.RenameColumns(ExpandFmly, {
        {"NtryRef", "BankRef"},
        {"BookgDt", "BookingDate"},
        {"Amt", "Amount"},
        {"CdtDbtInd", "CdtDbtInd"}
    }),

    // Convert amount to number
    ChangeType = Table.TransformColumnTypes(RenameColumns, {
        {"Amount", type number},
        {"BookingDate", type date}
    }),

    // Create signed amount: credit = positive, debit = negative
    SignedAmount = Table.AddColumn(ChangeType, "Signed_Amount_Bank",
        each if [CdtDbtInd] = "CRDT" then [Amount] else -[Amount],
        type number),

    // Extract period from booking date (YYYY-MM format)
    AddPeriod = Table.AddColumn(SignedAmount, "Bank_Period",
        each Date.ToText([BookingDate], "yyyy-MM"), type text),

    // Extract key reference from EndToEndId for fuzzy matching
    AddRefKey = Table.AddColumn(AddPeriod, "Bank_Ref_Key",
        each if [EndToEndId] = null then null
             else Text.Upper(Text.Replace(
                 Text.Select([EndToEndId],
                     {"A".."Z","a".."z","0".."9"}),
                 " ", "")),
        type text)

in
    AddRefKey


// ------------------------------------------------------------
// SECTION 3: LOAD GL DATA
// Load the latest CSV file from the GLStatement folder.
// ------------------------------------------------------------

let
    Source = Folder.Files(GLStatementFolder),
    FilterCSV = Table.SelectRows(Source, each [Extension] = ".csv"),
    LatestFile = Table.FirstN(
        Table.Sort(FilterCSV, {{"Date modified", Order.Descending}}), 1
    ),
    FilePath = LatestFile{0}[Folder Path] & LatestFile{0}[Name],

    // Load CSV
    GLData = Csv.Document(File.Contents(FilePath),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]),

    // Promote headers
    PromoteHeaders = Table.PromoteHeaders(GLData,
        [PromoteAllScalars=true]),

    // Set column types
    ChangeTypes = Table.TransformColumnTypes(PromoteHeaders, {
        {"posted_date", type date},
        {"doc_ref", type text},
        {"description", type text},
        {"paid_out", type number},
        {"received_in", type number},
        {"gl_category", type text}
    }),

    // Create signed amount for GL side
    GLSignedAmount = Table.AddColumn(ChangeTypes, "GL_Signed_Amount",
        each if [received_in] > 0 then [received_in] else -[paid_out],
        type number),

    // Create GL period column
    AddGLPeriod = Table.AddColumn(GLSignedAmount, "GL_Period",
        each Date.ToText([posted_date], "yyyy-MM"), type text),

    // Create GL reference key for fuzzy matching
    AddGLRefKey = Table.AddColumn(AddGLPeriod, "GL_Ref_Key",
        each if [doc_ref] = null then null
             else Text.Upper(Text.Replace(
                 Text.Select([doc_ref],
                     {"A".."Z","a".."z","0".."9"}),
                 " ", "")),
        type text)

in
    AddGLRefKey


// ------------------------------------------------------------
// SECTION 4: CONSOLIDATION — FULL OUTER JOIN
// Merge bank transactions and GL entries.
// Full outer join preserves all rows from both sides,
// so unmatched items appear in the output for review.
// ------------------------------------------------------------

let
    Source = Table.NestedJoin(
        BankTransaction,        // Left table: bank side
        {"EndToEndId"},         // Left key: bank reference
        GL,                     // Right table: GL side
        {"doc_ref"},            // Right key: GL reference
        "GL",                   // Name for expanded GL columns
        JoinKind.FullOuter      // Keep all rows from both sides
    ),

    // Expand GL columns into the consolidated table
    ExpandGL = Table.ExpandTableColumn(Source, "GL", {
        "doc_ref", "description", "GL_Signed_Amount",
        "GL_Period", "GL_Ref_Key", "gl_category"
    }, {
        "GL.doc_ref", "GL.Description", "GL.Signed_Amount",
        "GL.Period", "GL.GL_Ref_Key", "GL.gl_category"
    })

in
    ExpandGL


// ------------------------------------------------------------
// SECTION 5: RECONCILIATION STATUS COLUMN
// Two-pass matching logic.
// Pass 1: Full reference match (EndToEndId = doc_ref)
// Pass 2: Key reference match (Bank_Ref_Key = GL_Ref_Key)
// Remaining items flagged by responsible party.
// ------------------------------------------------------------

Table.AddColumn(Consolidation, "Recon_Status",
    each
        if [EndToEndId] <> null and [GL.doc_ref] <> null
            and [EndToEndId] = [GL.doc_ref]
        then "Done.Full Reference"

        else if [Bank_Ref_Key] <> null and [GL.GL_Ref_Key] <> null
            and [Bank_Ref_Key] = [GL.GL_Ref_Key]
        then "Done.Key Reference"

        else if [EndToEndId] = null and [GL.doc_ref] <> null
        then "Pending process by BANK"

        else "Pending process by GL",
    type text)


// ------------------------------------------------------------
// SECTION 6: RECONCILIATION NOTE COLUMN
// Human-readable note for dashboard display.
// ------------------------------------------------------------

Table.AddColumn(#"Recon_Status", "Recon_Note",
    each
        if [Recon_Status] = "Done.Full Reference"
            or [Recon_Status] = "Done.Key Reference"
        then "Match"
        else "Unmatch, Review required.",
    type text)


// ------------------------------------------------------------
// SECTION 7: CATEGORY CLASSIFICATION COLUMNS
// Cat1 — Revenue vs Expense
// Cat2 — Operating vs Financing
// Cat3 — Transaction detail category
// ------------------------------------------------------------

// Cat1: Revenue / Expense
Table.AddColumn(#"Recon_Note", "Cat1_RevExp",
    each
        if [CdtDbtInd] = "CRDT" then "1_Revenue"
        else "2_Expense",
    type text)


// Cat2: Operating / Financing
Table.AddColumn(#"Cat1_RevExp", "Cat2_OpFin",
    each
        if [Cd] = "PMNT" and [SubFmlyCd] <> "LBOX"
            and [SubFmlyCd] <> "SAVR"                  then "Operating"
        else if [Cd] = "PMNT" and [SubFmlyCd] = "LBOX" then "Financing"
        else if [Cd] = "ACMT" and [SubFmlyCd] = "LBOX" then "Financing"
        else if [Cd] = "PMNT" and [SubFmlyCd] = "SAVR" then "Financing"
        else if [Cd] = "ACMT" and [SubFmlyCd] = "SAVR" then "Financing"
        else if [Cd] = "ACMT" and [SubFmlyCd] = "CHRG" then "Operating"
        else if [Cd] = "PMNT" and [SubFmlyCd] = "CHRG" then "Operating"
        else if [Cd] = "ACMT" and [SubFmlyCd] = "INTR" then "Financing"
        else if [Cd] = "PMNT" and [SubFmlyCd] = "INTR" then "Financing"
        else "Other",
    type text)


// Cat3: Transaction detail category
Table.AddColumn(#"Cat2_OpFin", "Category",
    each
        if [SubFmlyCd] = "ESCT" and [CdtDbtInd] = "CRDT"
            then "Domestic Revenue"
        else if [SubFmlyCd] = "XBCT" and [CdtDbtInd] = "CRDT"
            then "Oversea Revenue"
        else if [SubFmlyCd] = "ESCT" and [CdtDbtInd] = "DBIT"
            then "Domestic Supplier Payment"
        else if [SubFmlyCd] = "XBCT" and [CdtDbtInd] = "DBIT"
            then "Oversea Supplier Payment"
        else if [SubFmlyCd] = "FCEX" and [CdtDbtInd] = "CRDT"
            then "FX Selling"
        else if [SubFmlyCd] = "INTR" and [Signed_Amount_Bank] > 0
            then "Interest Income"
        else if [SubFmlyCd] = "INTR" and [Signed_Amount_Bank] < 0
            then "Interest Expense"
        else if [SubFmlyCd] = "CHRG"
            then "Bank Charges"
        else if [SubFmlyCd] = "SALA"
            then "Salary & Employee Benefit"
        else if [SubFmlyCd] = "TAXS" and [KeyMat_1] = "VAT"
            then "VAT"
        else if [SubFmlyCd] = "TAXS" and [KeyMat_1] = "CUS"
            then "Import Duty Tax"
        else if [SubFmlyCd] = "LBOX" and [Signed_Amount_Bank] < 0
            then "Loan-Repayment"
        else if [SubFmlyCd] = "LBOX" and [Signed_Amount_Bank] > 0
            then "Loan-Drawdown"
        else if [SubFmlyCd] = "SAVR" and [Signed_Amount_Bank] < 0
            then "TD-Deposit"
        else if [SubFmlyCd] = "SAVR" and [Signed_Amount_Bank] > 0
            then "TD-Drawdown"
        else "Other - Please Review",
    type text)


// ------------------------------------------------------------
// SECTION 8: RUNNING BALANCE
// Row index is added after sorting by date.
// Running balance accumulates signed amounts from the
// period opening balance — dynamically per period.
//
// NOTE: Uses [Bank_Period] dynamically so that May, June,
// and future months each pick up their own opening balance.
// Earlier versions hardcoded "2026-04" here — that was a bug.
// ------------------------------------------------------------

// Step 1: Sort by booking date ascending
#"Sorted" = Table.Sort(#"Category",
    {{"BookingDate", Order.Ascending}}),

// Step 2: Add row index (0-based) after sort
#"Indexed" = Table.AddIndexColumn(#"Sorted", "RowIndex", 0, 1, Int64.Type),

// Step 3: Calculate running balance
#"RunningBalance" = Table.AddColumn(#"Indexed", "RunningBalance",
    each
        let
            // Dynamically get the period of the current row
            CurrentPeriod = [Bank_Period],

            // Look up opening balance for that period from HeaderMonthend
            BB = List.First(
                Table.SelectRows(
                    HeaderMonthend,
                    each [Monthend_Period] = CurrentPeriod
                    and [Balance_Period] = "Beginning Balance"
                )[#"Bal.Amt.Element:Text"]
            ),
            BBValue = if BB = null then 0 else BB
        in
            BBValue + List.Sum(
                List.Range(
                    #"Indexed"[Signed_Amount_Bank],
                    0,
                    [RowIndex] + 1
                )
            )
, type number)


// ------------------------------------------------------------
// SECTION 9: WEEK LABEL
// Weekly count starts Monday–Sunday within each month.
// W1 = first week of the month, regardless of calendar week.
// ------------------------------------------------------------

Table.AddColumn(#"RunningBalance", "WeekLabel",
    each "W" & Text.From(
        Number.RoundUp(
            (Date.Day([BookingDate]) +
             Date.DayOfWeek(
                 Date.StartOfMonth([BookingDate]),
                 Day.Monday
             )
            ) / 7
        )
    ),
    type text)

