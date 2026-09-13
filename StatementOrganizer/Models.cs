namespace StatementOrganizer;

/// <summary>
/// One input PDF file (may contain multiple statements) plus what the LLM
/// extracted from it.
/// </summary>
public class StatementFile
{
    public string FileName { get; set; } = string.Empty;

    /// <summary>Normalized category: "Fidelity" | "Vanguard" | "Bank" | "Other".</summary>
    public string Category { get; set; } = "Other";

    /// <summary>Exact institution name as printed on the statement (e.g. "Fidelity Investments").</summary>
    public string Institution { get; set; } = string.Empty;

    /// <summary>Kind of institution: "investment" | "bank" | "credit_card" | "unknown".</summary>
    public string InstitutionType { get; set; } = "unknown";

    /// <summary>Individual statements found in this file (a file may contain several).</summary>
    public List<Statement> Statements { get; set; } = new();

    /// <summary>Set to true if extraction failed.</summary>
    public bool Error { get; set; }

    public string? ErrorMessage { get; set; }
}

public class Statement
{
    /// <summary>Institution that issued this particular statement.</summary>
    public string Institution { get; set; } = string.Empty;

    /// <summary>Human-readable account name/type (e.g. "Roth IRA", "Checking ****1234").</summary>
    public string AccountName { get; set; } = string.Empty;

    /// <summary>Account / account-number fragment if present (last 4 digits preferred).</summary>
    public string? AccountNumber { get; set; }

    /// <summary>What kind of statement: e.g. "monthly investment statement", "credit card", "checking".</summary>
    public string StatementType { get; set; } = string.Empty;

    /// <summary>Opening balance (investment) / previous statement balance (bank).</summary>
    public decimal? OpeningBalance { get; set; }

    /// <summary>Ending / closing balance (investment) / current balance (bank).</summary>
    public decimal? ClosingBalance { get; set; }

    /// <summary>Best single "balance" figure to report (closing balance when available).</summary>
    public decimal? Balance { get; set; }

    /// <summary>Date the statement was issued ("statement date").</summary>
    public DateTime? StatementDate { get; set; }

    /// <summary>"As of" date the balances are valid for.</summary>
    public DateTime? AsOfDate { get; set; }

    /// <summary>Payment due date (bank / credit card only; null for investment statements).</summary>
    public DateTime? DueDate { get; set; }

    /// <summary>Amount due, if present (credit cards).</summary>
    public decimal? AmountDue { get; set; }

    public List<string> Notes { get; set; } = new();
}
