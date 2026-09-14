# OneNoteUpdate.ps1 - late-bound (IDispatch) OneNote page writer.
# The .NET PIA early-bound UpdatePageContent overloads are misaligned on this
# machine (they marshal as the 4-arg form and return 0x80042001). PowerShell's
# own COM interop dispatches the 1-arg overload correctly. This helper reads a
# full page XML from a file and commits it via the 1-arg UpdatePageContent.
param(
    [Parameter(Mandatory=$true)][string]$PageId,
    [Parameter(Mandatory=$true)][string]$XmlFile
)

$ErrorActionPreference = "Stop"
try {
    $asm = [Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
    $app = New-Object $asm.GetType("Microsoft.Office.Interop.OneNote.ApplicationClass")
    $xml = [System.IO.File]::ReadAllText($XmlFile, [System.Text.Encoding]::UTF8)
    [void]$app.UpdatePageContent($xml)
    Write-Output "UPDATE_OK"
    exit 0
}
catch {
    Write-Output ("UPDATE_FAIL: " + $_.Exception.Message -replace "`r?n", " ")
    exit 1
}
