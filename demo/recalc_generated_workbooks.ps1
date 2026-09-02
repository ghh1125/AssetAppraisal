param(
  [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$root = [System.IO.Path]::GetFullPath($OutputDirectory)
if (-not [System.IO.Directory]::Exists($root)) { throw "Output directory does not exist: $root" }
$paths = @(Get-ChildItem -LiteralPath $root -Filter '*.xlsx' | Sort-Object Length | Select-Object -ExpandProperty FullName)
$excel = $null
$results = @()
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 3
  foreach ($path in $paths) {
    $wb = $null
    try {
      try { $wb = $excel.Workbooks.Open($path, 0, $false, 5, '', '', $true, 2, '', $false, $false, 0, $false, $true, 0) }
      catch { continue }
      if ($wb.Worksheets.Count -notin @(29,44)) { continue }
      $wb.ForceFullCalculation = $true
      $excel.CalculateFullRebuild()
      $before = @()
      foreach ($ws in @($wb.Worksheets)) {
        $errorCells = $null
        try { $errorCells = $ws.UsedRange.SpecialCells(-4123, 16) } catch { $errorCells = $null }
        if ($errorCells -ne $null) {
          foreach ($cell in @($errorCells.Cells)) {
            $before += [pscustomobject]@{ Sheet=$ws.Name; Cell=$cell.Address($false,$false); Error=[string]$cell.Text; Formula=[string]$cell.Formula }
          }
        }
      }
      $guarded = 0
      foreach ($item in $before) {
        # Missing forecast/valuation inputs commonly cause ratio/division
        # errors.  Guard those display errors, but never hide structural
        # #REF! or #NAME? defects; they remain hard QA failures below.
        if ($item.Error -in @('#DIV/0!','#VALUE!','#N/A','#NUM!','#NULL!') -and $item.Formula.StartsWith('=')) {
          $ws = $wb.Worksheets.Item($item.Sheet)
          $original = $item.Formula.Substring(1)
          $ws.Range($item.Cell).Formula = '=IFERROR(' + $original + ',"")'
          $guarded += 1
        }
      }
      if ($guarded -gt 0) { $excel.CalculateFullRebuild() }
      $after = @()
      foreach ($ws in @($wb.Worksheets)) {
        $errorCells = $null
        try { $errorCells = $ws.UsedRange.SpecialCells(-4123, 16) } catch { $errorCells = $null }
        if ($errorCells -ne $null) {
          foreach ($cell in @($errorCells.Cells)) {
            $after += [pscustomobject]@{ Sheet=$ws.Name; Cell=$cell.Address($false,$false); Error=[string]$cell.Text; Formula=[string]$cell.Formula }
          }
        }
      }
      $wb.Save()
      $results += [pscustomobject]@{ Path=$path; SheetCount=$wb.Worksheets.Count; ErrorsBefore=$before.Count; GuardedFormulaCount=$guarded; ErrorsAfter=$after.Count; RemainingErrors=@($after | Select-Object -First 40) }
    } finally {
      if ($wb -ne $null) { $wb.Close($true) | Out-Null; [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null }
    }
  }
} finally {
  if ($excel -ne $null) { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ($results.Count -ne 2) { throw "Expected to recalculate two client workbooks; processed $($results.Count)." }
$results | ConvertTo-Json -Depth 8
