param(
    [string]$Path = "dotmngr.ps1"
)

$full = Resolve-Path -LiteralPath $Path
$code = Get-Content -Raw -LiteralPath $full

$tokens = $null
$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

if ($errors -ne $null -and $errors.Count -gt 0) {
    foreach ($e in $errors) {
        Write-Host "SYNTAX ERROR:" -ForegroundColor Red
        $e | Format-List * -Force
        Write-Host "----"
    }
    exit 1
}
else {
    Write-Host "Syntax OK" -ForegroundColor Green
}
