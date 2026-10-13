$ErrorActionPreference="Stop"
$env:GOOGLE_APPLICATION_CREDENTIALS='D:\SCARLETT_512\SCARLETT-329\SOM\Data\Sent Email\sales-som datawarehouse 490008.json'
$env:PYTHONUTF8="1"
$py=".\.venv\Scripts\python.exe"
function Step($n,[scriptblock]$b){ "=== $(Get-Date -f HH:mm:ss) $n"; & $b; if($LASTEXITCODE -ne 0){ "FAILED: $n exit $LASTEXITCODE"; exit 1 } }
Step "download_inputs" { & $py download_inputs.py }
Step "download_history" { & $py download_history.py }
Step "build workbooks" { & $py -u sales_monthly_report.py --month 2026-07 --month 2026-08 --month prev --month current --no-backup --anchanto-scope quarter }
Step "workbooks->parquet" { $x=Get-ChildItem work\output\*\*.xlsx | % FullName; & $py -u build_primary_parquet.py month @x }
Step "combine" { & $py -u build_primary_parquet.py combine work/monthly_parquet work/Primary_Sales.parquet }
Step "upload_outputs" { & $py upload_outputs.py }
Step "promote dry-run" { & $py -u promote_to_production.py }
"=== BUILT_OK $(Get-Date -f HH:mm:ss)"
