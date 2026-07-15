$base = 'C:\Users\D9352\OneDrive - 基恩士（中国）有限公司\IT-PartnerShare - 文档\71. AD日志\EventArchive'
$files = Get-ChildItem -Path $base -Recurse -File
foreach($file in $files) {
    Write-Host "=== $($file.FullName) ($($file.Length) bytes) ==="
}
Write-Host ""
Write-Host "Total files: $($files.Count)"
