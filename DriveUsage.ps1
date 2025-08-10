param(
    [string]$Path = "C:\",
    [int]$Depth = 1  # Increase for more detail, decrease for speed
)

function Get-FolderSize {
    param($Folder)

    $size = (Get-ChildItem -LiteralPath $Folder.FullName -Recurse -File -ErrorAction SilentlyContinue |
             Measure-Object -Property Length -Sum).Sum
    [PSCustomObject]@{
        Name   = $Folder.Name
        Path   = $Folder.FullName
        SizeGB = "{0:N2}" -f ($size / 1GB)
    }
}

Write-Host "Scanning $Path (Depth: $Depth) ..." -ForegroundColor Cyan

$folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue -Depth $Depth

$results = foreach ($folder in $folders) {
    Get-FolderSize -Folder $folder
}

$results | Sort-Object {[double]$_.SizeGB} -Descending | Format-Table -AutoSize
