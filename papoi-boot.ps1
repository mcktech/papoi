# Papoiongo - cargador (requiere PowerShell 7+)
# Solo lectura:  irm https://raw.githubusercontent.com/mcktech/papoi/main/papoi-boot.ps1 | iex
# Escritura:     definir antes $PapoiToken = 'github_pat_...' (en $PROFILE, nunca en el repo)
# Despues: papoi golpe | papoi -Add "golpe; pucopoi" | papoi -DestList | papoi -Help
function global:papoi {
    param(
        [string]$Add,
        [string]$Del,
        [string]$Edit,
        [Parameter(Position = 0)]
        [string]$Query,
        [switch]$DestList,
        [string]$Dest,
        [string]$File,
        [int]$TIMEOUT_S,
        [switch]$Help
    )
    $src = Invoke-RestMethod 'https://raw.githubusercontent.com/mcktech/papoi/main/papoi.ps1'
    $src = $src.TrimStart([char]0xFEFF)
    & ([scriptblock]::Create($src)) @PSBoundParameters
}
