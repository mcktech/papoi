#requires -Version 7
<#
.SYNOPSIS
    Consultas y ediciones al diccionario Papoiongo (papoi.md). Un solo comando por vez.
.EXAMPLE
    .\papoi-edit.ps1 -Add "golpe; pucopoi"
    .\papoi-edit.ps1 -Add "golpe=pucopoi / pucapoi"
    .\papoi-edit.ps1 -Add            (sin texto: pide los campos uno a uno)
    .\papoi-edit.ps1 -Add "golpe"    (palabra sola: 1ra columna dada, pide el resto)
    .\papoi-edit.ps1 -Query "golpe"
    .\papoi-edit.ps1 -Edit "pucopoi" -Dest ejemplos
    .\papoi-edit.ps1 -DestList
#>
param(
    [switch]$Add,
    [string]$Del,
    [string]$Edit,
    [Parameter(Position = 0)]
    [string]$Query,
    [switch]$DestList,
    [string]$Dest = 'Vocabulario',
    [string]$File,
    [int]$TIMEOUT_S = 60,
    [switch]$Help
)

# =====================================================================
#  Validos: Black DarkBlue DarkGreen DarkCyan DarkRed DarkMagenta DarkYellow
#           Gray DarkGray Blue Green Cyan Red Magenta Yellow White
# =====================================================================

# --- Generales ---
$ColTitulo         = 'Cyan'      # titulos: nombre de seccion, encabezados de pantalla, resumen final
$ColTexto          = 'White'     # texto destacado: registro a agregar, filas listadas, ejemplo de formato
$ColSecundario     = 'DarkGray'  # info secundaria: conteos, "Omitido", "Identico", ruta del backup, n de linea en el editor
$ColError          = 'Red'       # errores (destino invalido, formato, backup fallido, etc.)
$ColAviso          = 'Yellow'    # avisos: "Sin coincidencias", "--Coincidencias:", "El registro ya existe", "Sin cambios"
$ColCancelado      = 'Red'       # "Cancelado."
$ColOk             = 'Green'     # "Guardado: ..."

# --- Encabezado ---
$ColBanner         = 'DarkMagenta'   # recuadro de asteriscos con el titulo del script

# --- Preguntas y teclas (Enter=si  Otra=no (60s)) ---
$ColPregunta       = 'Yellow'    # texto de la pregunta (Sobreescribir?, Borrar?, Numero(s)...)
$ColTeclaPrompt    = 'Red'       # nombre de la tecla: "Enter", "Otra"
$ColAccionPrompt   = 'Cyan'      # accion de la tecla: "=si", "=no", "Arriba=volver"
$ColTimeout        = 'DarkGray'  # "(60s)" y "(sin respuesta)"
$ColTeclaEco       = 'DarkYellow'     # eco de la tecla/texto que tipea el usuario

# --- Diff estilo linux ---
$ColDiffQuita      = 'Red'       # lineas "- ..."
$ColDiffAgrega     = 'Green'     # lineas "+ ..."

# --- Listas numeradas y coincidencias ---
$ColListaNum       = 'Cyan'      # "[1]" en la lista para elegir
$ColListaLinea     = 'DarkGray'  # "(l.25)" numero de linea en el archivo
$ColListaFila      = 'White'     # fila en la lista para elegir
$ColCoincidencia   = 'Gray'      # fila listada en "--Coincidencias:" del -Add

# --- Ayuda (-Help) ---
$ColAyudaSubtitulo = 'White'     # descripcion junto al nombre del script
$ColAyudaUso       = 'Yellow'    # linea "Uso: ..."
$ColAyudaParam     = 'Green'     # nombre del parametro
$ColAyudaTexto     = 'Gray'      # descripciones y notas

# --- Editor interactivo (-Edit) ---
$ColEditAyuda      = 'DarkGray'  # linea de teclas (Enter=acepta ...)
$ColEditEtiqueta   = 'Gray'      # nombre de columnas ya completadas
$ColEditValor      = 'Green'     # valores ya completados
$ColEditCampo      = 'Yellow'    # "[2/3] Termino:" campo actual
$ColEditActual     = 'DarkGreen' # "actual: ..." (valor original)
$ColEditPrompt     = 'Green'     # "editar:"
$ColEditBuffer     = 'White'     # texto que se esta editando

# --- Resultado de -Query ---
$ColQueryEsp       = 'Magenta'      # termino en español
$ColQuerySep       = 'Gray'  # el " = "
$ColQueryPapoi     = 'Green'     # termino papoiongo y resto de las columnas (3ra en adelante: origen, formateado/short, etc.)

# --- Resultado de -DestList ---
$ColDestFormato    = 'Gray'      # "formato: ..." (encabezado de la seccion)
$ColDestEjemplo    = 'White'     # "ejemplo: ..." (primera fila)

try { [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false) } catch { }
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------- Origen de datos ----------
# Sin -File trabaja contra GitHub (lectura anonima; escritura con el token en $PapoiToken o $env:PAPOI_TOKEN). Con -File: archivo local.
$GhOwner = 'mcktech'; $GhRepo = 'papoi'; $GhPath = 'papoi.md'; $GhBranch = 'main'

# ---------- Encabezado ----------
$bannerTxt = '****  Editor de Papoiongo (by Mck)  ****'
$bannerBar = '*' * $bannerTxt.Length
Write-Host ''
Write-Host $bannerBar -ForegroundColor $ColBanner
Write-Host $bannerTxt -ForegroundColor $ColBanner
Write-Host $bannerBar -ForegroundColor $ColBanner
Write-Host ''

# ---------- Ayuda ----------
function Show-Help {
    function Hl($p, $d) { Write-Host ('  {0,-18}' -f $p) -ForegroundColor $ColAyudaParam -NoNewline; Write-Host $d -ForegroundColor $ColAyudaTexto }
    Write-Host ''
    Write-Host 'papoi-edit.ps1' -ForegroundColor $ColTitulo -NoNewline
    Write-Host ' - consulta / edita papoi.md (un comando por vez)' -ForegroundColor $ColAyudaSubtitulo
    Write-Host 'Uso: .\papoi-edit.ps1 <comando> "texto" [-Dest <seccion>]' -ForegroundColor $ColAyudaUso
    Write-Host ''
    Hl '-Add "esp; papoi"'   'Agrega al final de la seccion. Si hay coincidencias las lista; si el registro existe ofrece sobrescribir (diff + / -)'
    Hl ''                    'Acepta ";" "=" o ":" como separador. 3er dato opcional ("esp; papoi; origen"); los datos vacios al final no se guardan'
    Hl ''                    '"golpe=pucopoi / pucapoi" crea 2 registros (se divide el 2do campo por "/")'
    Hl '-Add'                'Sin texto: pide los campos uno a uno (como -Edit) y chequea si ya existe cada dato al confirmarlo'
    Hl '-Add "palabra"'      'Palabra sola (sin ; = :): la toma como 1ra columna, la chequea y pide el resto desde la 2da'
    Hl '-Del "texto"'        'Busca (esp o papoi, parcial) y borra la linea completa, con confirmacion'
    Hl '-Edit "texto"'       'Busca (esp o papoi, parcial) y edita los campos interactivamente'
    Hl '-Query "texto"'      'Busca en TODAS las secciones (o solo en -Dest si se indica) y lista <español> = <papoiongo>. Atajo: papoi-edit texto'
    Hl '-DestList'           'Lista las secciones detectadas y un ejemplo de formato'
    Hl '-Dest <seccion>'     'Seccion destino (default: Vocabulario). Case-insensitive, alcanza el comienzo del nombre'
    Hl ''                    'Al editar/borrar un termino papoi busca sus apariciones en otras secciones: -Edit ofrece cambiarlas (tambien las fusionadas) y -Del ofrece marcarlas [BORRADO]; siempre con diff'
    Hl '-File <ruta>'        "Archivo local. Sin -File se usa GitHub ($GhOwner/$GhRepo); sin token es solo lectura"
    Write-Host ''
    Write-Host 'Seccion = "## Nombre". Su 1er renglon no vacio es el encabezado: define separador (; = o :) y columnas. Las filas siguen hasta una linea en blanco.' -ForegroundColor $ColAyudaTexto
    Write-Host "Confirmaciones: Enter = si, otra tecla = no, sin respuesta en ${TIMEOUT_S}s = no." -ForegroundColor $ColAyudaTexto
    Write-Host 'Modo local: antes de cada escritura se guarda copia en <carpeta del md>\bk\papoi_bk.zip. Modo GitHub: cada guardado es un commit.' -ForegroundColor $ColAyudaTexto
    Write-Host ''
}

$cmds = @()
if ($Add.IsPresent)                           { $cmds += 'Add' }
if ($PSBoundParameters.ContainsKey('Del'))    { $cmds += 'Del' }
if ($PSBoundParameters.ContainsKey('Edit'))   { $cmds += 'Edit' }
if ($PSBoundParameters.ContainsKey('Query') -and -not $Add.IsPresent) { $cmds += 'Query' }
if ($DestList.IsPresent)                      { $cmds += 'DestList' }

if ($Help -or $cmds.Count -eq 0) { Show-Help; return }
if ($cmds.Count -gt 1) { Write-Host "Solo un comando a la vez (se indicaron: $($cmds -join ', '))." -ForegroundColor $ColError; return }
$cmd = $cmds[0]
if ($cmd -ne 'DestList') {
    # -Add es switch: el texto (si hay) llega por posicion en $Query. Sin texto = modo interactivo
    $arg = switch ($cmd) { 'Add' { $Query } 'Del' { $Del } 'Edit' { $Edit } 'Query' { $Query } }
    if ($cmd -ne 'Add' -and [string]::IsNullOrWhiteSpace($arg)) { Write-Host "-$cmd necesita un texto." -ForegroundColor $ColError; return }
    $arg = if ($null -eq $arg) { '' } else { $arg.Trim() }
}

$remoto = -not $PSBoundParameters.ContainsKey('File')
$token = if ($PapoiToken) { $PapoiToken } elseif ($env:PAPOI_TOKEN) { $env:PAPOI_TOKEN } else { $null }
if ($remoto -and ($cmd -in 'Add', 'Del', 'Edit') -and -not $token) {
    Write-Host "Modo solo lectura: para -$cmd falta el token de escritura (variable `$PapoiToken o `$env:PAPOI_TOKEN). Disponibles: -Query y -DestList." -ForegroundColor $ColAviso
    return
}
$GhApi = "https://api.github.com/repos/$GhOwner/$GhRepo/contents/$GhPath"
$GhHeaders = @{ Accept = 'application/vnd.github+json'; 'X-GitHub-Api-Version' = '2022-11-28' }
if ($token) { $GhHeaders['Authorization'] = "Bearer $token" }
$GhSha = $null
if ($remoto) {
    try {
        $gh = Invoke-RestMethod -Uri "${GhApi}?ref=$GhBranch" -Headers $GhHeaders -Method Get
        $origBytes = [Convert]::FromBase64String(($gh.content -replace '\s', ''))
        $GhSha = $gh.sha
    } catch {
        Write-Host "No se pudo leer de GitHub: $($_.Exception.Message)" -ForegroundColor $ColError
        return
    }
    $File = "github:$GhOwner/$GhRepo/$GhPath"
}
elseif (-not (Test-Path -LiteralPath $File)) { Write-Host "No se encontro el archivo: $File" -ForegroundColor $ColError; return }
if (-not $remoto) { $File = (Resolve-Path -LiteralPath $File).Path }

# ---------- Carga (UTF-8, preserva BOM y saltos de linea) ----------
if (-not $remoto) { $origBytes = [System.IO.File]::ReadAllBytes($File) }
$hasBom = ($origBytes.Length -ge 3 -and $origBytes[0] -eq 0xEF -and $origBytes[1] -eq 0xBB -and $origBytes[2] -eq 0xBF)
$bomLen = if ($hasBom) { 3 } else { 0 }
$text = [System.Text.UTF8Encoding]::new($false).GetString($origBytes, $bomLen, $origBytes.Length - $bomLen)
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
$endsNL = $text.EndsWith("`n")
$arr = @($text -split "`r?`n")
if ($endsNL -and $arr.Count -gt 0 -and $arr[-1] -eq '') { $arr = @($arr[0..($arr.Count - 2)]) }
$L = [System.Collections.Generic.List[string]]::new([string[]]$arr)

# ---------- Utilidades de pantalla / teclado ----------
function Read-Tecla([int]$seg = $TIMEOUT_S, [string]$def = 'N') {
    # Devuelve el nombre de la tecla ('Enter', 'UpArrow', 'N'...); sin tecla en $seg: $def
    $fin = (Get-Date).AddSeconds($seg)
    while (-not [Console]::KeyAvailable) {
        if ((Get-Date) -ge $fin) { Write-Host "$def (sin respuesta)" -ForegroundColor $ColTimeout; return $def }
        Start-Sleep -Milliseconds 100
    }
    $k = [Console]::ReadKey($true)
    if ([int]$k.KeyChar -ge 32) { Write-Host $k.KeyChar.ToString().ToUpper() -ForegroundColor $ColTeclaEco } else { Write-Host '' }
    $k.Key.ToString()
}

function Write-Prompt([string]$msg, [string]$extra = '') {
    Write-Host "$msg " -ForegroundColor $ColPregunta -NoNewline
    Write-Host 'Enter' -ForegroundColor $ColTeclaPrompt -NoNewline
    Write-Host '=si  ' -ForegroundColor $ColAccionPrompt -NoNewline
    if ($extra) { Write-Host $extra -ForegroundColor $ColAccionPrompt -NoNewline; Write-Host '  ' -NoNewline }
    Write-Host 'Otra' -ForegroundColor $ColTeclaPrompt -NoNewline
    Write-Host '=no ' -ForegroundColor $ColAccionPrompt -NoNewline
    Write-Host "(${TIMEOUT_S}s) " -ForegroundColor $ColTimeout -NoNewline
}

function Confirm-Enter([string]$msg) {
    Write-Prompt $msg
    return ((Read-Tecla $TIMEOUT_S 'N') -eq 'Enter')
}

function Read-LineT([int]$seg = $TIMEOUT_S) {
    # Linea con timeout. Enter devuelve el texto; Esc o timeout devuelven $null
    $sb = [System.Text.StringBuilder]::new()
    $fin = (Get-Date).AddSeconds($seg)
    while ($true) {
        while (-not [Console]::KeyAvailable) {
            if ((Get-Date) -ge $fin) { Write-Host ' (sin respuesta)' -ForegroundColor $ColTimeout; return $null }
            Start-Sleep -Milliseconds 100
        }
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq 'Enter') { Write-Host ''; return $sb.ToString() }
        if ($k.Key -eq 'Escape') { Write-Host ''; return $null }
        if ($k.Key -eq 'Backspace') {
            if ($sb.Length -gt 0) { [void]$sb.Remove($sb.Length - 1, 1); Write-Host "`b `b" -NoNewline }
            continue
        }
        if ([int]$k.KeyChar -ge 32) { [void]$sb.Append($k.KeyChar); Write-Host $k.KeyChar -NoNewline -ForegroundColor $ColTeclaEco }
    }
}

# ---------- Secciones ----------

function Get-Sections {
    $res = [System.Collections.Generic.List[object]]::new()
    $heads = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $L.Count; $i++) {
        if ($L[$i] -match '^##\s+(?<n>.+?)\s*$') { $heads.Add([pscustomobject]@{ I = $i; Name = $Matches['n'] }) }
    }
    for ($h = 0; $h -lt $heads.Count; $h++) {
        $start = $heads[$h].I
        $end = if ($h + 1 -lt $heads.Count) { $heads[$h + 1].I } else { $L.Count }
        $i = $start + 1
        while ($i -lt $end -and $L[$i].Trim() -eq '') { $i++ }
        foreach ($once in 1) {
            if ($i -lt $end) {
                # El 1er renglon no vacio bajo el "##" es el encabezado: define separador y columnas
                $sepCh = $null; $joiner = '; '
                foreach ($ch in ';', '=', ':') {
                    if ($L[$i].Contains($ch)) { $sepCh = $ch; $joiner = [regex]::Match($L[$i], '\s*' + [regex]::Escape($ch) + '\s*').Value; break }
                }
                $cols = if ($sepCh) { @($L[$i] -split [regex]::Escape($sepCh) | ForEach-Object { $_.Trim() }) } else { @($L[$i].Trim()) }
                $rows = [System.Collections.Generic.List[object]]::new()
                $j = $i + 1
                while ($j -lt $end -and $L[$j].Trim() -ne '') {
                    $f = if ($sepCh) { @($L[$j] -split [regex]::Escape($sepCh) | ForEach-Object { $_.Trim() }) } else { @($L[$j].Trim()) }
                    while ($f.Count -lt $cols.Count) { $f += '' }
                    $rows.Add([pscustomobject]@{ Idx = $j; Raw = $L[$j]; F = $f })
                    $j++
                }
                $res.Add([pscustomobject]@{ Name = $heads[$h].Name; Cols = $cols; Header = $L[$i]; Sep = $sepCh; Joiner = $joiner; Last = $j - 1; Rows = $rows })
                break
            }
        }
    }
    return , $res
}

function Find-Section([string]$name) {
    $secs = Get-Sections
    $ex = @($secs | Where-Object { $_.Name -eq $name })
    if ($ex.Count -eq 1) { return $ex[0] }
    $pre = @($secs | Where-Object { $_.Name.StartsWith($name, [StringComparison]::OrdinalIgnoreCase) })
    if ($pre.Count -eq 1) { return $pre[0] }
    if ($pre.Count -gt 1) {
        Write-Host "ERROR(DEST): destino ambiguo '$name' (coincide con: $(($pre | ForEach-Object { $_.Name }) -join ' | '))" -ForegroundColor $ColError
        return $null
    }
    Write-Host "ERROR(DEST): destino '$name' no encontrado. Ver -DestList." -ForegroundColor $ColError
    return $null
}

# Celda igual al termino (o a alguna de sus alternativas separadas por / o ,)
# Con -Strict compara con tildes (solo ignora mayusculas); sin -Strict ignora tildes y dieresis.
function Test-Cell([string]$cell, [string]$term, [switch]$Strict) {
    $t = if ($Strict) { $term.Trim() } else { Get-Norm $term.Trim() }
    foreach ($a in (@($cell) + @($cell -split '[/,]'))) {
        $x = if ($Strict) { $a.Trim() } else { Get-Norm $a.Trim() }
        if ($x -ieq $t) { return $true }
    }
    return $false
}
# Busqueda parcial sin tildes ni mayusculas
function Test-Sub([string]$cell, [string]$term) { return (Get-Norm $cell).Contains((Get-Norm $term)) }

# Normaliza para buscar: minusculas, sin tildes ni dieresis; la ñ se conserva (ña != na)
function Get-Norm([string]$s) {
    $ene = [string][char]0xE000   # marcador temporal para proteger la ñ
    $d = $s.Normalize([System.Text.NormalizationForm]::FormC).ToLowerInvariant().Replace('ñ', $ene).Normalize([System.Text.NormalizationForm]::FormD)
    $sb = [System.Text.StringBuilder]::new()
    foreach ($c in $d.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
    }
    return $sb.ToString()
}

function Find-Rows($sec, [string]$term) {
    # Filas que contienen el termino. Exact = alguna celda coincide completa (sin tildes/mayusculas); si no, es parcial
    $out = @()
    foreach ($r in $sec.Rows) {
        if ((Test-Sub $r.F[0] $term) -or (Test-Sub $r.F[1] $term)) {
            $ex = (Test-Cell $r.F[0] $term) -or (Test-Cell $r.F[1] $term)
            $out += [pscustomobject]@{ Idx = $r.Idx; Raw = $r.Raw; F = $r.F; Exact = $ex }
        }
    }
    return , $out
}

function Select-Rows($hits, [switch]$Multi) {
    # Devuelve las filas elegidas (nada si se cancela)
    if ($hits.Count -eq 1 -and $hits[0].Exact -ne $false) { return $hits[0] }
    for ($k = 0; $k -lt $hits.Count; $k++) {
        $parcial = ($hits[$k].Exact -eq $false)
        Write-Host ('  [{0}] ' -f ($k + 1)) -ForegroundColor $ColListaNum -NoNewline
        Write-Host ('(l.{0}) ' -f ($hits[$k].Idx + 1)) -ForegroundColor $ColListaLinea -NoNewline
        Write-Host $hits[$k].Raw -ForegroundColor $ColListaFila -NoNewline:$parcial
        if ($parcial) { Write-Host '  ~parcial' -ForegroundColor $ColAviso }
    }
    if ($hits.Count -eq 1) {
        if (Confirm-Enter 'Coincidencia parcial. Continuar?') { return $hits[0] } else { return }
    }
    $q = if ($Multi) { 'Numero(s) (ej: 1,3 | T = todas | Enter = cancelar): ' } else { 'Numero (Enter = cancelar): ' }
    Write-Host $q -ForegroundColor $ColPregunta -NoNewline
    $s = Read-LineT
    if ([string]::IsNullOrWhiteSpace($s)) { return }
    if ($Multi -and $s.Trim() -ieq 'T') { return $hits }
    $out = @()
    foreach ($t in ($s -split '[,\s]+' | Where-Object { $_ })) {
        $n = 0
        if ([int]::TryParse($t, [ref]$n) -and $n -ge 1 -and $n -le $hits.Count) { $out += $hits[$n - 1] }
        else { Write-Host "Seleccion invalida: $t" -ForegroundColor $ColError; return }
        if (-not $Multi) { break }
    }
    return $out
}

function Show-Diff([string]$old, [string]$new) {
    if ($old) { Write-Host "- $old" -ForegroundColor $ColDiffQuita }
    if ($new) { Write-Host "+ $new" -ForegroundColor $ColDiffAgrega }
}

function Join-Row($f, $sec) {
    # Une con el separador de la seccion, sin dejar campos vacios al final
    $a = @($f | ForEach-Object { "$_" })
    while ($a.Count -gt 0 -and $a[-1] -eq '') { $a = @($a | Select-Object -First ($a.Count - 1)) }
    $jn = if ($sec -and $sec.Joiner) { $sec.Joiner } else { '; ' }
    return ($a -join $jn)
}

# ---------- Guardado con backup ----------
function New-Backup {
    $dir = Join-Path (Split-Path -Parent $File) 'bk'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $base = [IO.Path]::GetFileNameWithoutExtension($File)
    $zipPath = Join-Path $dir ($base + '_bk.zip')
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $ext = [IO.Path]::GetExtension($File)
    $name = "${base}_${stamp}${ext}"
    $zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $c = 1
        while ($zip.GetEntry($name)) { $name = "${base}_${stamp}_$c${ext}"; $c++ }
        $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $s = $e.Open()
        try { $s.Write($origBytes, 0, $origBytes.Length) } finally { $s.Dispose() }
    } finally { $zip.Dispose() }
    return "$zipPath > $name"
}

function Save-Remote {
    $t = ($L -join $nl)
    if ($endsNL) { $t += $nl }
    $enc = [System.Text.UTF8Encoding]::new($hasBom)
    $bytes = [byte[]]($enc.GetPreamble() + $enc.GetBytes($t))
    $msg = ("papoi: -$cmd $arg").Trim()
    if ($msg.Length -gt 100) { $msg = $msg.Substring(0, 100) }
    $json = @{ message = $msg; content = [Convert]::ToBase64String($bytes); sha = $GhSha; branch = $GhBranch } | ConvertTo-Json
    try {
        $res = Invoke-RestMethod -Uri $GhApi -Headers $GhHeaders -Method Put -Body ([System.Text.Encoding]::UTF8.GetBytes($json)) -ContentType 'application/json; charset=utf-8'
        Write-Host "Guardado en GitHub: $File" -ForegroundColor $ColOk
        if ($res.commit.sha) { Write-Host "Commit  : $($res.commit.sha.Substring(0, 7))" -ForegroundColor $ColSecundario }
    } catch {
        $sc = 0
        try { $sc = [int]$_.Exception.Response.StatusCode } catch { }
        $det = ''
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $det = $_.ErrorDetails.Message }
        if ($sc -eq 409 -or $sc -eq 422) { Write-Host 'NO SE GUARDO: el archivo cambio en GitHub desde que se leyo (otra edicion). Volve a correr el comando.' -ForegroundColor $ColError }
        elseif ($sc -in 401, 403, 404) { Write-Host "NO SE GUARDO ($sc): token invalido, vencido o sin permiso Contents: Read and write sobre $GhOwner/$GhRepo." -ForegroundColor $ColError }
        else { Write-Host "NO SE GUARDO: $($_.Exception.Message)" -ForegroundColor $ColError }
        if ($det) { Write-Host $det -ForegroundColor $ColSecundario }
    }
}

function Save-File {
    if ($remoto) { Save-Remote; return }
    $bk = $null
    try { $bk = New-Backup } catch { Write-Host "FALLO EL BACKUP (se guarda igual): $($_.Exception.Message)" -ForegroundColor $ColError }
    $t = ($L -join $nl)
    if ($endsNL) { $t += $nl }
    [System.IO.File]::WriteAllText($File, $t, [System.Text.UTF8Encoding]::new($hasBom))
    Write-Host "Guardado: $File" -ForegroundColor $ColOk
    if ($bk) { Write-Host "Backup  : $bk" -ForegroundColor $ColSecundario }
}

# ---------- Editor de campo (estilo Append-Csv) ----------
function Read-Field([string]$buf) {
    $pos = $buf.Length
    $left = [Console]::CursorLeft
    $top = [Console]::CursorTop
    $w = [Console]::BufferWidth
    while ($true) {
        [Console]::SetCursorPosition($left, $top)
        Write-Host ($buf + ' ') -ForegroundColor $ColEditBuffer -NoNewline
        $t = $left + $pos
        [Console]::SetCursorPosition([int]($t % $w), [int][Math]::Min($top + [Math]::Floor($t / $w), [Console]::BufferHeight - 1))
        $k = [Console]::ReadKey($true)
        $ctrlS = ($k.Key -eq 'S') -and (($k.Modifiers -band [ConsoleModifiers]::Control) -ne 0)
        if ($ctrlS) { return [pscustomobject]@{ Key = 'CtrlS'; Buf = $buf } }
        elseif ($k.Key -eq 'Enter') { return [pscustomobject]@{ Key = 'Enter'; Buf = $buf } }
        elseif ($k.Key -eq 'UpArrow') { return [pscustomobject]@{ Key = 'Up'; Buf = $buf } }
        elseif ($k.Key -eq 'Escape') { return [pscustomobject]@{ Key = 'Esc'; Buf = $buf } }
        elseif ($k.Key -eq 'LeftArrow') { if ($pos -gt 0) { $pos-- } }
        elseif ($k.Key -eq 'RightArrow') { if ($pos -lt $buf.Length) { $pos++ } }
        elseif ($k.Key -eq 'Home') { $pos = 0 }
        elseif ($k.Key -eq 'End') { $pos = $buf.Length }
        elseif ($k.Key -eq 'Backspace') { if ($pos -gt 0) { $buf = $buf.Remove($pos - 1, 1); $pos-- } }
        elseif ($k.Key -eq 'Delete') { if ($pos -lt $buf.Length) { $buf = $buf.Remove($pos, 1) } }
        elseif ([int]$k.KeyChar -ge 32) { $buf = $buf.Insert($pos, [string]$k.KeyChar); $pos++ }
    }
}

function Edit-Row($sec, $row) {
    # Devuelve $true si modifico $L
    $n = $sec.Cols.Count
    $orig = @($row.F)
    $vals = @($orig | ForEach-Object { "$_" })
    $oldNorm = Join-Row $orig $sec
    $i = 0
    while ($true) {
        if ($i -ge $n) {
            $newLine = Join-Row $vals $sec
            Write-Host ''
            if ($newLine -ceq $oldNorm) { Write-Host 'Sin cambios.' -ForegroundColor $ColAviso; return $false }
            Show-Diff $row.Raw $newLine
            Write-Host ''
            Write-Prompt 'Guardar?' 'Arriba=volver'
            $t = Read-Tecla $TIMEOUT_S 'N'
            if ($t -eq 'Enter') { $L[$row.Idx] = $newLine; $script:editOldPap = $orig[1]; $script:editNewPap = $vals[1]; return $true }
            elseif ($t -eq 'UpArrow') { $i = $n - 1; continue }
            else { Write-Host 'Cancelado.' -ForegroundColor $ColCancelado; return $false }
        }
        Clear-Host
        Write-Host 'papoi-edit -Edit' -ForegroundColor $ColTitulo -NoNewline
        Write-Host "  [$($sec.Name)]  l.$($row.Idx + 1)" -ForegroundColor $ColSecundario
        Write-Host 'Enter=acepta  Izq/Der=cursor  Arriba=anterior  Ctrl+S=guardar ya  Esc=cancelar' -ForegroundColor $ColEditAyuda
        Write-Host ''
        for ($k = 0; $k -lt $i; $k++) {
            Write-Host ('  {0}: ' -f $sec.Cols[$k]) -ForegroundColor $ColEditEtiqueta -NoNewline
            Write-Host $vals[$k] -ForegroundColor $ColEditValor
        }
        Write-Host ''
        Write-Host ('[{0}/{1}] {2}:' -f ($i + 1), $n, $sec.Cols[$i]) -ForegroundColor $ColEditCampo
        Write-Host "actual: $($orig[$i])" -ForegroundColor $ColEditActual
        Write-Host ''
        Write-Host 'editar: ' -ForegroundColor $ColEditPrompt -NoNewline
        $r = Read-Field $vals[$i]
        $b = $r.Buf.Replace(';', ',')
        switch ($r.Key) {
            'Enter' { $vals[$i] = $b; $i++ }
            'Up'    { $vals[$i] = $b; if ($i -gt 0) { $i-- } }
            'CtrlS' { $vals[$i] = $b; $i = $n }
            'Esc'   {
                Write-Host ''
                if (Confirm-Enter 'Cancelar la edicion?') { Write-Host 'Cancelado.' -ForegroundColor $ColCancelado; return $false }
            }
        }
    }
}

# ---------- Otras secciones: apariciones, reemplazo y marca de borrado ----------
# Solo se propagan terminos papoi (2da columna de la seccion editada); las palabras en español no se tocan.
$MarcaBorrado = '[BORRADO]'

function Get-WordRegex([string]$term, [switch]$SinMarca) {
    # Palabra completa (sin letras/numeros pegados), sin distinguir mayusculas. -SinMarca: ignora las que ya tienen la marca
    $p = '(?<![\p{L}\p{N}])' + [regex]::Escape($term) + '(?![\p{L}\p{N}])'
    if ($SinMarca) { $p += '(?! ' + [regex]::Escape($MarcaBorrado) + ')' }
    return [regex]::new($p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Get-Root([string]$t) {
    # Raiz de un termino papoi: sin la particula final (ñapoi -> ña, putoi -> puto)
    if ($t -match '(?i)^(.+)poi$') { return $Matches[1] }
    if ($t -match '(?i)^(.+)i$') { return $Matches[1] }
    return $t
}

function Get-KnownTerms {
    # Terminos papoi que existen como entrada propia en alguna seccion (para no confundirlos con palabras fusionadas)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($s in (Get-Sections)) {
        if ($s.Cols.Count -lt 2) { continue }
        foreach ($r in $s.Rows) {
            foreach ($a in ($r.F[1] -split '[/,]')) { $t = $a.Trim(); if ($t) { [void]$set.Add($t) } }
        }
    }
    return , $set
}

function Find-Occurrences($destSec, [string]$term) {
    # Filas de OTRAS secciones donde aparece el termino como palabra completa
    $rx = Get-WordRegex $term
    $out = @()
    foreach ($s in (Get-Sections)) {
        if ($s.Name -eq $destSec.Name) { continue }
        foreach ($r in $s.Rows) {
            foreach ($c in $r.F) {
                if ($c -and $rx.IsMatch($c)) { $out += [pscustomobject]@{ Sec = $s.Name; Idx = $r.Idx; Raw = $r.Raw }; break }
            }
        }
    }
    return , $out
}

function Get-WholeChanges($destSec, [string[]]$terms, [string]$mode, [string]$newTerm) {
    # Palabra completa en otras secciones. mode 'Mark': agrega [BORRADO] junto al termino; mode 'Replace': cambia $terms[0] por $newTerm
    $out = @()
    foreach ($s in (Get-Sections)) {
        if ($s.Name -eq $destSec.Name) { continue }
        $single = -not $s.Sep   # seccion de 1 columna sin separador: se trabaja sobre la linea cruda para no perder la sangria
        foreach ($r in $s.Rows) {
            $f = if ($single) { @($r.Raw) } else { @($r.F) }; $hit = $false
            for ($c = 0; $c -lt $f.Count; $c++) {
                foreach ($t in $terms) {
                    $rx = if ($mode -eq 'Mark') { Get-WordRegex $t -SinMarca } else { Get-WordRegex $t }
                    if ($f[$c] -and $rx.IsMatch($f[$c])) {
                        $rep = if ($mode -eq 'Mark') { '$0 ' + $MarcaBorrado } else { $newTerm.Replace('$', '$$') }
                        $f[$c] = $rx.Replace($f[$c], $rep); $hit = $true
                    }
                }
            }
            if ($hit) {
                $nl2 = if ($single) { $f[0] } else { Join-Row $f $s }
                if ($nl2 -cne $r.Raw) { $out += [pscustomobject]@{ Sec = $s.Name; Idx = $r.Idx; Old = $r.Raw; New = $nl2 } }
            }
        }
    }
    return , $out
}

function Get-FusionChanges($destSec, [string]$oldTerm, [string]$newTerm) {
    # Palabras unidas que contienen la raiz del termino (ñapoi -> "ña" dentro de ñakachepoi), terminan en poi/i y no existen como termino propio.
    # Con $newTerm propone cambiar la raiz; sin $newTerm solo informa (New = $null)
    $out = @()
    $root = Get-Root $oldTerm
    if ($root.Length -lt 2) { return , $out }
    $newRoot = if ($newTerm) { Get-Root $newTerm } else { $null }
    $known = Get-KnownTerms
    foreach ($s in (Get-Sections)) {
        if ($s.Name -eq $destSec.Name) { continue }
        $single = -not $s.Sep
        foreach ($r in $s.Rows) {
            $f = if ($single) { @($r.Raw) } else { @($r.F) }; $hit = $false
            # Las fusiones son solo de terminos papoi: en secciones con columnas se saltea la 1ra (español)
            $c0 = if ($single -or $s.Cols.Count -lt 2) { 0 } else { 1 }
            for ($c = $c0; $c -lt $f.Count; $c++) {
                if (-not $f[$c]) { continue }
                $toks = @([regex]::Matches($f[$c], "[\p{L}\p{N}']+") | ForEach-Object { $_.Value } | Select-Object -Unique)
                foreach ($tk in $toks) {
                    if ($tk -ieq $oldTerm -or $tk.IndexOf($root, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                    if ($tk -notmatch '(?i)(poi|i)$' -or $known.Contains($tk)) { continue }
                    $hit = $true
                    if ($newRoot) {
                        $ntk = [regex]::Replace($tk, [regex]::Escape($root), $newRoot.Replace('$', '$$'), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        $f[$c] = (Get-WordRegex $tk).Replace($f[$c], $ntk.Replace('$', '$$'))
                    }
                }
            }
            if ($hit) {
                $nl2 = if ($newRoot) { if ($single) { $f[0] } else { Join-Row $f $s } } else { $null }
                if ($newRoot -and $nl2 -ceq $r.Raw) { continue }
                $out += [pscustomobject]@{ Sec = $s.Name; Idx = $r.Idx; Old = $r.Raw; New = $nl2 }
            }
        }
    }
    return , $out
}

function Select-Changes($changes, [string]$titulo) {
    # Muestra cada cambio como diff (- rojo / + verde) y devuelve los elegidos
    Write-Host ''
    Write-Host $titulo -ForegroundColor $ColAviso
    for ($k = 0; $k -lt $changes.Count; $k++) {
        Write-Host ('  [{0}] ' -f ($k + 1)) -ForegroundColor $ColListaNum -NoNewline
        Write-Host "[$($changes[$k].Sec)] " -ForegroundColor $ColTitulo -NoNewline
        Write-Host ('(l.{0})' -f ($changes[$k].Idx + 1)) -ForegroundColor $ColListaLinea
        Show-Diff $changes[$k].Old $changes[$k].New
    }
    if ($changes.Count -eq 1) { if (Confirm-Enter 'Aplicar?') { return $changes[0] } else { return } }
    Write-Host 'Aplicar (ej: 1,3 | T = todas | Enter = ninguna): ' -ForegroundColor $ColPregunta -NoNewline
    $s = Read-LineT
    if ([string]::IsNullOrWhiteSpace($s)) { return }
    if ($s.Trim() -ieq 'T') { return $changes }
    $out = @()
    foreach ($t in ($s -split '[,\s]+' | Where-Object { $_ })) {
        $n = 0
        if ([int]::TryParse($t, [ref]$n) -and $n -ge 1 -and $n -le $changes.Count) { $out += $changes[$n - 1] }
        else { Write-Host "Seleccion invalida: $t" -ForegroundColor $ColError; return }
    }
    return $out
}

function Invoke-PropEdit($sec, [string]$oldPap, [string]$newPap) {
    # Tras editar el termino papoi: ofrece cambiarlo en otras secciones (palabra completa y fusionadas)
    $oldPap = $oldPap.Trim(); $newPap = $newPap.Trim()
    if (-not $oldPap -or -not $newPap -or $oldPap -ceq $newPap) { return }
    if ($oldPap -match '[/,]' -or $newPap -match '[/,]') {
        Write-Host '(celda con alternativas: no se propaga a otras secciones)' -ForegroundColor $ColSecundario
        return
    }
    $whole = Get-WholeChanges $sec @($oldPap) 'Replace' $newPap
    if ($whole.Count -gt 0) {
        $sel = @(Select-Changes $whole "Apariciones de '$oldPap' en otras secciones (cambiar por '$newPap'):")
        foreach ($c in $sel) { $L[$c.Idx] = $c.New }
    }
    $fus = Get-FusionChanges $sec $oldPap $newPap
    if ($fus.Count -gt 0) {
        Write-Host ''
        Write-Host "Hay $($fus.Count) linea(s) con el termino fusionado (raiz '$(Get-Root $oldPap)' -> '$(Get-Root $newPap)')." -ForegroundColor $ColAviso
        $sel = @(Select-Changes $fus 'Cambiar tambien las fusionadas?')
        foreach ($c in $sel) { $L[$c.Idx] = $c.New }
    }
}

function Invoke-PropDel($sec, [string[]]$terms) {
    # Antes de borrar: ofrece marcar con [BORRADO] las apariciones (palabra completa) en otras secciones; las fusionadas solo se avisan
    $whole = Get-WholeChanges $sec $terms 'Mark' ''
    if ($whole.Count -gt 0) {
        $sel = @(Select-Changes $whole "Apariciones en otras secciones (marcar con $MarcaBorrado):")
        foreach ($c in $sel) { $L[$c.Idx] = $c.New }
    }
    $warn = @()
    foreach ($t in $terms) {
        $x = Get-FusionChanges $sec $t $null
        foreach ($w in $x) { if (-not ($warn | Where-Object { $_.Idx -eq $w.Idx })) { $warn += $w } }
    }
    if ($warn.Count -gt 0) {
        Write-Host ''
        Write-Host 'Fusionadas con el termino borrado (revisar a mano):' -ForegroundColor $ColAviso
        foreach ($w in $warn) {
            Write-Host '* ' -ForegroundColor $ColAviso -NoNewline
            Write-Host "[$($w.Sec)] " -ForegroundColor $ColTitulo -NoNewline
            Write-Host ('(l.{0}) ' -f ($w.Idx + 1)) -ForegroundColor $ColListaLinea -NoNewline
            Write-Host $w.Old -ForegroundColor $ColCoincidencia
        }
    }
}

# ---------- Parseo de -Add ----------
function Split-AddInput([string]$s, $sec) {
    # Devuelve lista de arrays de campos, o $null si hay error
    $nCols = $sec.Cols.Count
    if ($nCols -lt 2) { $o = [System.Collections.Generic.List[object]]::new(); $o.Add(@($s.Trim())); return , $o }
    $sep = $null
    foreach ($ch in (@($sec.Sep) + ';', '=', ':')) { if ($ch -and $s.Contains($ch)) { $sep = $ch; break } }
    if (-not $sep) { Write-Host "Formato de la seccion: $($sec.Header)  (se acepta ; = o : como separador)" -ForegroundColor $ColError; return $null }
    $p = @($s -split $sep, 3 | ForEach-Object { $_.Trim() })
    if ($p.Count -lt 2 -or $p[0] -eq '' -or $p[1] -eq '') { Write-Host 'Faltan datos: se necesitan al menos español y papoi.' -ForegroundColor $ColError; return $null }
    if ($p.Count -gt $nCols -and -not ($p.Count -eq 3 -and $p[2] -eq '')) { Write-Host "La seccion destino tiene $nCols columnas." -ForegroundColor $ColError; return $null }
    $vars = @($p[1] -split '/' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($vars.Count -eq 0) { Write-Host 'Falta el termino papoi.' -ForegroundColor $ColError; return $null }
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($v in $vars) {
        $f = @($p[0], $v)
        if ($p.Count -ge 3 -and $p[2] -ne '') { $f += $p[2] }
        $out.Add($f)
    }
    return , $out
}

# ---------- -Add interactivo ----------
function Test-HasSep([string]$s, $sec) {
    # Mismo criterio que Split-AddInput: el texto trae algun separador (; = :)
    foreach ($ch in (@($sec.Sep) + ';', '=', ':')) { if ($ch -and $s.Contains($ch)) { return $true } }
    return $false
}

function Write-HitLine($r, [string]$suf) {
    Write-Host '  * ' -ForegroundColor $ColAviso -NoNewline
    Write-Host ('(l.{0}) ' -f ($r.Idx + 1)) -ForegroundColor $ColListaLinea -NoNewline
    Write-Host $r.Raw -ForegroundColor $ColCoincidencia -NoNewline
    Write-Host $suf -ForegroundColor $ColAviso
}

function Show-AddChecks($sec, $vals, [int]$upto) {
    # Existencia de los datos ya confirmados (columnas 0..$upto-1). Solo informa; max 6 lineas por bloque
    $max = 6
    if ($upto -ge 1 -and $vals[0] -ne '') {
        $h = @($sec.Rows | Where-Object { Test-Sub $_.F[0] $vals[0] })
        if ($h.Count -gt 0) {
            Write-Host "  --Coincidencias en $($sec.Cols[0]):" -ForegroundColor $ColAviso
            foreach ($r in ($h | Select-Object -First $max)) {
                $suf = if (Test-Cell $r.F[0] $vals[0]) { '' } else { '  ~parcial' }
                Write-HitLine $r $suf
            }
            if ($h.Count -gt $max) { Write-Host "  ... (+$($h.Count - $max) mas)" -ForegroundColor $ColSecundario }
        }
    }
    if ($sec.Cols.Count -ge 2 -and $upto -ge 2 -and $vals[1] -ne '') {
        foreach ($v in @($vals[1] -split '/' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
            $h = @($sec.Rows | Where-Object { Test-Sub $_.F[1] $v })
            if ($h.Count -gt 0) {
                Write-Host "  --Coincidencias en $($sec.Cols[1]) '${v}':" -ForegroundColor $ColAviso
                foreach ($r in ($h | Select-Object -First $max)) {
                    $suf = ''
                    if (-not (Test-Cell $r.F[1] $v)) { $suf = '  ~parcial' }
                    elseif ((Test-Cell $r.F[0] $vals[0] -Strict) -and (Test-Cell $r.F[1] $v -Strict)) { $suf = '  <= ya existe' }
                    Write-HitLine $r $suf
                }
                if ($h.Count -gt $max) { Write-Host "  ... (+$($h.Count - $max) mas)" -ForegroundColor $ColSecundario }
            }
            $occ = Find-Occurrences $sec $v
            if ($occ.Count -gt 0) {
                Write-Host "  --Aparece en otras secciones '${v}':" -ForegroundColor $ColAviso
                foreach ($o in ($occ | Select-Object -First $max)) {
                    Write-Host '  * ' -ForegroundColor $ColAviso -NoNewline
                    Write-Host "[$($o.Sec)] " -ForegroundColor $ColTitulo -NoNewline
                    Write-Host ('(l.{0}) ' -f ($o.Idx + 1)) -ForegroundColor $ColListaLinea -NoNewline
                    Write-Host $o.Raw -ForegroundColor $ColCoincidencia
                }
                if ($occ.Count -gt $max) { Write-Host "  ... (+$($occ.Count - $max) mas)" -ForegroundColor $ColSecundario }
            }
        }
    }
}

function Read-AddInteractive($sec, [string]$first) {
    # Pide los campos de a uno (como -Edit) y devuelve la linea "esp; papoi; ..." para Split-AddInput, o $null si se cancela.
    # $first = 1ra columna ya dada: arranca por la 2da. Obligatorias: las 2 primeras columnas (1 si la seccion tiene una sola)
    $n = $sec.Cols.Count
    $req = [Math]::Min(2, $n)
    $sepCh = if ($sec.Sep) { $sec.Sep } else { ';' }
    $vals = @(1..$n | ForEach-Object { '' })
    $i = 0
    if ($first) { $vals[0] = $first; $i = 1 }
    $err = ''
    while ($true) {
        if ($i -ge $n) {
            $line = Join-Row $vals $sec
            Write-Host ''
            Show-Diff '' $line
            Write-Host ''
            Write-Prompt 'Agregar?' 'Arriba=volver'
            $t = Read-Tecla $TIMEOUT_S 'N'
            if ($t -eq 'Enter') { return $line }
            elseif ($t -eq 'UpArrow') { $i = $n - 1; continue }
            else { return $null }
        }
        Clear-Host
        Write-Host 'papoi -Add' -ForegroundColor $ColTitulo -NoNewline
        Write-Host "  [$($sec.Name)]" -ForegroundColor $ColSecundario
        Write-Host 'Enter=acepta  Izq/Der=cursor  Arriba=anterior  Ctrl+S=guardar ya  Esc=cancelar' -ForegroundColor $ColEditAyuda
        Write-Host ''
        for ($k = 0; $k -lt $i; $k++) {
            Write-Host ('  {0}: ' -f $sec.Cols[$k]) -ForegroundColor $ColEditEtiqueta -NoNewline
            Write-Host $vals[$k] -ForegroundColor $ColEditValor
        }
        Show-AddChecks $sec $vals $i
        Write-Host ''
        if ($err) { Write-Host $err -ForegroundColor $ColError; $err = '' }
        $opt = if ($i -ge $req) { '  (opcional)' } else { '' }
        Write-Host ('[{0}/{1}] {2}:{3}' -f ($i + 1), $n, $sec.Cols[$i], $opt) -ForegroundColor $ColEditCampo
        Write-Host ''
        Write-Host 'agregar: ' -ForegroundColor $ColEditPrompt -NoNewline
        $r = Read-Field $vals[$i]
        $b = $r.Buf.Replace($sepCh, ',').Trim()
        switch ($r.Key) {
            'Enter' {
                if ($i -lt $req -and $b -eq '') { $err = "Falta: $($sec.Cols[$i])" }
                else { $vals[$i] = $b; $i++ }
            }
            'Up'    { $vals[$i] = $b; if ($i -gt 0) { $i-- } }
            'CtrlS' {
                $vals[$i] = $b
                $miss = -1
                for ($k = 0; $k -lt $req; $k++) { if ($vals[$k] -eq '') { $miss = $k; break } }
                if ($miss -ge 0) { $err = "Faltan datos obligatorios: $($sec.Cols[$miss])"; $i = $miss }
                else { $i = $n }
            }
            'Esc'   {
                Write-Host ''
                if (Confirm-Enter 'Cancelar?') { return $null }
            }
        }
    }
}

# ================= Comandos =================
$changed = $false

switch ($cmd) {

    'DestList' {
        $secs = Get-Sections
        if ($secs.Count -eq 0) { Write-Host 'No se detectaron secciones.' -ForegroundColor $ColAviso; break }
        Write-Host ''
        foreach ($s in $secs) {
            Write-Host $s.Name -ForegroundColor $ColTitulo -NoNewline
            Write-Host ("  ({0} cols, {1} filas)" -f $s.Cols.Count, $s.Rows.Count) -ForegroundColor $ColSecundario
            Write-Host "  formato: $($s.Header)" -ForegroundColor $ColDestFormato
            $ej = if ($s.Rows.Count -gt 0) { $s.Rows[0].Raw } else { '(sin filas)' }
            Write-Host "  ejemplo: $ej" -ForegroundColor $ColDestEjemplo
        }
        Write-Host ''
    }

    'Query' {
        # Sin -Dest busca en todas las secciones; con -Dest, solo en esa
        $todas = -not $PSBoundParameters.ContainsKey('Dest')
        if ($todas) { $targets = Get-Sections }
        else {
            $one = Find-Section $Dest
            if (-not $one) { break }
            $targets = @($one)
        }
        $total = 0
        Write-Host ''
        foreach ($sec in $targets) {
            $hits = Find-Rows $sec $arg
            if ($hits.Count -eq 0) { continue }
            $total += $hits.Count
            if ($todas) { Write-Host "[$($sec.Name)]" -ForegroundColor $ColTitulo }
            $dos = ($sec.Cols.Count -gt 1)
            foreach ($r in $hits) {
                # Columnas 3+ (si tienen dato): se muestran detras, separadas por " ; "
                $extra = @()
                if ($r.F.Count -gt 2) { $extra = @($r.F[2..($r.F.Count - 1)] | Where-Object { $_ -ne '' -and $_ -ne '-' }) }
                Write-Host $r.F[0] -ForegroundColor $ColQueryEsp -NoNewline:$dos
                if ($dos) {
                    Write-Host ' = ' -ForegroundColor $ColQuerySep -NoNewline
                    Write-Host $r.F[1] -ForegroundColor $ColQueryPapoi -NoNewline:($extra.Count -gt 0)
                    for ($k = 0; $k -lt $extra.Count; $k++) {
                        Write-Host ' ; ' -ForegroundColor $ColQuerySep -NoNewline
                        Write-Host $extra[$k] -ForegroundColor $ColQueryPapoi -NoNewline:($k -lt $extra.Count - 1)
                    }
                }
            }
            Write-Host ''
        }
        if ($total -eq 0) { Write-Host "Sin coincidencias para '$arg'." -ForegroundColor $ColAviso; Write-Host '' }
        else { Write-Host "$total coincidencia(s)" -ForegroundColor $ColSecundario; Write-Host '' }
    }

    'Del' {
        $sec = Find-Section $Dest
        if (-not $sec) { break }
        $hits = Find-Rows $sec $arg
        Write-Host ''
        if ($hits.Count -eq 0) { Write-Host "Sin coincidencias para '$arg' en [$($sec.Name)]." -ForegroundColor $ColAviso; break }
        Write-Host "[$($sec.Name)] coincidencias: $($hits.Count)" -ForegroundColor $ColTitulo
        $sel = @(Select-Rows $hits -Multi)
        if ($sel.Count -eq 0) { Write-Host 'Cancelado.' -ForegroundColor $ColCancelado; break }
        Write-Host ''
        foreach ($r in $sel) { Show-Diff $r.Raw '' }
        Write-Host ''
        if (Confirm-Enter "Borrar $($sel.Count) linea(s)?") {
            # Otras secciones: se procesan ANTES de borrar para que los indices de linea sigan validos
            if ($sec.Cols.Count -ge 2) {
                $terms = @($sel | ForEach-Object { $_.F[1] -split '[/,]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
                if ($terms.Count -gt 0) { Invoke-PropDel $sec $terms }
            }
            foreach ($r in ($sel | Sort-Object Idx -Descending)) { $L.RemoveAt($r.Idx) }
            $changed = $true
        } else { Write-Host 'Cancelado.' -ForegroundColor $ColCancelado }
    }

    'Edit' {
        $sec = Find-Section $Dest
        if (-not $sec) { break }
        $hits = Find-Rows $sec $arg
        Write-Host ''
        if ($hits.Count -eq 0) { Write-Host "Sin coincidencias para '$arg' en [$($sec.Name)]." -ForegroundColor $ColAviso; break }
        Write-Host "[$($sec.Name)] coincidencias: $($hits.Count)" -ForegroundColor $ColTitulo
        $sel = @(Select-Rows $hits)
        if ($sel.Count -eq 0) { Write-Host 'Cancelado.' -ForegroundColor $ColCancelado; break }
        if (Edit-Row $sec $sel[0]) {
            $changed = $true
            if ($sec.Cols.Count -ge 2) { Invoke-PropEdit $sec $script:editOldPap $script:editNewPap }
        }
    }

    'Add' {
        $sec = Find-Section $Dest
        if (-not $sec) { break }
        # Interactivo: sin texto, o palabra sola (sin separador) en una seccion con columnas. $arg pasa a ser la linea armada
        $interactivo = ($arg -eq '') -or ($sec.Cols.Count -ge 2 -and -not (Test-HasSep $arg $sec))
        if ($interactivo) {
            $arg = Read-AddInteractive $sec $arg
            if ($null -eq $arg) { Write-Host 'Cancelado.' -ForegroundColor $ColCancelado; break }
        }
        $recs = Split-AddInput $arg $sec
        if (-not $recs) { break }
        $snap = $sec.Rows   # coincidencias se evaluan contra lo que habia al empezar
        $nAdd = 0; $nOver = 0
        Write-Host ''
        Write-Host "[$($sec.Name)]" -ForegroundColor $ColTitulo
        foreach ($f in $recs) {
            $line = Join-Row $f $sec
            $esp = $f[0]; $pap = if ($f.Count -gt 1) { $f[1] } else { '' }
            Write-Host "Registro: $line" -ForegroundColor $ColTexto
            # hits = parciales (contienen el termino) + completas; full = celda igual sin tildes (solo estas piden confirmacion); exact = igual con tildes
            $hits = @($snap | Where-Object { (Test-Sub $_.F[0] $esp) -or ($pap -ne '' -and (Test-Sub $_.F[1] $pap)) })
            $exact = @($hits | Where-Object { (Test-Cell $_.F[0] $esp -Strict) -and ($pap -eq '' -or (Test-Cell $_.F[1] $pap -Strict)) })
            $full = @($hits | Where-Object { (Test-Cell $_.F[0] $esp) -or ($pap -ne '' -and (Test-Cell $_.F[1] $pap)) })
            if ($hits.Count -gt 0 -and -not $interactivo) {
                Write-Host '--Coincidencias:' -ForegroundColor $ColAviso
                foreach ($h in $hits) {
                    $parcial = ($full -notcontains $h)
                    Write-Host '* ' -ForegroundColor $ColAviso -NoNewline
                    Write-Host ('(l.{0}) ' -f ($h.Idx + 1)) -ForegroundColor $ColListaLinea -NoNewline
                    Write-Host $h.Raw -ForegroundColor $ColCoincidencia -NoNewline:$parcial
                    if ($parcial) { Write-Host '  ~parcial' -ForegroundColor $ColAviso }
                }
            }
            if ($pap -ne '' -and $sec.Cols.Count -ge 2 -and -not $interactivo) {
                $occ = Find-Occurrences $sec $pap
                if ($occ.Count -gt 0) {
                    Write-Host '--Aparece en otras secciones:' -ForegroundColor $ColAviso
                    foreach ($o in $occ) {
                        Write-Host '* ' -ForegroundColor $ColAviso -NoNewline
                        Write-Host "[$($o.Sec)] " -ForegroundColor $ColTitulo -NoNewline
                        Write-Host ('(l.{0}) ' -f ($o.Idx + 1)) -ForegroundColor $ColListaLinea -NoNewline
                        Write-Host $o.Raw -ForegroundColor $ColCoincidencia
                    }
                }
            }
            if ($exact.Count -gt 0) {
                Write-Host '-- El registro ya existe:' -ForegroundColor $ColAviso
                $cand = @($exact | Where-Object { (Join-Row $_.F $sec) -cne $line })
                if ($cand.Count -eq 0) { Write-Host 'Identico, nada que hacer.' -ForegroundColor $ColSecundario; Write-Host ''; continue }
                $sel = @($cand)
                if ($cand.Count -gt 1) {
                    Write-Host 'Cual sobreescribir?' -ForegroundColor $ColPregunta
                    $sel = @(Select-Rows $cand -Multi)
                }
                if ($sel.Count -eq 0) { Write-Host 'Omitido.' -ForegroundColor $ColSecundario; Write-Host ''; continue }
                foreach ($o in $sel) { Show-Diff $o.Raw $line }
                if (Confirm-Enter 'Sobreescribir?') {
                    foreach ($o in $sel) {
                        $ix = $L.IndexOf($o.Raw)
                        if ($ix -lt 0) { Write-Host 'No se encontro la linea original (cambio en esta corrida); se omite.' -ForegroundColor $ColError }
                        else { $L[$ix] = $line; $nOver++; $changed = $true }
                    }
                } else { Write-Host 'Omitido.' -ForegroundColor $ColSecundario }
            } else {
                if ($full.Count -gt 0 -and -not $interactivo -and -not (Confirm-Enter 'Agregar igual?')) { Write-Host 'Omitido.' -ForegroundColor $ColSecundario; Write-Host ''; continue }
                $cur = Find-Section $Dest
                $L.Insert($cur.Last + 1, $line)
                Write-Host "+ $line" -ForegroundColor $ColDiffAgrega
                $nAdd++; $changed = $true
            }
            Write-Host ''
        }
        if ($changed) { Write-Host "Agregados: $nAdd | Sobreescritos: $nOver" -ForegroundColor $ColTitulo }
    }
}

if ($changed) { Save-File }
