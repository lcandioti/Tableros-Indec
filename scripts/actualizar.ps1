<#
  Descarga los indicadores del Catalogo API-INDEC y genera data/datos.json y data/datos.js.

  Uso:
    .\scripts\actualizar.ps1              Actualiza los datos (comercio exterior: incremental)
    .\scripts\actualizar.ps1 -Completo    Vuelve a bajar comercio exterior desde $AnioInicioComercio
    .\scripts\actualizar.ps1 -Descubrir   Lista los valores de cada dimension en data/_catalogo.txt

  Token: variable de entorno INDEC_TOKEN (local: de usuario; GitHub Actions: secret).
  Funciona en Windows PowerShell 5.1 y en PowerShell 7.
#>
[CmdletBinding()]
param([switch]$Completo, [switch]$Descubrir)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Base = 'https://api.indec.gob.ar/v1'
$AnioInicioComercio = 2016
$Raiz = Split-Path -Parent $PSScriptRoot
$DirDatos = Join-Path $Raiz 'data'
$Inv = [Globalization.CultureInfo]::InvariantCulture

$Token = $env:INDEC_TOKEN
if (-not $Token) { $Token = [Environment]::GetEnvironmentVariable('INDEC_TOKEN', 'User') }
if (-not $Token) {
    throw 'Falta el token. Generalo en https://api.indec.gob.ar/ y guardalo en la variable de entorno INDEC_TOKEN.'
}
$Token = $Token.Trim()
$Cabeceras = @{ Authorization = "Bearer $Token"; Accept = 'application/json' }

# ---------------------------------------------------------------- API

function Get-Pagina([string]$Url) {
    for ($intento = 1; ; $intento++) {
        try {
            $r = Invoke-WebRequest -Uri $Url -Headers $Cabeceras -UseBasicParsing -TimeoutSec 120
            # Se decodifica a mano: PowerShell 5.1 asume otra codificacion y rompe los acentos.
            $texto = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
            return @($texto | ConvertFrom-Json) | ForEach-Object { $_ }
        }
        catch {
            $codigo = 0
            if ($_.Exception.Response) { $codigo = [int]$_.Exception.Response.StatusCode }
            if ($codigo -eq 401) { throw 'La API respondio 401: el token falta, es invalido o vencio (dura 12 meses).' }
            if ($codigo -eq 400 -or $codigo -eq 403 -or $intento -ge 5) { throw "Fallo $Url (HTTP $codigo): $($_.Exception.Message)" }
            Start-Sleep -Seconds ([Math]::Pow(2, $intento))
        }
    }
}

# Recorre una tabla completa. La API devuelve hasta 1000 filas y exige ordenar por la clave para paginar.
function Get-Filas([string]$Tabla, [string]$Orden, [string]$Filtros = '') {
    $filas = New-Object System.Collections.Generic.List[object]
    $offset = 0
    while ($true) {
        $url = "$Base/$Tabla`?order=$Orden&limit=1000&offset=$offset"
        if ($Filtros) { $url += "&$Filtros" }
        $pagina = @(Get-Pagina $url)
        foreach ($f in $pagina) { $filas.Add($f) }
        if ($pagina.Count -lt 1000) { break }
        $offset += 1000
        Start-Sleep -Milliseconds 150
    }
    Write-Host ("  {0}: {1} filas" -f $Tabla, $filas.Count)
    return , $filas
}

# ---------------------------------------------------------------- Modo descubrimiento

if ($Descubrir) {
    $tablas = @(
        @{ t = 'estimador-mensual-de-actividad-economica-nivel-general'; o = 'periodo.desc' },
        @{ t = 'estimador-mensual-de-actividad-economica-sectores'; o = 'periodo.desc' },
        @{ t = 'comercio_exterior_exportacion_bienes'; o = 'Periodo.desc' },
        @{ t = 'comercio_exterior_importacion_bienes'; o = 'Periodo.desc' },
        @{ t = 'ipc_nacional'; o = 'periodo.desc' },
        @{ t = 'indice_salarios'; o = 'periodo.desc' },
        @{ t = 'icc_precios_construccion'; o = 'periodo.desc' },
        @{ t = 'sipm_precios_mayoristas'; o = 'periodo.desc' },
        @{ t = 'mercado_trabajo_eph'; o = 'periodo.desc' },
        @{ t = 'comercio_exterior_nomenclaturas'; o = 'nombre_columna' }
    )
    $salida = New-Object System.Collections.Generic.List[string]
    foreach ($x in $tablas) {
        $salida.Add("### $($x.t)")
        $filas = @(Get-Pagina "$Base/$($x.t)?order=$($x.o)&limit=1000")
        if ($filas.Count -eq 0) { $salida.Add('  (sin filas)'); continue }
        foreach ($col in $filas[0].PSObject.Properties.Name) {
            $vals = @($filas | ForEach-Object { "$($_.$col)" } | Sort-Object -Unique)
            $muestra = ($vals | Select-Object -First 40) -join ' | '
            $salida.Add("  $col ($($vals.Count)): $muestra")
        }
    }
    if (-not (Test-Path $DirDatos)) { New-Item -ItemType Directory $DirDatos | Out-Null }
    $destino = Join-Path $DirDatos '_catalogo.txt'
    [IO.File]::WriteAllLines($destino, $salida, (New-Object Text.UTF8Encoding($false)))
    Write-Host "Catalogo escrito en $destino"
    return
}

# ---------------------------------------------------------------- Armado de series

function Fecha([string]$Periodo) { '{0}-{1}' -f $Periodo.Substring(0, 4), $Periodo.Substring(4, 2) }

function Numero($v) {
    if ($null -eq $v -or "$v" -eq '') { return $null }
    return [double]::Parse("$v", $Inv)
}

# $mapa: nombre de serie -> (fecha -> valor). Devuelve lista de {id, nombre, puntos} ordenados.
function Nueva-Serie([string]$Id, [string]$Nombre, $Mapa) {
    $puntos = New-Object System.Collections.Generic.List[object]
    foreach ($k in ($Mapa.Keys | Sort-Object)) { $puntos.Add(@($k, [Math]::Round([double]$Mapa[$k], 3))) }
    return [ordered]@{ id = $Id; nombre = $Nombre; puntos = $puntos.ToArray() }
}

# Agrupa filas en series segun el id que devuelve $Clasificar (o $null para descartar la fila).
function Agrupar($Filas, [scriptblock]$Clasificar, [string]$CampoValor = 'valor', [string]$CampoPeriodo = 'periodo') {
    $series = @{}
    foreach ($f in $Filas) {
        $v = Numero $f.$CampoValor
        if ($null -eq $v) { continue }
        $id = & $Clasificar $f
        if (-not $id) { continue }
        if (-not $series.ContainsKey($id)) { $series[$id] = @{} }
        $series[$id][(Fecha "$($f.$CampoPeriodo)")] = $v
    }
    return $series
}

function A-Lista($Series, $Nombres) {
    $lista = New-Object System.Collections.Generic.List[object]
    foreach ($id in ($Series.Keys | Sort-Object)) {
        $n = $id
        if ($Nombres -and $Nombres.Contains($id)) { $n = $Nombres[$id] }
        $lista.Add((Nueva-Serie $id $n $Series[$id]))
    }
    return , $lista.ToArray()
}

# Si falta una serie imprescindible se corta: nunca se publican datos a medias.
function Exigir($Series, [string[]]$Ids, [string]$Tabla, $Filas, [string[]]$Columnas) {
    foreach ($id in $Ids) {
        if (-not $Series.ContainsKey($id)) {
            $pista = foreach ($c in $Columnas) { "$c = " + ((@($Filas | ForEach-Object { "$($_.$c)" } | Sort-Object -Unique) | Select-Object -First 25) -join ' | ') }
            throw "No se encontro la serie '$id' en $Tabla. Valores recibidos:`n$($pista -join "`n")"
        }
    }
}

function Es-Indice($f) {
    $t = "$($f.tipo_de_indicador) $($f.tipo_de_indicador_descripcion)"
    return ($t -match '(?i)[ií]ndice|^\s*i\b|nivel') -and ($t -notmatch '(?i)variaci|var_|incidencia|peso|precio promedio')
}

$Indicadores = [ordered]@{}

Write-Host 'Descargando EMAE...'
$filas = Get-Filas 'estimador-mensual-de-actividad-economica-nivel-general' 'tipo_de_indicador_descripcion,anio,mes,periodo,tipo_de_indicador'
$s = Agrupar ($filas | Where-Object { Es-Indice $_ }) {
    param($f) $t = "$($f.tipo_de_indicador) $($f.tipo_de_indicador_descripcion)"
    if ($t -match '(?i)desest') { 'desest' } elseif ($t -match '(?i)tend') { 'tendencia' } else { 'original' }
}
Exigir $s @('original', 'desest', 'tendencia') 'EMAE' $filas @('tipo_de_indicador', 'tipo_de_indicador_descripcion')
$Indicadores.emae = A-Lista $s @{ original = 'Serie original'; desest = 'Desestacionalizada'; tendencia = 'Tendencia-ciclo' }

$filas = Get-Filas 'estimador-mensual-de-actividad-economica-sectores' 'anio,mes,codigo,descripcion,periodo,tipo_de_indicador'
$s = Agrupar ($filas | Where-Object { Es-Indice $_ }) { param($f) "$($f.descripcion)".Trim() }
if ($s.Count -lt 5) { Exigir $s @('sectores') 'EMAE sectores' $filas @('tipo_de_indicador', 'descripcion') }
$Indicadores.emae_sectores = A-Lista $s $null

Write-Host 'Descargando IPC...'
$filtro = 'region=eq.Nacional&tipo_de_indicador=eq.Indice&nivel=in.' + [uri]::EscapeDataString('(Nivel General,División,Categoría)')
$filas = Get-Filas 'ipc_nacional' 'clasificador,nivel,codigo,descripcion,periodo,anio,mes,indicador,tipo_de_indicador,tipo_de_indicador_descripcion,region,unidad_de_medida' $filtro
$s = Agrupar $filas { param($f) if ($f.nivel -eq 'Nivel General') { 'general' } }
Exigir $s @('general') 'IPC' $filas @('tipo_de_indicador', 'clasificador', 'nivel', 'region')
$Indicadores.ipc = A-Lista $s @{ general = 'IPC nivel general' }
$s = Agrupar $filas { param($f) if ($f.nivel -eq 'División') { "$($f.descripcion)".Trim() } }
if ($s.Count -lt 12) { Exigir $s @('12 divisiones') 'IPC' $filas @('nivel', 'descripcion') }
$Indicadores.ipc_divisiones = A-Lista $s $null
$s = Agrupar $filas {
    param($f) if ($f.nivel -eq 'Categoría') { switch -Regex ("$($f.descripcion)") { '^Núcleo' { 'nucleo' } '^Regulados' { 'regulados' } '^Estacional' { 'estacionales' } } }
}
Exigir $s @('nucleo', 'regulados', 'estacionales') 'IPC' $filas @('nivel', 'descripcion')
$Indicadores.ipc_categorias = A-Lista $s @{ nucleo = 'Núcleo'; regulados = 'Regulados'; estacionales = 'Estacionales' }

Write-Host 'Descargando salarios...'
$filas = Get-Filas 'indice_salarios' 'clasificador,nivel,codigo,descripcion,periodo,anio,mes,indicador,tipo_de_indicador,tipo_de_indicador_descripcion,unidad_de_medida' 'tipo_de_indicador=eq.Indice'
$cods = @{ '0' = 'total'; '110' = 'privado_registrado'; '120' = 'publico'; '200' = 'no_registrado' }
$s = Agrupar $filas { param($f) $cods["$($f.codigo)"] }
Exigir $s @('total', 'privado_registrado', 'publico', 'no_registrado') 'Salarios' $filas @('codigo', 'descripcion')
$Indicadores.salarios = A-Lista $s @{ total = 'Total'; privado_registrado = 'Privado registrado'; publico = 'Público'; no_registrado = 'Privado no registrado' }

Write-Host 'Descargando ICC...'
$filtro = 'tipo_de_indicador=eq.Indice&nivel=in.' + [uri]::EscapeDataString('(Nivel general,Capitulo)')
$filas = Get-Filas 'icc_precios_construccion' 'indicador,codigo_apertura,codigo_CPC_1,clasificador,nivel,descripcion,periodo,anio,mes,tipo_de_indicador,tipo_de_indicador_descripcion,region,unidad_de_medida' $filtro
$s = Agrupar $filas {
    param($f) switch ("$($f.descripcion)".Trim()) { 'Nivel general' { 'general' } 'Materiales' { 'materiales' } 'Mano de obra' { 'mano_obra' } 'Gastos generales' { 'gastos' } }
}
Exigir $s @('general', 'materiales', 'mano_obra', 'gastos') 'ICC' $filas @('nivel', 'descripcion')
$Indicadores.icc = A-Lista $s @{ general = 'Nivel general'; materiales = 'Materiales'; mano_obra = 'Mano de obra'; gastos = 'Gastos generales' }

Write-Host 'Descargando precios mayoristas...'
$filtro = 'tipo_de_indicador=eq.Indice&nivel=in.' + [uri]::EscapeDataString('(Nivel general,Origen)')
$filas = Get-Filas 'sipm_precios_mayoristas' 'operativo,indicador,clasificador,nivel,codigo_apertura,descripcion,periodo,anio,mes,tipo_de_indicador,tipo_de_indicador_descripcion,region,unidad_de_medida' $filtro
$s = Agrupar $filas {
    param($f) $i = "$($f.indicador)"; $d = "$($f.descripcion)"
    if ($f.nivel -eq 'Nivel general') { $i.ToLower() }
    elseif ($i -eq 'IPIM' -and $d -match '(?i)nacional') { 'ipim_nacionales' }
    elseif ($i -eq 'IPIM' -and $d -match '(?i)importad') { 'ipim_importados' }
}
Exigir $s @('ipim', 'ipib', 'ipp', 'ipim_nacionales', 'ipim_importados') 'SIPM' $filas @('indicador', 'nivel', 'descripcion')
$Indicadores.sipm = A-Lista $s @{ ipim = 'IPIM'; ipib = 'IPIB'; ipp = 'IPP'; ipim_nacionales = 'IPIM productos nacionales'; ipim_importados = 'IPIM productos importados' }

Write-Host 'Descargando mercado de trabajo...'
$filas = Get-Filas 'mercado_trabajo_eph' 'area_geografica_nom,periodo'
$tasas = [ordered]@{
    tasa_de_actividad = 'Actividad'; tasa_de_empleo = 'Empleo'; tasa_de_empleo_formal = 'Empleo formal'; tasa_de_empleo_informal = 'Empleo informal'
    tasa_de_desocupacion = 'Desocupación'; tasa_de_subocupacion = 'Subocupación'; tasa_de_ocupacion_demandante = 'Ocupados demandantes'
}
$s = @{}
foreach ($f in $filas) {
    if ("$($f.area_geografica_nom)" -notmatch '31') { continue }
    # periodo AAAANT (ej. 20162T): se ubica el trimestre en su ultimo mes para compartir eje con las series mensuales.
    $per = "$($f.periodo)"
    $clave = '{0}-{1:00}' -f $per.Substring(0, 4), ([int]$per.Substring(4, 1) * 3)
    foreach ($c in $tasas.Keys) {
        $v = Numero $f.$c
        if ($null -eq $v) { continue }
        if (-not $s.ContainsKey($c)) { $s[$c] = @{} }
        $s[$c][$clave] = $v
    }
}
Exigir $s @('tasa_de_desocupacion') 'EPH' $filas @('area_geografica_nom', 'periodo')
$Indicadores.trabajo = A-Lista $s $tasas

# ---------------------------------------------------------------- Comercio exterior

Write-Host 'Descargando nomenclaturas de comercio exterior...'
$nom = @{}
foreach ($f in (Get-Filas 'comercio_exterior_nomenclaturas' 'nombre_columna,codigo')) {
    $nom["$($f.nombre_columna)".ToLower() + '|' + "$($f.codigo)"] = "$($f.codigo_descripcion)".Trim()
}
function Nombre-De([string]$Columna, [string]$Codigo) {
    $k = $Columna.ToLower() + '|' + $Codigo
    if ($nom.ContainsKey($k)) { $nom[$k] } else { $Codigo }
}

$PrevPath = Join-Path $DirDatos 'datos.json'
$Prev = $null
if ((Test-Path $PrevPath) -and -not $Completo) { $Prev = [IO.File]::ReadAllText($PrevPath) | ConvertFrom-Json }

# Las tablas de comercio tienen una fila por mes/rubro/subrubro/pais: se agregan aca y solo se guarda lo que usa la pagina.
function Bajar-Comercio([string]$Tabla, [string]$Grupo, [string]$CampoGrupo, [string]$CampoGrupo2, [string]$CampoValor, [string]$ClavePrev) {
    $desde = $AnioInicioComercio
    $porGrupo = @{}
    # Modo incremental: reutiliza lo ya descargado y refresca los ultimos 2 anios por revisiones.
    if ($Prev -and $Prev.indicadores.$ClavePrev) {
        $desde = (Get-Date).Year - 2
        foreach ($serie in $Prev.indicadores.$ClavePrev) {
            $porGrupo[$serie.id] = @{}
            foreach ($p in $serie.puntos) { if ([int]$p[0].Substring(0, 4) -lt $desde) { $porGrupo[$serie.id][$p[0]] = [double]$p[1] } }
        }
    }
    $filtro = "Anio=gte.$desde&select=Periodo,$CampoGrupo,pais_cod,$CampoValor"
    $orden = "Anio,Mes,Periodo,$CampoGrupo,$CampoGrupo2,pais_cod"
    $filas = Get-Filas $Tabla $orden $filtro
    $ultimos = @{}
    $delMes = @{}
    foreach ($f in $filas) {
        $v = Numero $f.$CampoValor
        if ($null -eq $v) { continue }   # valores suprimidos por secreto estadistico
        $fecha = Fecha "$($f.Periodo)"
        $g = "$($f.$CampoGrupo)"
        if (-not $delMes.ContainsKey($fecha)) { $delMes[$fecha] = @{} }
        if (-not $delMes[$fecha].ContainsKey($g)) { $delMes[$fecha][$g] = 0.0 }
        $delMes[$fecha][$g] += $v
    }
    $meses = @($delMes.Keys | Sort-Object)
    $ventana = @($meses | Select-Object -Last 12)
    foreach ($f in $filas) {
        $fecha = Fecha "$($f.Periodo)"
        if ($ventana -notcontains $fecha) { continue }
        $v = Numero $f.$CampoValor
        if ($null -eq $v) { continue }
        $p = "$($f.pais_cod)"
        if (-not $ultimos.ContainsKey($p)) { $ultimos[$p] = 0.0 }
        $ultimos[$p] += $v
    }
    foreach ($m in $meses) {
        foreach ($g in $delMes[$m].Keys) {
            if (-not $porGrupo.ContainsKey($g)) { $porGrupo[$g] = @{} }
            $porGrupo[$g][$m] = $delMes[$m][$g] / 1e6    # millones de USD
        }
    }
    $nombres = @{}
    foreach ($g in $porGrupo.Keys) { $nombres[$g] = Nombre-De $CampoGrupo $g }
    $Indicadores[$ClavePrev] = A-Lista $porGrupo $nombres
    $paises = $ultimos.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
        [ordered]@{ nombre = (Nombre-De 'pais_cod' $_.Key); valor = [Math]::Round($_.Value / 1e6, 1) }
    }
    $Indicadores[$Grupo] = [ordered]@{ desde = $ventana[0]; hasta = $ventana[-1]; paises = @($paises) }
}

Write-Host 'Descargando exportaciones (la primera vez tarda bastante)...'
Bajar-Comercio 'comercio_exterior_exportacion_bienes' 'expo_destinos' 'rubro_cod' 'subrubro_cod' 'valor_FOB' 'expo'
Write-Host 'Descargando importaciones (la primera vez tarda bastante)...'
Bajar-Comercio 'comercio_exterior_importacion_bienes' 'impo_origenes' 'uso_economico_cod' 'gce_cod' 'valor_CIF' 'impo'

# ---------------------------------------------------------------- Salida

if (-not (Test-Path $DirDatos)) { New-Item -ItemType Directory $DirDatos | Out-Null }
$jsonInd = $Indicadores | ConvertTo-Json -Depth 8 -Compress
if (Test-Path $PrevPath) {
    $previo = [IO.File]::ReadAllText($PrevPath)
    if ($previo.EndsWith('"indicadores":' + $jsonInd + '}')) { Write-Host 'Sin cambios en los datos.'; return }
}
$salida = [ordered]@{ generado = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', $Inv); fuente = 'INDEC, Catálogo API-INDEC'; indicadores = $Indicadores }
$json = $salida | ConvertTo-Json -Depth 8 -Compress
$utf8 = New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText($PrevPath, $json, $utf8)
[IO.File]::WriteAllText((Join-Path $DirDatos 'datos.js'), "window.INDEC_DATA=$json;", $utf8)
Write-Host "Listo: $PrevPath"
