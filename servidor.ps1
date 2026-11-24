# Servidor web minimo para abrir el inventario desde el celular por la red local (LAN).
# - Sirve los archivos de esta carpeta.
# - Recibe los registros de los celulares, los junta en inventario.json y guarda Inventario.xlsx.
# Cerrar la ventana detiene el servidor.
$puerto = 8080
$raiz = [IO.Path]::GetFullPath($PSScriptRoot)
$archivoDatos = Join-Path $raiz 'inventario.json'
$archivoExcel = Join-Path $raiz 'Inventario.xlsx'
$tipos = @{
  '.html' = 'text/html; charset=utf-8'; '.js' = 'application/javascript'; '.css' = 'text/css'
  '.json' = 'application/json'; '.wasm' = 'application/wasm'; '.png' = 'image/png'; '.ico' = 'image/x-icon'
}

function Leer-Peticion($stream) {
  # Encabezados byte por byte hasta la linea vacia; luego el cuerpo segun Content-Length.
  $buf = New-Object System.Collections.Generic.List[byte]
  while ($true) {
    $b = $stream.ReadByte()
    if ($b -lt 0) { break }
    $buf.Add([byte]$b)
    $n = $buf.Count
    if ($n -ge 4 -and $buf[$n-4] -eq 13 -and $buf[$n-3] -eq 10 -and $buf[$n-2] -eq 13 -and $buf[$n-1] -eq 10) { break }
  }
  $lineas = [Text.Encoding]::ASCII.GetString($buf.ToArray()) -split "`r`n"
  $partes = $lineas[0] -split ' '
  $largo = 0
  foreach ($l in $lineas) { if ($l -match '^Content-Length:\s*(\d+)') { $largo = [int]$matches[1] } }
  $cuerpo = New-Object byte[] $largo
  $leido = 0
  while ($leido -lt $largo) {
    $r = $stream.Read($cuerpo, $leido, $largo - $leido)
    if ($r -le 0) { break }
    $leido += $r
  }
  return @{ Metodo = $partes[0]; Ruta = $partes[1]; Cuerpo = $cuerpo }
}

function Responder($stream, [string]$estado, [string]$tipo, [byte[]]$cuerpo) {
  $enc = "HTTP/1.1 $estado`r`nContent-Type: $tipo`r`nContent-Length: $($cuerpo.Length)`r`nCache-Control: no-cache`r`nConnection: close`r`n`r`n"
  $bytes = [Text.Encoding]::ASCII.GetBytes($enc)
  $stream.Write($bytes, 0, $bytes.Length)
  $stream.Write($cuerpo, 0, $cuerpo.Length)
}

function Json($texto) { return [Text.Encoding]::UTF8.GetBytes($texto) }

function Juntar-Registros([string]$jsonNuevos) {
  # Junta por id: un registro editado en el celular reemplaza al anterior.
  $todos = @{}
  if (Test-Path -LiteralPath $archivoDatos) {
    $previos = [IO.File]::ReadAllText($archivoDatos) | ConvertFrom-Json
    foreach ($r in $previos) { if ($r.id) { $todos[$r.id] = $r } }
  }
  $nuevos = $jsonNuevos | ConvertFrom-Json
  foreach ($r in $nuevos) { if ($r.id) { $todos[$r.id] = $r } }
  $lista = @($todos.Values | Sort-Object id)
  $salida = ConvertTo-Json -InputObject $lista -Depth 5 -Compress
  if ($lista.Count -eq 0) { $salida = '[]' }
  [IO.File]::WriteAllText($archivoDatos, $salida, (New-Object Text.UTF8Encoding $false))
  return @{ Json = $salida; Total = $lista.Count; Nuevos = @($nuevos).Count }
}

try {
  $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, $puerto)
  $listener.Start()
} catch {
  Write-Host "No se pudo abrir el puerto $puerto. Puede que otro programa lo este usando." -ForegroundColor Red
  Read-Host "Presiona Enter para salir"; exit 1
}

Write-Host ""
Write-Host "  Servidor de inventario funcionando." -ForegroundColor Green
Write-Host "  Abre en el navegador del celular (conectado al mismo Wi-Fi) una de estas direcciones:"
Write-Host ""
Get-NetIPAddress -AddressFamily IPv4 |
  Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } |
  ForEach-Object { Write-Host "     http://$($_.IPAddress):$puerto" -ForegroundColor Cyan }
Write-Host ""
Write-Host "  Los registros enviados se guardan en: $archivoExcel"
Write-Host "  Deja esta ventana abierta mientras se captura. Cierrala para apagar el servidor."
Write-Host ""

while ($true) {
  $cliente = $listener.AcceptTcpClient()
  $cliente.ReceiveTimeout = 15000
  try {
    $stream = $cliente.GetStream()
    $p = Leer-Peticion $stream
    if (-not $p.Ruta) { continue }
    $ruta = [Uri]::UnescapeDataString(($p.Ruta -split '\?')[0])
    $hora = Get-Date -Format 'HH:mm:ss'

    if ($p.Metodo -eq 'POST' -and $ruta -eq '/api/enviar') {
      try {
        $res = Juntar-Registros ([Text.Encoding]::UTF8.GetString($p.Cuerpo))
        Responder $stream '200 OK' 'application/json; charset=utf-8' (Json $res.Json)
        Write-Host "  $hora  Recibidos $($res.Nuevos) registros de $($cliente.Client.RemoteEndPoint.Address). Total: $($res.Total)" -ForegroundColor Green
      } catch {
        Responder $stream '500 Internal Server Error' 'application/json' (Json '{"error":"Error al guardar los registros en la PC"}')
        Write-Host "  $hora  ERROR al juntar registros: $_" -ForegroundColor Red
      }
    }
    elseif ($p.Metodo -eq 'POST' -and $ruta -eq '/api/excel') {
      try {
        $bytes = [Convert]::FromBase64String([Text.Encoding]::ASCII.GetString($p.Cuerpo))
        [IO.File]::WriteAllBytes($archivoExcel, $bytes)
        Responder $stream '200 OK' 'application/json' (Json '{"archivo":"Inventario.xlsx"}')
        Write-Host "  $hora  Inventario.xlsx actualizado" -ForegroundColor Green
      } catch {
        $msg = '{"error":"Los registros SI llegaron a la PC, pero no se pudo actualizar Inventario.xlsx. Si esta abierto en Excel, cierralo y vuelve a enviar."}'
        Responder $stream '423 Locked' 'application/json' (Json $msg)
        Write-Host "  $hora  No se pudo escribir Inventario.xlsx (esta abierto en Excel?)" -ForegroundColor Yellow
      }
    }
    else {
      if ($ruta -eq '/') { $ruta = '/index.html' }
      $archivo = [IO.Path]::GetFullPath((Join-Path $raiz $ruta.TrimStart('/')))
      if ($archivo.StartsWith($raiz, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $archivo -PathType Leaf)) {
        $tipo = $tipos[[IO.Path]::GetExtension($archivo).ToLower()]
        if (-not $tipo) { $tipo = 'application/octet-stream' }
        Responder $stream '200 OK' $tipo ([IO.File]::ReadAllBytes($archivo))
      } else {
        Responder $stream '404 Not Found' 'text/plain; charset=utf-8' (Json 'No encontrado')
      }
    }
  } catch {
    # Conexion interrumpida por el celular; se ignora.
  } finally {
    $cliente.Close()
  }
}
