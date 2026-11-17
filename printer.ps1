# =====================================================================
#  printer.ps1 - migra portas de impressoras para o NOME DO NÓ,
#  identificando cada aparelho por número de série / MAC (SNMP + ARP)
# =====================================================================

# ===================== CONFIGURAÇÃO (edite se quiser) =====================
$Simular            = $false   # $true = só mostra o que faria, sem alterar impressoras
$RemoverPortaAntiga = $true    # remove a porta antiga se nenhuma impressora a usar mais
$FiltroNome         = '*'      # ex.: '*Brother*' para migrar só algumas impressoras
$IgnorarFax         = $true    # ignora impressoras/drivers de fax (ex.: Brother PC-FAX)
$Forcar             = $false   # $true = migra mesmo sem confirmar nome/resposta (não recomendado)
$BuscarNaRede       = $true    # procura a impressora na rede local quando o IP antigo não vale mais
$VerificarDepois    = $true    # testa a impressora após a troca e REVERTE se não responder / for outro aparelho
$GravarInventario   = $true    # grava série/MAC/nome de cada impressora em ProgramData (mesmo em simulação)
$ModoAssistido      = $true    # pergunta no terminal quando não há como identificar sozinho (só interativo e sem simulação)
$RedesExtras        = @()      # sub-redes /24 extras para varrer, ex.: @('192.168.21','192.168.3')
$Comunidades        = @('public')   # comunidades SNMP a tentar (as das portas já são incluídas)

# Mapa manual (opcional): nome da fila no Windows = nome do nó. Exemplo:
# $Manual = @{ 'Tributos 1' = 'BRWBCF4D443678E'; 'Tributos 2' = 'SELB-8634' }
$Manual = @{ }
# ==========================================================================

$ErrorActionPreference = 'Stop'

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { Write-Host "Abra o PowerShell como ADMINISTRADOR e rode de novo." -ForegroundColor Red; return }

# ---------------------------------------------------------------- Funções ----

function Test-EhIPv4 { param([string]$Texto); return [bool]($Texto -match '^\d{1,3}(\.\d{1,3}){3}$') }

function Get-NomeNetBios {
    param([string]$IP)
    $linha = nbtstat -A $IP 2>$null | Select-String '<00>\s+UNIQUE' | Select-Object -First 1
    if ($linha -and $linha.ToString() -match '^\s*(.+?)\s*<00>\s+UNIQUE') { return $Matches[1].Trim() }
    return $null
}

function Get-NomeDoNoPorIP {
    param([string]$IP)
    $dns = (Resolve-DnsName -Name $IP -Type PTR -ErrorAction SilentlyContinue).NameHost
    if ($dns) { return ($dns | Select-Object -First 1).Split('.')[0] }
    return (Get-NomeNetBios -IP $IP)
}

function Get-IPdoNome {
    param([string]$Nome)
    if (-not $Nome) { return $null }
    try {
        $a = [System.Net.Dns]::GetHostAddresses($Nome) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1
        if ($a) { return $a.IPAddressToString }
    } catch { }
    try {
        $r = Resolve-DnsName -Name $Nome -NetbiosFallback -ErrorAction Stop | Where-Object { $_.IPAddress } | Select-Object -First 1
        if ($r) { return $r.IPAddress }
    } catch { }
    return $null
}

function Test-Porta {
    param([string]$Alvo, [int]$Porta = 9100, [int]$TimeoutMs = 1500)
    $c = New-Object System.Net.Sockets.TcpClient
    try { $t = $c.ConnectAsync($Alvo, $Porta); return ($t.Wait($TimeoutMs) -and $c.Connected) }
    catch { return $false }
    finally { $c.Dispose() }
}

function Test-ImpressoraResponde {
    param([string]$Nome, [int]$PortaTcp)
    for ($i = 0; $i -lt 3; $i++) {
        if (Test-Porta -Alvo $Nome -Porta $PortaTcp -TimeoutMs 3000) { return $true }
        Start-Sleep -Seconds 2
    }
    return $false
}

# MAC (sem separadores, maiúsculo) lido da tabela ARP; só existe na mesma sub-rede
function Get-MacPorIP {
    param([string]$IP)
    $n = Get-NetNeighbor -IPAddress $IP -AddressFamily IPv4 -ErrorAction SilentlyContinue |
         Where-Object { $_.LinkLayerAddress -and $_.LinkLayerAddress -notmatch '^(00-00-00-00-00-00|FF-FF-FF-FF-FF-FF)$' } |
         Select-Object -First 1
    if ($n) { return ($n.LinkLayerAddress -replace '[-:]', '').ToUpper() }
    return $null
}

function Get-PaginaWeb {
    param([string]$IP)
    try { return [string](Invoke-WebRequest -Uri "http://$IP" -UseBasicParsing -TimeoutSec 3 -ErrorAction Stop).Content }
    catch { return $null }
}

$MarcasRegex = [ordered]@{
    Brother = 'brother'; Samsung = 'samsung|syncthru'; Canon = 'canon'; HP = '\bHP\b|hewlett'
    Epson = 'epson'; Ricoh = 'ricoh|lanier|savin'; Kyocera = 'kyocera|ecosys'; Xerox = 'xerox'
    Lexmark = 'lexmark'; OKI = '\bOKI\b'; Konica = 'konica|minolta|bizhub'; Sharp = '\bsharp\b'
}

function Get-MarcaPorTexto {
    param([string]$Texto)
    if (-not $Texto) { return $null }
    foreach ($k in $MarcasRegex.Keys) { if ($Texto -match $MarcasRegex[$k]) { return $k } }
    return $null
}

# Marca da fila (pelo driver + nome), mesmo que a fila tenha sido renomeada
function Get-MarcaImpressora {
    param($Imp)
    return (Get-MarcaPorTexto -Texto "$($Imp.DriverName) $($Imp.Name)")
}

# $true = modelo do driver aparece no texto do aparelho; $false = não aparece; $null = não dá para saber
function Test-ModeloCombina {
    param([string]$Driver, [string]$TextoNorm)
    if (-not $Driver -or -not $TextoNorm) { return $null }
    $tokens = @([regex]::Matches($Driver, '[A-Za-z0-9\-]*\d[A-Za-z0-9\-]*') | ForEach-Object { $_.Value } |
        Where-Object { $_.Length -ge 3 -and $_ -notmatch '^(PCL\d*|PS\d*|XPS|v\d+)$' })
    if ($tokens.Count -eq 0) { return $null }
    foreach ($tok in $tokens) {
        $t = ($tok -replace '[^A-Za-z0-9]', '')
        if ($tok -cmatch 'x$') { $re = [regex]::Escape($t.Substring(0, $t.Length - 1).ToUpper()) + '[A-Z0-9]' }   # "M337x" -> M337?
        else { $re = [regex]::Escape($t.ToUpper()) }
        if ($TextoNorm -match $re) { return $true }
    }
    return $false
}

# ------------------------------------------------------------------ SNMP ----

function Read-Tlv {
    param([byte[]]$B, [int]$P)
    $tag = $B[$P]; $l = [int]$B[$P + 1]; $s = $P + 2
    if ($l -band 0x80) {
        $n = $l -band 0x7F; $l = 0
        for ($i = 0; $i -lt $n; $i++) { $l = ($l -shl 8) + $B[$P + 2 + $i] }
        $s = $P + 2 + $n
    }
    [PSCustomObject]@{ Tag = $tag; Len = $l; Start = $s; End = $s + $l }
}

function New-Tlv {
    param([byte]$Tag, [byte[]]$Val)
    $len = $Val.Length
    $hdr = if ($len -lt 128) { [byte[]]@($Tag, $len) } else { [byte[]]@($Tag, 0x81, $len) }
    return [byte[]]($hdr + $Val)
}

# SNMP v1 GET de um OID. Devolve texto/número, ou $null se não houver resposta.
function Get-Snmp {
    param([string]$IP, [string]$Oid, [string]$Comunidade = 'public', [int]$TimeoutMs = 1500)
    try {
        $arcs = @($Oid.Split('.') | ForEach-Object { [int]$_ })
        $ob = New-Object System.Collections.Generic.List[byte]
        $ob.Add([byte](40 * $arcs[0] + $arcs[1]))
        foreach ($a in $arcs[2..($arcs.Count - 1)]) {
            if ($a -lt 128) { $ob.Add([byte]$a) }
            else {
                $partes = @(); $x = $a
                $partes += ($x -band 0x7F); $x = $x -shr 7
                while ($x -gt 0) { $partes += (($x -band 0x7F) -bor 0x80); $x = $x -shr 7 }
                [array]::Reverse($partes)
                foreach ($p in $partes) { $ob.Add([byte]$p) }
            }
        }
        $vb  = New-Tlv 0x30 ([byte[]]((New-Tlv 0x06 $ob.ToArray()) + (New-Tlv 0x05 ([byte[]]@()))))
        $vbl = New-Tlv 0x30 $vb
        $pduCorpo = [byte[]]((New-Tlv 0x02 ([byte[]]@(1))) + (New-Tlv 0x02 ([byte[]]@(0))) + (New-Tlv 0x02 ([byte[]]@(0))) + $vbl)
        $pdu = New-Tlv 0xA0 $pduCorpo
        $msgCorpo = [byte[]]((New-Tlv 0x02 ([byte[]]@(0))) + (New-Tlv 0x04 ([System.Text.Encoding]::ASCII.GetBytes($Comunidade))) + $pdu)
        $pkt = New-Tlv 0x30 $msgCorpo

        $u = New-Object System.Net.Sockets.UdpClient
        try {
            $u.Client.ReceiveTimeout = $TimeoutMs
            $u.Connect($IP, 161)
            [void]$u.Send($pkt, $pkt.Length)
            $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
            $r = $u.Receive([ref]$ep)
        } finally { $u.Close() }

        $t0 = Read-Tlv $r 0;         $t1 = Read-Tlv $r $t0.Start
        $t2 = Read-Tlv $r $t1.End;   $t3 = Read-Tlv $r $t2.End
        $e1 = Read-Tlv $r $t3.Start; $e2 = Read-Tlv $r $e1.End
        if ($r[$e2.Start] -ne 0) { return $null }
        $e3 = Read-Tlv $r $e2.End;   $l1 = Read-Tlv $r $e3.End
        $l2 = Read-Tlv $r $l1.Start; $o  = Read-Tlv $r $l2.Start
        $v  = Read-Tlv $r $o.End
        if ($v.Tag -eq 0x04) { return ([System.Text.Encoding]::ASCII.GetString($r, $v.Start, $v.Len)).Trim() }
        if ($v.Tag -eq 0x02) { $n = 0; for ($i = 0; $i -lt $v.Len; $i++) { $n = ($n -shl 8) + $r[$v.Start + $i] }; return $n }
        return $null
    } catch { return $null }
}

# ------------------------------------------------------------ Identidade ----

# Lê tudo o que identifica o aparelho que responde em $IP (SNMP, ARP, nome, marca)
function Get-Identidade {
    param([string]$IP, [string[]]$Coms)
    $id = [PSCustomObject]@{
        IP = $IP; Vivo = $false; Snmp = $false; Serie = $null; Modelo = $null; SysName = $null
        Descr = $null; Mac = $null; Nome = $null; Marca = $null; Texto = $null
    }
    $com = $null
    foreach ($c in $Coms) {
        $d = Get-Snmp -IP $IP -Oid '1.3.6.1.2.1.1.1.0' -Comunidade $c
        if ($d) { $id.Descr = [string]$d; $id.Snmp = $true; $com = $c; break }
    }
    if ($id.Snmp) {
        $id.Vivo = $true
        $s = Get-Snmp -IP $IP -Oid '1.3.6.1.2.1.43.5.1.1.17.1' -Comunidade $com
        if (-not $s -and $id.Descr -match 'S/N\s+([A-Za-z0-9]+)') { $s = $Matches[1] }
        if ($s -and ([string]$s).Trim().Length -ge 4) { $id.Serie = ([string]$s).Trim().ToUpper() }
        $id.Modelo  = Get-Snmp -IP $IP -Oid '1.3.6.1.2.1.25.3.2.1.3.1' -Comunidade $com
        $id.SysName = Get-Snmp -IP $IP -Oid '1.3.6.1.2.1.1.5.0' -Comunidade $com
    }
    else {
        # Sem SNMP: considera "vivo" se alguma porta de impressão abrir
        foreach ($pt in 9100, 515, 631) { if (Test-Porta -Alvo $IP -Porta $pt -TimeoutMs 800) { $id.Vivo = $true; break } }
    }
    if (-not $id.Vivo) { return $id }

    $id.Mac  = Get-MacPorIP -IP $IP
    $id.Nome = Get-NomeDoNoPorIP -IP $IP
    $txt = "$($id.Modelo) $($id.Descr)"
    if (-not $id.Snmp) { $web = Get-PaginaWeb -IP $IP; if ($web) { $txt = "$txt $web" } }
    $id.Marca = Get-MarcaPorTexto -Texto $txt
    $n = ($txt -replace '[^A-Za-z0-9]', '').ToUpper()
    if ($n.Length -gt 40000) { $n = $n.Substring(0, 40000) }
    $id.Texto = $n
    return $id
}

function Get-IdentidadeCache {
    param([string]$IP)
    if (-not $cacheIdent.ContainsKey($IP)) { $cacheIdent[$IP] = Get-Identidade -IP $IP -Coms $comunidadesTodas }
    return $cacheIdent[$IP]
}

# Escolhe um nome do aparelho que RESOLVA para o IP dele (ou para um aparelho com mesma série/MAC)
function Select-NomeValido {
    param($Dev, [string[]]$Extras = @())
    $lista = @()
    foreach ($c in (@($Dev.Nome, $Dev.SysName) + @($Extras))) { if ($c) { $lista += ([string]$c).Split('.')[0] } }
    $lista = @($lista | Select-Object -Unique)
    foreach ($c in $lista) {
        $ipc = Get-IPdoNome -Nome $c
        if (-not $ipc) { continue }
        if ($ipc -eq $Dev.IP) { return $c }
        $x = Get-IdentidadeCache -IP $ipc
        if ($x.Serie -and $Dev.Serie -and $x.Serie -eq $Dev.Serie) { return $c }
        if ($x.Mac -and $Dev.Mac -and $x.Mac -eq $Dev.Mac) { return $c }
    }
    if ($Forcar -and $lista.Count -gt 0) { return $lista[0] }
    return $null
}

# Acha na lista de aparelhos o que corresponde ao registro do inventário (série > MAC > nome)
function Find-NoInventario {
    param($Inv, $Lista)
    if ($Inv.Serie) {
        $d = @($Lista | Where-Object { $_.Serie -eq $Inv.Serie }) | Select-Object -First 1
        if ($d) { return [PSCustomObject]@{ Dev = $d; Criterio = 'número de série' } }
    }
    if ($Inv.Mac) {
        $d = @($Lista | Where-Object { $_.Mac -eq $Inv.Mac -and -not ($_.Serie -and $Inv.Serie -and $_.Serie -ne $Inv.Serie) }) | Select-Object -First 1
        if ($d) { return [PSCustomObject]@{ Dev = $d; Criterio = 'MAC' } }
    }
    if ($Inv.Nome) {
        $d = @($Lista | Where-Object { $_.Nome -and $_.Nome -eq $Inv.Nome -and -not ($_.Serie -and $Inv.Serie -and $_.Serie -ne $Inv.Serie) }) | Select-Object -First 1
        if ($d) { return [PSCustomObject]@{ Dev = $d; Criterio = 'nome do nó' } }
    }
    return $null
}

# Devolve o nome da fila que já usa este nó com OUTRO aparelho
function Test-ConflitoNo {
    param([string]$Nome, [string]$IdIP, [string]$IdMac, [string]$IdSerie)
    $r = $nomesAtribuidos[$Nome.ToUpper()]
    if (-not $r) { return $null }
    if ($IdSerie -and $r.Serie) { if ($IdSerie -ne $r.Serie) { return $r.Fila } else { return $null } }
    if ($IdIP -and $r.IP -and $IdIP -ne $r.IP) {
        if ($IdMac -and $r.Mac -and $IdMac -eq $r.Mac) { return $null }
        return $r.Fila
    }
    return $null
}

# Varre as sub-redes locais (/24) atrás de aparelhos com porta de impressão aberta e lê a identidade de cada um
function Get-ImpressorasNaRede {
    Write-Host "Procurando impressoras na rede local (pode levar alguns segundos)..." -ForegroundColor Cyan

    $bases = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.IPAddress -notmatch '^(127\.|169\.254\.)' -and
            $_.InterfaceAlias -notmatch 'vEthernet|VirtualBox|VMware|Loopback|Bluetooth|Hyper-V|VPN|WSL' -and
            $_.AddressState -eq 'Preferred'
        } | ForEach-Object { ($_.IPAddress -split '\.')[0..2] -join '.' })
    $bases = @(($bases + @($RedesExtras)) | Where-Object { $_ } | Select-Object -Unique)

    $tentativas = New-Object System.Collections.Generic.List[object]
    foreach ($b in $bases) {
        foreach ($i in 1..254) {
            foreach ($pt in 9100, 515, 631) {
                $ip = "$b.$i"
                $c  = New-Object System.Net.Sockets.TcpClient
                $tentativas.Add([PSCustomObject]@{ IP = $ip; Client = $c; Task = $c.ConnectAsync($ip, $pt) })
            }
        }
    }
    try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]@($tentativas | ForEach-Object { $_.Task }), 4000) } catch { }

    $abertos = @($tentativas | Where-Object { $_.Task.Status -eq 'RanToCompletion' -and $_.Client.Connected } |
        ForEach-Object { $_.IP } | Select-Object -Unique)
    foreach ($t in $tentativas) { try { $t.Client.Dispose() } catch { } }

    $achadas = New-Object System.Collections.Generic.List[object]
    foreach ($ip in $abertos) { $achadas.Add((Get-IdentidadeCache -IP $ip)) }

    if ($achadas.Count -gt 0) {
        Write-Host "Dispositivos de impressão encontrados:" -ForegroundColor Cyan
        $achadas | Format-Table IP, Nome, Marca, Modelo, Serie, Mac -AutoSize | Out-String | Write-Host
    } else { Write-Host "Nenhum dispositivo de impressão encontrado." -ForegroundColor Yellow }
    return $achadas
}

# --------------------------------------------------------- Modo assistido ----

function Send-PaginaIdentificacao {
    param([string]$IP, [string]$Texto)
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        if (-not $c.ConnectAsync($IP, 9100).Wait(2000)) { return $false }
        $bytes = [System.Text.Encoding]::ASCII.GetBytes("`r`n`r`n$Texto`r`n`r`n`f")
        $s = $c.GetStream(); $s.Write($bytes, 0, $bytes.Length); $s.Flush()
        return $true
    } catch { return $false }
    finally { $c.Dispose() }
}

# Lista candidatos e deixa você escolher (ou imprimir uma página de identificação no candidato)
function Invoke-Assistido {
    param($Fila, $Cands)
    $Cands = @($Cands)
    if ($Cands.Count -eq 0) { return $null }
    Write-Host ""
    Write-Host "A fila '$($Fila.Imp.Name)' (driver: $($Fila.Imp.DriverName)) não tem registro de identidade. Candidatos na rede:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $Cands.Count; $i++) {
        $c = $Cands[$i]
        Write-Host ("  [{0}] {1,-15} {2,-18} {3,-30} S/N {4}  MAC {5}" -f ($i + 1), $c.IP, $c.Nome, $c.Modelo, $c.Serie, $c.Mac)
    }
    try {
        while ($true) {
            $r = Read-Host "Número para escolher | 'p N' imprime página de identificação no candidato N | Enter pula"
            if (-not $r) { return $null }
            if ($r -match '^[pP]\s*(\d+)$') {
                $n = [int]$Matches[1]
                if ($n -ge 1 -and $n -le $Cands.Count) {
                    $alvo = $Cands[($n - 1)]
                    $ok = Send-PaginaIdentificacao -IP $alvo.IP -Texto "IDENTIFICACAO - candidato $n ($($alvo.IP)) para a fila: $($Fila.Imp.Name)"
                    if ($ok) { Write-Host "  Página enviada ao candidato $n. Veja qual impressora imprimiu." -ForegroundColor Cyan }
                    else { Write-Host "  Não consegui enviar a página ao candidato $n (porta 9100 fechada ou formato não aceito)." -ForegroundColor Yellow }
                } else { Write-Host "  Candidato inválido." -ForegroundColor Yellow }
            }
            elseif ($r -match '^\d+$') {
                $n = [int]$r
                if ($n -ge 1 -and $n -le $Cands.Count) {
                    $alvo = $Cands[($n - 1)]
                    $conf = Read-Host "Associar '$($Fila.Imp.Name)' ao candidato $n ($($alvo.IP), S/N $($alvo.Serie))? (S/N)"
                    if ($conf -match '^[sSyY]') { return $alvo }
                } else { Write-Host "  Candidato inválido." -ForegroundColor Yellow }
            }
            else { Write-Host "  Não entendi. Use um número, 'p N' ou Enter." -ForegroundColor Yellow }
        }
    } catch { return $null }
}

# ------------------------------------------------------------- Inventário ----

$pastaInv = Join-Path $env:ProgramData 'ReinstalarImpressora'
$arqInv   = Join-Path $pastaInv 'inventario.json'
$inventario = @{}

function Read-Inventario {
    if (-not (Test-Path $arqInv)) { return }
    try {
        $lido = Get-Content -Path $arqInv -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($e in @($lido)) { if ($e.Fila) { $inventario[$e.Fila] = $e } }
    } catch { Write-Host "Aviso: não consegui ler o inventário ($arqInv)." -ForegroundColor Yellow }
}

function Update-Inventario {
    param([string]$Fila, [string]$Driver, [string]$Marca, [string]$Nome, [string]$IP, [string]$Mac, [string]$Serie, [string]$Modelo)
    if (-not $Serie -and -not $Mac -and -not $IP) { return }
    $velho = $inventario[$Fila]
    $novo = [PSCustomObject]@{
        Fila = $Fila; Driver = $Driver; Marca = $Marca; Nome = $Nome; IP = $IP; Mac = $Mac
        Serie = $Serie; Modelo = $Modelo; Atualizado = (Get-Date -Format 's')
    }
    if ($velho) {
        foreach ($campo in 'Driver', 'Marca', 'Nome', 'IP', 'Mac', 'Serie', 'Modelo') {
            if (-not $novo.$campo -and $velho.$campo) { $novo.$campo = $velho.$campo }
        }
    }
    $inventario[$Fila] = $novo
}

function Write-Inventario {
    if (-not $GravarInventario -or $inventario.Count -eq 0) { return }
    try {
        if (-not (Test-Path $pastaInv)) { New-Item -ItemType Directory -Path $pastaInv -Force | Out-Null }
        ConvertTo-Json -InputObject @($inventario.Values) -Depth 3 | Set-Content -Path $arqInv -Encoding UTF8
        Write-Host "Inventário salvo em: $arqInv" -ForegroundColor Cyan
    } catch { Write-Host "Aviso: não consegui gravar o inventário: $($_.Exception.Message)" -ForegroundColor Yellow }
}

# --------------------------------------------------- Coleta das filas (1) ----

$wsdPorts = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\WSD Port\Ports\*" -ErrorAction SilentlyContinue |
    ForEach-Object { $wsdPorts[$_.PSChildName] = $_ }

function Get-InfoFila {
    param($Imp)
    $porta = Get-PrinterPort -Name $Imp.PortName -ErrorAction SilentlyContinue
    $reg = Get-ItemProperty -LiteralPath "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports\$($Imp.PortName)" -ErrorAction SilentlyContinue
    $f = [PSCustomObject]@{
        Imp = $Imp; Porta = $porta; Reg = $reg; Tipo = $null; HostAddr = $porta.PrinterHostAddress
        Ip = $null; Nome = $null; Marca = (Get-MarcaImpressora -Imp $Imp)
        Fax = ($IgnorarFax -and ($Imp.Name -match 'fax' -or $Imp.DriverName -match 'fax'))
        Ident = $null; Diverge = $false
    }
    if ($f.HostAddr) {
        if (Test-EhIPv4 $f.HostAddr) { $f.Tipo = 'TCP/IP (IP estático)'; $f.Ip = $f.HostAddr }
        else { $f.Tipo = 'TCP/IP (já usa nome)'; $f.Nome = $f.HostAddr; $f.Ip = Get-IPdoNome -Nome $f.HostAddr }
    }
    elseif ($Imp.PortName -match '^WSD' -or $porta.Description -match 'WSD') {
        $f.Tipo = 'WSD'
        $wsd = $wsdPorts[$Imp.PortName]
        if ($wsd) { $f.Ip = $wsd.IPAddress; $f.Nome = $wsd.HostName }
        if (-not $f.Ip -or -not $f.Nome) {
            $pnp = Get-PnpDeviceProperty -InstanceId "*$($Imp.PortName)*" -ErrorAction SilentlyContinue
            if (-not $f.Ip)   { $f.Ip   = ($pnp | Where-Object KeyName -match 'IPAddress').Data }
            if (-not $f.Nome) { $f.Nome = ($pnp | Where-Object KeyName -match 'HostName').Data }
        }
    }
    elseif (-not $porta -or $Imp.PortName -match '^(PORTPROMPT|FILE|nul|SHRFAX|XPS|PDF|USB|LPT\d|COM\d|TS\d|Microsoft|\\\\\.\\pipe)' -or $porta.Description -match 'Local|Fax|Arquivo|File') {
        $f.Tipo = 'Local/Virtual'
    }
    else { $f.Tipo = 'Monitor proprietário' }
    return $f
}

# ---------------------------------------------------------------- Principal ----

Write-Host ("printer.ps1 | usuário: {0} | computador: {1}" -f (whoami), $env:COMPUTERNAME) -ForegroundColor DarkGray

try { $impressoras = @(Get-Printer -Name $FiltroNome -ErrorAction Stop) }
catch { $impressoras = @(); Write-Host "Não consegui listar as impressoras: $($_.Exception.Message)" -ForegroundColor Red }

if ($impressoras.Count -eq 0) {
    Write-Host "Nenhuma impressora encontrada para este usuário/computador (filtro: '$FiltroNome'). Nada a fazer." -ForegroundColor Yellow
    return
}
Write-Host "Impressoras encontradas: $($impressoras.Count)" -ForegroundColor DarkGray

if ($Simular) { Write-Host "`n*** MODO SIMULAÇÃO: nenhuma impressora será alterada (o inventário é gravado) ***`n" -ForegroundColor Yellow }

Read-Inventario

$comunidadesTodas = @(@($Comunidades) + @(Get-PrinterPort -ErrorAction SilentlyContinue |
    Where-Object { $_.SNMPCommunity } | ForEach-Object { $_.SNMPCommunity }) | Where-Object { $_ } | Select-Object -Unique)
if ($comunidadesTodas.Count -eq 0) { $comunidadesTodas = @('public') }
$cacheIdent = @{}

$filas = @(foreach ($imp in $impressoras) { Get-InfoFila -Imp $imp })

# --- Identidade do aparelho que responde no IP de cada fila ---
foreach ($f in $filas) {
    if ($f.Fax -or $f.Tipo -notin 'TCP/IP (IP estático)', 'TCP/IP (já usa nome)', 'WSD' -or -not $f.Ip) { continue }
    $f.Ident = Get-IdentidadeCache -IP $f.Ip
    $inv = $inventario[$f.Imp.Name]
    if ($f.Ident.Vivo) {
        $motivo = $null
        if ($inv -and $inv.Serie -and $f.Ident.Serie -and $inv.Serie -ne $f.Ident.Serie) { $motivo = "S/N $($f.Ident.Serie) difere do inventário ($($inv.Serie))" }
        elseif ($f.Marca -and $f.Ident.Marca -and $f.Marca -ne $f.Ident.Marca) { $motivo = "marca $($f.Ident.Marca) difere da fila ($($f.Marca))" }
        elseif ($f.Ident.Modelo -and (Test-ModeloCombina -Driver $f.Imp.DriverName -TextoNorm $f.Ident.Texto) -eq $false) { $motivo = "modelo '$($f.Ident.Modelo)' não combina com o driver '$($f.Imp.DriverName)'" }
        if ($motivo) {
            $f.Diverge = $true
            Write-Host "  [$($f.Imp.Name)] o IP $($f.Ip) agora é de OUTRO aparelho ($motivo)." -ForegroundColor DarkYellow
        }
    }
}

# --- Nós já em uso, nós por marca e quantos APARELHOS distintos há por marca ---
$conhecidos = @{}
$nomesAtribuidos = @{}
$grupos = New-Object System.Collections.ArrayList
foreach ($f in $filas) {
    if ($f.Tipo -eq 'TCP/IP (já usa nome)') {
        $nomesAtribuidos[$f.HostAddr.ToUpper()] = @{ IP = $f.Ip; Mac = $f.Ident.Mac; Serie = $f.Ident.Serie; Fila = $f.Imp.Name }
        if ($f.Marca) {
            if (-not $conhecidos.ContainsKey($f.Marca)) { $conhecidos[$f.Marca] = @() }
            if ($conhecidos[$f.Marca] -notcontains $f.HostAddr) { $conhecidos[$f.Marca] += $f.HostAddr }
        }
    }
    if ($f.Marca -and -not $f.Fax -and $f.Tipo -in 'TCP/IP (IP estático)', 'TCP/IP (já usa nome)', 'WSD') {
        $ids = New-Object System.Collections.Generic.HashSet[string]
        if ($f.HostAddr) { [void]$ids.Add($f.HostAddr.ToUpper()) }
        if ($f.Tipo -eq 'WSD') {
            if ($f.Ip)   { [void]$ids.Add(([string]$f.Ip).ToUpper()) }
            if ($f.Nome) { [void]$ids.Add(([string]$f.Nome).ToUpper()) }
            if ($ids.Count -eq 0) { [void]$ids.Add('WSD:' + $f.Imp.PortName.ToUpper()) }
        }
        $e = $inventario[$f.Imp.Name]
        if ($e) { foreach ($v in $e.Nome, $e.IP, $e.Mac, $e.Serie) { if ($v) { [void]$ids.Add(([string]$v).ToUpper()) } } }
        if ($f.Ident -and $f.Ident.Vivo -and -not $f.Diverge) {
            foreach ($v in $f.Ident.Serie, $f.Ident.Mac) { if ($v) { [void]$ids.Add(([string]$v).ToUpper()) } }
        }
        if ($ids.Count -gt 0) {
            $alvo = $null
            foreach ($g in @($grupos)) {
                if ($g.Marca -ne $f.Marca) { continue }
                if (@($ids | Where-Object { $g.Ids.Contains($_) }).Count -gt 0) {
                    if ($null -eq $alvo) { $alvo = $g; $alvo.Ids.UnionWith($ids) }
                    else { $alvo.Ids.UnionWith($g.Ids); $grupos.Remove($g) }
                }
            }
            if ($null -eq $alvo) { [void]$grupos.Add(@{ Marca = $f.Marca; Ids = $ids }) }
        }
    }
}
$qtdDisp = @{}
foreach ($g in $grupos) { $qtdDisp[$g.Marca] = 1 + [int]$qtdDisp[$g.Marca] }
foreach ($k in @($qtdDisp.Keys)) { Write-Host "Aparelhos $k identificados nas filas: $($qtdDisp[$k])" -ForegroundColor DarkGray }

$resultado = New-Object System.Collections.Generic.List[object]
$portasAntigas = New-Object System.Collections.Generic.HashSet[string]
$descobertas = $null

foreach ($f in $filas) {

    $imp = $f.Imp; $m = $f.Marca; $tipo = $f.Tipo
    $ip = $f.Ip; $nome = $null; $dev = $null; $ipAtual = $null
    $status = $null; $novaPorta = $null; $origemNome = $null

    if ($f.Fax) { $status = 'Ignorada: impressora de fax' }
    elseif ($tipo -eq 'Local/Virtual') { $status = 'Ignorada: porta local/virtual' }
    elseif ($tipo -eq 'Monitor proprietário') {
        $status = "Ignorada: porta de monitor proprietário ('$($f.Porta.Description)'), ainda não suportada"
    }
    else {
        # Captura de identidade (inventário) enquanto o aparelho responde e é o mesmo registrado
        if ($f.Ident -and $f.Ident.Vivo -and -not $f.Diverge -and ($f.Ident.Serie -or $f.Ident.Mac)) {
            $nomeCap = if ($f.Nome) { $f.Nome } elseif ($f.Ident.Nome) { $f.Ident.Nome } else { $f.Ident.SysName }
            Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $m -Nome $nomeCap -IP $f.Ip -Mac $f.Ident.Mac -Serie $f.Ident.Serie -Modelo $f.Ident.Modelo
        }
        $inv = $inventario[$imp.Name]

        if ($tipo -eq 'TCP/IP (já usa nome)') { $status = 'Ignorada: já usa nome do nó' }
        else {
            $bloqueioMotivo = $null; $motivoFalha = $null
            $reservaLiberada = ($m -and [int]$qtdDisp[$m] -le 1)

            # 1) Mapa manual
            if ($Manual.ContainsKey($imp.Name)) {
                $nome = $Manual[$imp.Name]; $origemNome = 'manual'
                $ipm = Get-IPdoNome -Nome $nome
                if ($ipm) { $dev = Get-IdentidadeCache -IP $ipm }
            }
            # 2) O aparelho ainda responde no IP da fila e é o registrado: usa um nome DELE que resolva para o IP dele
            elseif ($f.Ident -and $f.Ident.Vivo -and -not $f.Diverge) {
                $dev = $f.Ident
                $nome = Select-NomeValido -Dev $dev -Extras @($inv.Nome)
                if ($nome) { $origemNome = 'aparelho no IP da fila' }
                else {
                    $bloqueioMotivo = "o aparelho em $ip responde, mas nenhum nome dele resolve para o IP dele; use `$Manual"
                    $dev = $null
                }
            }

            # 3) Inventário: acha o aparelho na rede (série > MAC > nome)
            if (-not $nome -and -not $bloqueioMotivo -and $BuscarNaRede -and $inv -and ($inv.Serie -or $inv.Mac -or $inv.Nome)) {
                if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
                $ach = Find-NoInventario -Inv $inv -Lista $descobertas
                if ($ach) {
                    $n3 = Select-NomeValido -Dev $ach.Dev -Extras @($inv.Nome)
                    if ($n3) { $nome = $n3; $dev = $ach.Dev; $origemNome = "inventário: achado na rede por $($ach.Criterio), IP atual $($ach.Dev.IP)" }
                    else { $motivoFalha = "achei o aparelho do inventário em $($ach.Dev.IP), mas nenhum nome dele resolve para esse IP" }
                }
            }

            # 4) Caminhos de reserva (só com UM aparelho dessa marca instalado)
            if (-not $nome -and -not $bloqueioMotivo -and $reservaLiberada) {
                # 4a) outra fila da mesma marca já usa um nó: é o mesmo aparelho
                if ($conhecidos[$m] -and @($conhecidos[$m]).Count -eq 1) {
                    $cand = @($conhecidos[$m])[0]
                    $ipc = Get-IPdoNome -Nome $cand
                    if ($ipc) { $nome = $cand; $origemNome = 'mesmo nó de outra fila'; $dev = Get-IdentidadeCache -IP $ipc }
                }
                # 4b) procura na rede por marca (e modelo, se houver mais de um)
                if (-not $nome -and $BuscarNaRede) {
                    if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
                    $cands = @($descobertas | Where-Object {
                        $_.Marca -eq $m -and -not ($_.Serie -and $inv.Serie -and $_.Serie -ne $inv.Serie) })
                    if ($cands.Count -gt 1) {
                        $porModelo = @($cands | Where-Object { (Test-ModeloCombina -Driver $imp.DriverName -TextoNorm $_.Texto) -eq $true })
                        if ($porModelo.Count -ge 1) { $cands = $porModelo }
                    }
                    if ($cands.Count -eq 1) {
                        $n4 = Select-NomeValido -Dev $cands[0]
                        if ($n4) {
                            $mod = Test-ModeloCombina -Driver $imp.DriverName -TextoNorm $cands[0].Texto
                            $obs = if ($mod -eq $true) { 'modelo confirmado' } elseif ($mod -eq $false) { 'modelo NÃO confirmado' } else { 'modelo não verificável' }
                            $nome = $n4; $dev = $cands[0]; $origemNome = "achada na rede por marca ($obs)"
                        }
                        else { $motivoFalha = "achei uma $m em $($cands[0].IP), mas nenhum nome dela resolve para esse IP" }
                    }
                    elseif ($cands.Count -gt 1) { $motivoFalha = "$($cands.Count) impressoras $m na rede ($(($cands.IP) -join ', ')); sem como escolher sozinho" }
                    else { $motivoFalha = "nenhuma impressora $m encontrada na rede" }
                }
            }

            # 5) Modo assistido: pergunta no terminal (só interativo e fora da simulação)
            if (-not $nome -and -not $bloqueioMotivo -and $ModoAssistido -and $BuscarNaRede -and -not $Simular -and $Host.Name -match 'ConsoleHost|ISE') {
                if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
                $cands5 = @($descobertas | Where-Object { ($_.Snmp -or $_.Marca) -and (-not $m -or -not $_.Marca -or $_.Marca -eq $m) })
                $esc = Invoke-Assistido -Fila $f -Cands $cands5
                if ($esc) {
                    $n5 = Select-NomeValido -Dev $esc
                    if ($n5) { $nome = $n5; $dev = $esc; $origemNome = 'escolha assistida' }
                    else { $motivoFalha = "o aparelho escolhido ($($esc.IP)) não tem nome que resolva para o IP dele" }
                }
            }

            # Sem nome: explica por quê
            if (-not $nome) {
                if ($bloqueioMotivo) { $status = "Ignorada: $bloqueioMotivo" }
                elseif ($motivoFalha) { $status = "Ignorada: $motivoFalha" }
                elseif ($m -and -not $reservaLiberada) {
                    $status = "Ignorada: há $([int]$qtdDisp[$m]) aparelhos $m instalados e não consegui identificar este com segurança (rode antes da troca de rede para gravar o inventário, use o modo assistido ou `$Manual)"
                }
                else { $status = 'Ignorada: não consegui identificar o aparelho desta fila (sem inventário e sem resposta no IP antigo)' }
            }
            else {
                $serieDev = if ($dev -and $dev.Serie) { $dev.Serie } elseif ($inv) { $inv.Serie } else { $null }
                $macDev   = if ($dev -and $dev.Mac) { $dev.Mac } elseif ($inv) { $inv.Mac } else { $null }
                $ipAtual  = if ($dev) { $dev.IP } else { Get-IPdoNome -Nome $nome }

                # Confere o aparelho contra o inventário e contra outras filas
                if ($origemNome -ne 'manual' -and $dev -and $dev.Serie -and $inv -and $inv.Serie -and $dev.Serie -ne $inv.Serie) {
                    $status = "Ignorada: o aparelho encontrado (S/N $($dev.Serie)) não é o registrado para esta fila (S/N $($inv.Serie))"
                }
                $outraFila = $null
                if (-not $status -and $origemNome -ne 'manual') {
                    $outraFila = Test-ConflitoNo -Nome $nome -IdIP $ipAtual -IdMac $macDev -IdSerie $serieDev
                }
                if (-not $status) {
                    if (-not (Get-IPdoNome -Nome $nome) -and -not $Forcar) { $status = "Ignorada: '$nome' não resolve neste PC (use `$Forcar = `$true para ignorar)" }
                    elseif ($outraFila) { $status = "Ignorada: o nó '$nome' já é usado pela fila '$outraFila' com outro aparelho; use `$Manual se forem o mesmo" }
                }

                if (-not $status) {
                    $novaPorta = $nome
                    $extra = if ($origemNome) { " [$origemNome]" } else { '' }

                    if ($Simular) {
                        $status = "SIMULAÇÃO: trocaria '$($imp.PortName)' por '$novaPorta'$extra"
                        Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $m -Nome $nome -IP $ipAtual -Mac $macDev -Serie $serieDev -Modelo $dev.Modelo
                        $nomesAtribuidos[$nome.ToUpper()] = @{ IP = $ipAtual; Mac = $macDev; Serie = $serieDev; Fila = $imp.Name }
                    }
                    else {
                        try {
                            # Protocolo da porta antiga (Registro): 1 = RAW, 2 = LPR
                            $prot = if ($f.Reg -and $f.Reg.Protocol) { [int]$f.Reg.Protocol } else { 1 }
                            $filaLpr = if ($f.Reg -and $f.Reg.Queue) { [string]$f.Reg.Queue } else { [string]$f.Porta.LprQueueName }
                            $usaLpr = ($prot -eq 2)
                            $tcp = if ($usaLpr) { 515 } elseif ($f.Reg -and $f.Reg.PortNumber) { [int]$f.Reg.PortNumber } elseif ($f.Porta.PortNumber) { [int]$f.Porta.PortNumber } else { 9100 }
                            if ($usaLpr -and -not $filaLpr) { throw "a porta antiga é LPR mas não achei o nome da fila LPR" }
                            $criou = $false

                            if (-not (Get-PrinterPort -Name $novaPorta -ErrorAction SilentlyContinue)) {
                                $p = @{ Name = $novaPorta }
                                if ($usaLpr) {
                                    $p.LprHostAddress = $nome
                                    $p.LprQueueName = $filaLpr
                                    if (($f.Reg -and $f.Reg.'Byte Count' -eq 1) -or $f.Porta.LprByteCounting) { $p.LprByteCounting = $true }
                                } else {
                                    $p.PrinterHostAddress = $nome
                                    $p.PortNumber = $tcp
                                }
                                if ($f.Porta.SNMPEnabled) {
                                    $p.SNMP = $f.Porta.SNMPIndex
                                    if ($f.Porta.SNMPCommunity) { $p.SNMPCommunity = $f.Porta.SNMPCommunity }
                                }
                                Add-PrinterPort @p
                                $criou = $true
                            }
                            Set-Printer -Name $imp.Name -PortName $novaPorta

                            # Verifica: responde E é o mesmo aparelho (série, ou MAC se não houver série)
                            $motivoRev = $null
                            if ($VerificarDepois -and -not $Forcar) {
                                if (-not (Test-ImpressoraResponde -Nome $novaPorta -PortaTcp $tcp)) { $motivoRev = "'$novaPorta' não respondeu na porta $tcp" }
                                else {
                                    $ipN = Get-IPdoNome -Nome $novaPorta
                                    if ($ipN) {
                                        $x = Get-Identidade -IP $ipN -Coms $comunidadesTodas
                                        if ($x.Serie -and $serieDev -and $x.Serie -ne $serieDev) { $motivoRev = "'$novaPorta' respondeu, mas é OUTRO aparelho (S/N $($x.Serie); esperado $serieDev)" }
                                        elseif (-not ($x.Serie -and $serieDev) -and $x.Mac -and $macDev -and $x.Mac -ne $macDev) { $motivoRev = "'$novaPorta' respondeu, mas é OUTRO aparelho (MAC $($x.Mac); esperado $macDev)" }
                                    }
                                }
                            }
                            if ($motivoRev) {
                                Set-Printer -Name $imp.Name -PortName $imp.PortName
                                if ($criou -and -not (Get-Printer | Where-Object PortName -eq $novaPorta)) {
                                    Remove-PrinterPort -Name $novaPorta -ErrorAction SilentlyContinue
                                }
                                $status = "REVERTIDA: $motivoRev; voltou para '$($imp.PortName)'$extra"
                                $novaPorta = $null
                            }
                            else {
                                [void]$portasAntigas.Add($imp.PortName)
                                $status = "OK: '$($imp.PortName)' -> '$novaPorta'$extra"
                                Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $m -Nome $nome -IP $ipAtual -Mac $macDev -Serie $serieDev -Modelo $dev.Modelo
                                $nomesAtribuidos[$nome.ToUpper()] = @{ IP = $ipAtual; Mac = $macDev; Serie = $serieDev; Fila = $imp.Name }
                            }
                        }
                        catch { $status = "ERRO: $($_.Exception.Message)"; $novaPorta = $null }
                    }
                }
            }
        }
    }

    $cor = switch -Wildcard ($status) { 'OK*' {'Green'} 'SIMULA*' {'Cyan'} 'ERRO*' {'Red'} 'REVERT*' {'Yellow'} default {'DarkGray'} }
    Write-Host ("{0,-40} {1}" -f $imp.Name, $status) -ForegroundColor $cor

    $resultado.Add([PSCustomObject]@{
        Impressora = $imp.Name; TipoPorta = $tipo; IPAntigo = $ip; IPAtual = $ipAtual
        NomeDoNo = $nome; Serie = $(if ($dev) { $dev.Serie } else { $null }); PortaAntiga = $imp.PortName; PortaNova = $novaPorta
    })
}

if (-not $Simular -and $RemoverPortaAntiga) {
    foreach ($pa in $portasAntigas) {
        if (-not (Get-Printer | Where-Object PortName -eq $pa)) {
            try { Remove-PrinterPort -Name $pa; Write-Host "Porta antiga removida: $pa" -ForegroundColor Green }
            catch { Write-Host "Não foi possível remover '$pa': $($_.Exception.Message)" -ForegroundColor Yellow }
        } else { Write-Host "Porta '$pa' mantida (ainda em uso por outra impressora)." -ForegroundColor Yellow }
    }
}

Write-Host ""
$resultado | Format-Table -AutoSize
Write-Inventario
