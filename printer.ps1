# =====================================================================
#  printer.ps1 - restaura portas de impressoras para o NOME DO NÓ após
#  mudança de IP/rede, identificando cada aparelho por série, UUID, MAC
#  e hostname (SNMP, IPP, PJL, HTTP, WSD, mDNS, UPnP e inventário local).
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

# --- Identificação multi-fonte (todas as sondas são somente leitura) ---
$UsarIpp            = $true    # IPP (631) só quando o SNMP falhar ou não trouxer série
$UsarPjl            = $true    # PJL over 9100 (INFO ID/PRODINFO/CONFIG)
$UsarHttpIdentidade = $true    # HP LEDM / páginas de estado, último recurso
$UsarWsd            = $true    # WS-Discovery (Resolve por UUID e Probe)
$UsarMdns           = $true    # mDNS/DNS-SD (_ipp._tcp etc.)
$UsarUpnp           = $true    # SSDP/UPnP Printer:1
$UsarSlp            = $false   # SLP 427 (HP Jetdirect), opcional
$PermitirPortaPorIP = $true    # sem hostname: usa o IP atual e avisa
$CriarAliasHosts    = $false   # reservado: criar alias no ficheiro hosts (não implementado)
$MigrarHPStandardTCPIP = $true # migra filas do monitor HP Standard TCP/IP
$TimeoutSondaMs     = 1500     # timeout padrão de cada sonda ativa

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
    $nomes = @($Nome)
    if ($Nome -notmatch '\.' -and -not (Test-EhIPv4 $Nome)) { $nomes += "$Nome.local" }
    foreach ($n in $nomes) {
        try {
            $a = [System.Net.Dns]::GetHostAddresses($n) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1
            if ($a) { return $a.IPAddressToString }
        } catch { }
        try {
            $r = Resolve-DnsName -Name $n -NetbiosFallback -ErrorAction Stop |
                Where-Object { $_.IPAddress } | Select-Object -First 1
            if ($r) { return $r.IPAddress }
        } catch { }
        if ($n -notmatch '\.') {
            try {
                $r = Resolve-DnsName -Name $n -LlmnrNetbiosOnly -ErrorAction Stop |
                    Where-Object { $_.IPAddress } | Select-Object -First 1
                if ($r) { return $r.IPAddress }
            } catch { }
        }
    }
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

# ------------------------------------------------------------------ IPP ----

function Add-IppTexto {
    param([System.Collections.Generic.List[byte]]$Lista, [byte]$Tag, [string]$Nome, [string[]]$Valores)
    $Lista.Add($Tag)
    $n = [System.Text.Encoding]::UTF8.GetBytes($Nome)
    $Lista.Add([byte](($n.Length -shr 8) -band 0xFF))
    $Lista.Add([byte]($n.Length -band 0xFF))
    $Lista.AddRange($n)
    $vals = @($Valores)
    for ($i = 0; $i -lt $vals.Count; $i++) {
        $v = [System.Text.Encoding]::UTF8.GetBytes([string]$vals[$i])
        $Lista.Add([byte](($v.Length -shr 8) -band 0xFF))
        $Lista.Add([byte]($v.Length -band 0xFF))
        $Lista.AddRange($v)
        if ($i -lt $vals.Count - 1) { $Lista.Add($Tag) }
    }
}

function Invoke-IppPedido {
    param([string]$IP, [string]$Caminho, [string]$Versao = '2.0', [int]$TimeoutMs = 1500)
    $major = 2; $minor = 0
    if ($Versao -eq '1.1') { $major = 1; $minor = 1 }

    $b = New-Object System.Collections.Generic.List[byte]
    $b.AddRange([byte[]]@($major, $minor, 0x00, 0x0B, 0, 0, 0, 1))
    $b.Add(0x01)
    Add-IppTexto -Lista $b -Tag 0x47 -Nome 'attributes-charset' -Valores @('utf-8')
    Add-IppTexto -Lista $b -Tag 0x48 -Nome 'attributes-natural-language' -Valores @('en')
    Add-IppTexto -Lista $b -Tag 0x45 -Nome 'printer-uri' -Valores @("ipp://$IP$Caminho")
    Add-IppTexto -Lista $b -Tag 0x44 -Nome 'requested-attributes' -Valores @(
        'printer-uuid', 'printer-name', 'printer-info', 'printer-make-and-model',
        'printer-device-id', 'printer-dns-sd-name', 'printer-location'
    )
    $b.Add(0x03)
    $corpo = $b.ToArray()

    $cab = "POST $Caminho HTTP/1.1`r`nHost: $IP`r`nContent-Type: application/ipp`r`nAccept: application/ipp`r`nContent-Length: $($corpo.Length)`r`nConnection: close`r`n`r`n"
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        if (-not $c.ConnectAsync($IP, 631).Wait($TimeoutMs)) { return $null }
        $s = $c.GetStream()
        $s.ReadTimeout = $TimeoutMs
        $s.WriteTimeout = $TimeoutMs
        $hb = [System.Text.Encoding]::ASCII.GetBytes($cab)
        $s.Write($hb, 0, $hb.Length)
        $s.Write($corpo, 0, $corpo.Length)
        $s.Flush()

        $ms = New-Object System.IO.MemoryStream
        $buf = New-Object byte[] 8192
        $fim = (Get-Date).AddMilliseconds($TimeoutMs * 2)
        while ((Get-Date) -lt $fim) {
            try {
                $l = $s.Read($buf, 0, $buf.Length)
                if ($l -le 0) { break }
                $ms.Write($buf, 0, $l)
            } catch { break }
        }
        $resp = $ms.ToArray()
        if ($resp.Length -lt 16) { return $null }

        $txt = [System.Text.Encoding]::ASCII.GetString($resp)
        $idx = $txt.IndexOf("`r`n`r`n")
        if ($idx -lt 0) { return $null }
        $header = $txt.Substring(0, $idx)
        if ($header -notmatch '^HTTP/\d\.\d\s+200') { return $null }
        $ini = $idx + 4
        if ($ini -ge $resp.Length) { return $null }

        $body = New-Object byte[] ($resp.Length - $ini)
        [Array]::Copy($resp, $ini, $body, 0, $body.Length)
        if ($body.Length -lt 8) { return $null }
        $status = ([int]$body[2] -shl 8) + [int]$body[3]
        if ($status -ne 0) { return $null }

        $attrs = @{}
        $pos = 8
        $ultimo = $null
        while ($pos -lt $body.Length) {
            $tag = [int]$body[$pos]; $pos++
            if ($tag -eq 0x03) { break }
            if ($tag -lt 0x10) { continue }
            if ($pos + 2 -gt $body.Length) { break }
            $nlen = ([int]$body[$pos] -shl 8) + [int]$body[$pos + 1]; $pos += 2
            if ($nlen -gt 0) {
                if ($pos + $nlen -gt $body.Length) { break }
                $nome = [System.Text.Encoding]::UTF8.GetString($body, $pos, $nlen)
                $pos += $nlen
                $ultimo = $nome
            } else { $nome = $ultimo }
            if ($pos + 2 -gt $body.Length) { break }
            $vlen = ([int]$body[$pos] -shl 8) + [int]$body[$pos + 1]; $pos += 2
            if ($vlen -gt 0 -and $pos + $vlen -le $body.Length -and $nome) {
                $attrs[$nome] = [System.Text.Encoding]::UTF8.GetString($body, $pos, $vlen)
            }
            $pos += $vlen
        }
        if ($attrs.Count -gt 0) { return $attrs }
        return $null
    } catch { return $null }
    finally { $c.Dispose() }
}

function Get-IppAtributos {
    param([string]$IP, [int]$TimeoutMs = 1500)
    if (-not $UsarIpp -or -not $IP) { return $null }
    if (-not (Test-Porta -Alvo $IP -Porta 631 -TimeoutMs 800)) { return $null }
    $tentativas = @(
        @{ Caminho = '/ipp/print'; Versao = '2.0' },
        @{ Caminho = '/ipp';       Versao = '2.0' },
        @{ Caminho = '/ipp/port1'; Versao = '2.0' },
        @{ Caminho = '/ipp/print'; Versao = '1.1' },
        @{ Caminho = '/';          Versao = '1.1' }
    )
    foreach ($t in $tentativas) {
        $r = Invoke-IppPedido -IP $IP -Caminho $t.Caminho -Versao $t.Versao -TimeoutMs $TimeoutMs
        if ($r) { return $r }
    }
    return $null
}

# ------------------------------------------------------------------ PJL ----

function Get-PjlInfo {
    param([string]$IP, [int]$TimeoutMs = 1500)
    if (-not $UsarPjl -or -not $IP) { return $null }
    if (-not (Test-Porta -Alvo $IP -Porta 9100 -TimeoutMs 800)) { return $null }

    $esc = [string][char]27
    $comandos = "$esc%-12345X@PJL INFO ID`r`n@PJL INFO PRODINFO`r`n@PJL INFO CONFIG`r`n$esc%-12345X"
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        if (-not $c.ConnectAsync($IP, 9100).Wait($TimeoutMs)) { return $null }
        $s = $c.GetStream()
        $s.ReadTimeout = 700
        $s.WriteTimeout = $TimeoutMs
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($comandos)
        $s.Write($bytes, 0, $bytes.Length)
        $s.Flush()

        $ms = New-Object System.IO.MemoryStream
        $buf = New-Object byte[] 8192
        $fim = (Get-Date).AddMilliseconds($TimeoutMs)
        while ((Get-Date) -lt $fim -and $ms.Length -lt 262144) {
            try {
                $l = $s.Read($buf, 0, $buf.Length)
                if ($l -le 0) { break }
                $ms.Write($buf, 0, $l)
            } catch { break }
        }
        $texto = [System.Text.Encoding]::ASCII.GetString($ms.ToArray())
        if (-not $texto) { return $null }
        $texto = $texto -replace "`0", ''

        # PRODINFO costuma vir como "CHAVE"="VALOR" em cada linha
        $campos = @{}
        foreach ($linha in ($texto -split '\r?\n')) {
            $mm = [regex]::Match($linha, '^\s*"([^"]+)"\s*=\s*"([^"]*)"')
            if ($mm.Success) { $campos[$mm.Groups[1].Value.Trim().ToUpperInvariant()] = $mm.Groups[2].Value.Trim() }
        }

        $serie = $null
        foreach ($k in @($campos.Keys | Where-Object { $_ -match 'SERIAL|S/?N' })) {
            $serie = Normalize-Serie $campos[$k]; if ($serie) { break }
        }
        if (-not $serie) {
            $m = [regex]::Match($texto, '(?im)^\s*"?\s*(?:ProductSerialNumber|SerialNumber|Serial\s+Number|SERIAL\s*NUMBER|S/N)\s*"?\s*[=:]\s*"?([A-Z0-9][A-Z0-9\-]{3,})')
            if ($m.Success) { $serie = Normalize-Serie $m.Groups[1].Value }
        }

        $mac = $null
        foreach ($k in @($campos.Keys | Where-Object { $_ -match 'HW\s*ADDRESS|MAC' })) {
            $mac = Normalize-Mac $campos[$k]; if ($mac) { break }
        }
        if (-not $mac) {
            $m = [regex]::Match($texto, '(?im)^\s*"?\s*(?:HWAddress|MAC(?:\s*Address)?)\s*"?\s*[=:]\s*"?([0-9A-Fa-f:\-]{12,17})')
            if ($m.Success) { $mac = Normalize-Mac $m.Groups[1].Value }
        }

        $modelo = $null
        foreach ($k in @($campos.Keys | Where-Object { $_ -match 'MODEL|PRODUCT' })) {
            if ($campos[$k]) { $modelo = $campos[$k]; break }
        }
        if (-not $modelo) {
            $m = [regex]::Match($texto, '(?im)^\s*@PJL\s+INFO\s+ID\s*$[\r\n]+\s*(.{3,80})$')
            if ($m.Success) { $modelo = $m.Groups[1].Value.Trim() }
        }
        if (-not $modelo) {
            $m = [regex]::Match($texto, '(?im)^\s*(?:MODEL|PRODUCT|PrinterModel)\s*[=:]\s*(.{3,80})$')
            if ($m.Success) { $modelo = $m.Groups[1].Value.Trim() }
        }

        return [PSCustomObject]@{ Texto = $texto; Serie = $serie; Mac = $mac; Modelo = $modelo }
    } catch { return $null }
    finally { $c.Dispose() }
}

# ----------------------------------------------------------------- HTTP ----

function Get-HttpIdentidade {
    param([string]$IP, [string]$Marca = $null, [int]$TimeoutMs = 2000)
    if (-not $UsarHttpIdentidade -or -not $IP) { return $null }
    $urls = @("http://$IP/DevMgmt/ProductConfigDyn.xml")
    if ($Marca -eq 'Brother') {
        $urls += "http://$IP/general/status.html", "http://$IP/net/status.html", "http://$IP/home/status.html"
    } elseif ($Marca -eq 'Epson') {
        $urls += "http://$IP/PRESENTATION/HTML/TOP/PRTINFO.html"
    }
    $urls += "http://$IP/"

    $r = [PSCustomObject]@{ Serie = $null; Modelo = $null; Uuid = $null; Mac = $null; Texto = $null }
    foreach ($u in $urls) {
        $html = $null
        try { $html = [string](Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec ([math]::Max(1, [int]($TimeoutMs / 1000))) -ErrorAction Stop).Content }
        catch { continue }
        if (-not $html) { continue }

        if (-not $r.Serie) {
            $m = [regex]::Match($html, '(?is)<serialNumber>\s*([^<]{4,40})\s*</serialNumber>')
            if (-not $m.Success) { $m = [regex]::Match($html, '(?i)(?:serial(?:\s*number)?|s/?n|n[úu]mero\s+de\s+s[ée]rie)\s*[:=#]?\s*["'']?([A-Z0-9][A-Z0-9\-]{3,})') }
            if ($m.Success) { $r.Serie = Normalize-Serie $m.Groups[1].Value }
        }
        if (-not $r.Modelo) {
            $m = [regex]::Match($html, '(?is)<modelName>\s*([^<]{3,80})\s*</modelName>')
            if (-not $m.Success) { $m = [regex]::Match($html, '(?is)<title>\s*([^<]{3,120})\s*</title>') }
            if (-not $m.Success) { $m = [regex]::Match($html, '(?im)^\s*(?:MODEL|MODELO|PRODUCT)\s*[:=]\s*(.{3,80})$') }
            if ($m.Success) { $r.Modelo = ($m.Groups[1].Value.Trim() -replace '\s+', ' ') }
        }
        if (-not $r.Uuid) {
            $m = [regex]::Match($html, '(?is)<deviceUUID>\s*([^<]{6,60})\s*</deviceUUID>')
            if (-not $m.Success) { $m = [regex]::Match($html, '(?i)(?:uuid|deviceUuid)\s*[:=]\s*["'']?([0-9a-f\-]{20,40})') }
            if ($m.Success) { $r.Uuid = Normalize-Uuid $m.Groups[1].Value }
        }
        if (-not $r.Mac) {
            $m = [regex]::Match($html, '(?is)<macAddress>\s*([^<]{12,20})\s*</macAddress>')
            if ($m.Success) { $r.Mac = Normalize-Mac $m.Groups[1].Value }
        }

        $txt = ($html -replace '(?is)<script.*?</script>', ' ' -replace '(?is)<style.*?</style>', ' ' -replace '<[^>]+>', ' ')
        if ($txt.Length -gt 20000) { $txt = $txt.Substring(0, 20000) }
        $r.Texto = "$($r.Texto) $txt"
        if ($r.Serie -and $r.Modelo) { break }
    }
    if ($r.Texto) { $r.Texto = ($r.Texto -replace '\s+', ' ').Trim() }
    if ($r.Serie -or $r.Modelo -or $r.Uuid) { return $r }
    return $null
}

# --------------------------------------------------------- WS-Discovery ----

function Get-WsdIpDoXAddrs {
    param([string]$XAddrs)
    if (-not $XAddrs) { return $null }
    foreach ($x in ($XAddrs -split '\s+')) {
        if (-not $x) { continue }
        $hostX = $null
        if ($x -match '^https?://') {
            try { $hostX = ([uri]$x).Host } catch { }
        } elseif ($x -match '^([^\s/:]+)(?::\d+)?(?:/|$)') {
            $hostX = $Matches[1]
        }
        if ($hostX) {
            if (Test-EhIPv4 $hostX) { return $hostX }
            $ipX = Get-IPdoNome -Nome $hostX
            if ($ipX) { return $ipX }
            return $hostX
        }
    }
    return $null
}

function Invoke-WsdMensagem {
    param([string]$Mensagem, [int]$TimeoutMs = 2500)
    $u = New-Object System.Net.Sockets.UdpClient
    try {
        $u.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]::Socket, [System.Net.Sockets.SocketOptionName]::ReuseAddress, $true)
        $u.Client.Bind((New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)))
        $u.Client.ReceiveTimeout = 500
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Mensagem)
        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse('239.255.255.250'), 3702)
        [void]$u.Send($bytes, $bytes.Length, $ep)

        $pacotes = @{}
        $fim = (Get-Date).AddMilliseconds($TimeoutMs)
        while ((Get-Date) -lt $fim) {
            try {
                $rem = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
                $r = $u.Receive([ref]$rem)
                $pacotes[[Convert]::ToBase64String($r)] = [System.Text.Encoding]::UTF8.GetString($r)
            } catch { }
        }
        return @($pacotes.Values)
    } catch { return @() }
    finally { $u.Close() }
}

function Resolve-WsdDevice {
    param([string]$Uuid, [int]$TimeoutMs = 2500)
    if (-not $UsarWsd -or -not $Uuid) { return $null }
    $u = Normalize-Uuid $Uuid
    if (-not $u) { return $null }
    $msgId = [guid]::NewGuid().ToString()
    $msg = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" xmlns:wsa="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:wsd="http://schemas.xmlsoap.org/ws/2005/04/discovery">
  <soap:Header>
    <wsa:To>urn:schemas-xmlsoap-org:ws:2005:04:discovery</wsa:To>
    <wsa:Action>http://schemas.xmlsoap.org/ws/2005/04/discovery/Resolve</wsa:Action>
    <wsa:MessageID>urn:uuid:$msgId</wsa:MessageID>
  </soap:Header>
  <soap:Body>
    <wsd:Resolve>
      <wsa:EndpointReference><wsa:Address>urn:uuid:$u</wsa:Address></wsa:EndpointReference>
    </wsd:Resolve>
  </soap:Body>
</soap:Envelope>
"@
    $pacotes = Invoke-WsdMensagem -Mensagem $msg -TimeoutMs $TimeoutMs
    foreach ($p in $pacotes) {
        foreach ($m in [regex]::Matches($p, '(?is)<wsd:ResolveMatch>(.*?)</wsd:ResolveMatch>')) {
            $bloco = $m.Groups[1].Value
            $addr = [regex]::Match($bloco, '(?is)<wsa:Address>\s*urn:uuid:([0-9a-fA-F\-]{20,40})\s*</wsa:Address>')
            if ($addr.Success -and (Normalize-Uuid $addr.Groups[1].Value) -ne $u) { continue }
            $xa = [regex]::Match($bloco, '(?is)<wsd:XAddrs>(.*?)</wsd:XAddrs>')
            if ($xa.Success) {
                $ip = Get-WsdIpDoXAddrs $xa.Groups[1].Value
                if ($ip) { return [PSCustomObject]@{ Uuid = $u; Ip = $ip; XAddrs = $xa.Groups[1].Value.Trim() } }
            }
        }
    }
    return $null
}

function Get-WsdDevices {
    param([int]$TimeoutMs = 2500)
    if (-not $UsarWsd) { return @() }
    $msgId = [guid]::NewGuid().ToString()
    $msg = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" xmlns:wsa="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:wsd="http://schemas.xmlsoap.org/ws/2005/04/discovery" xmlns:wsdp="http://schemas.xmlsoap.org/ws/2006/02/devprof" xmlns:nprt="http://schemas.microsoft.com/windows/2006/08/wdp/print">
  <soap:Header>
    <wsa:To>urn:schemas-xmlsoap-org:ws:2005:04:discovery</wsa:To>
    <wsa:Action>http://schemas.xmlsoap.org/ws/2005/04/discovery/Probe</wsa:Action>
    <wsa:MessageID>urn:uuid:$msgId</wsa:MessageID>
  </soap:Header>
  <soap:Body>
    <wsd:Probe><wsd:Types>wsdp:Device nprt:PrintDeviceType</wsd:Types></wsd:Probe>
  </soap:Body>
</soap:Envelope>
"@
    $pacotes = Invoke-WsdMensagem -Mensagem $msg -TimeoutMs $TimeoutMs
    $lista = New-Object System.Collections.Generic.List[object]
    $vistos = @{}
    foreach ($p in $pacotes) {
        foreach ($m in [regex]::Matches($p, '(?is)<wsd:ProbeMatch>(.*?)</wsd:ProbeMatch>')) {
            $bloco = $m.Groups[1].Value
            $addr = [regex]::Match($bloco, '(?is)<wsa:Address>\s*urn:uuid:([0-9a-fA-F\-]{20,40})\s*</wsa:Address>')
            $xa = [regex]::Match($bloco, '(?is)<wsd:XAddrs>(.*?)</wsd:XAddrs>')
            if (-not $xa.Success) { continue }
            $ip = Get-WsdIpDoXAddrs $xa.Groups[1].Value
            if (-not $ip) { continue }
            $uuid = if ($addr.Success) { Normalize-Uuid $addr.Groups[1].Value } else { $null }
            $chave = "$ip|$uuid"
            if ($vistos.ContainsKey($chave)) { continue }
            $vistos[$chave] = $true
            $lista.Add([PSCustomObject]@{ Uuid = $uuid; Ip = $ip; XAddrs = $xa.Groups[1].Value.Trim() })
        }
    }
    return $lista.ToArray()
}

# ----------------------------------------------------------------- mDNS ----

function Get-DnsNome {
    param([byte[]]$P, [ref]$Pos)
    $partes = New-Object System.Collections.Generic.List[string]
    $idx = [int]$Pos.Value
    $fim = $idx
    $seguindo = $true
    $saltos = 0
    while ($idx -lt $P.Length -and $saltos -lt 20) {
        $l = [int]$P[$idx]
        if ($l -eq 0) {
            $idx = $idx + 1
            if ($seguindo) { $fim = $idx }
            break
        }
        if (($l -band 0xC0) -eq 0xC0) {
            if ($idx + 1 -ge $P.Length) { break }
            $ptr = (($l -band 0x3F) -shl 8) -bor [int]$P[$idx + 1]
            if ($seguindo) { $fim = $idx + 2; $seguindo = $false }
            $idx = $ptr
            $saltos = $saltos + 1
        } else {
            if ($idx + 1 + $l -gt $P.Length) { break }
            $partes.Add([System.Text.Encoding]::ASCII.GetString($P, $idx + 1, $l))
            $idx = $idx + $l + 1
            if ($seguindo) { $fim = $idx }
        }
    }
    $Pos.Value = $fim
    return ($partes -join '.')
}

function Read-MdnsPacote {
    param([byte[]]$P)
    if (-not $P -or $P.Length -lt 12) { return @() }
    $qd = ([int]$P[4] -shl 8) + [int]$P[5]
    $an = ([int]$P[6] -shl 8) + [int]$P[7]
    $ns = ([int]$P[8] -shl 8) + [int]$P[9]
    $ar = ([int]$P[10] -shl 8) + [int]$P[11]
    $pos = 12
    $i = 0
    while ($i -lt $qd) {
        [void](Get-DnsNome -P $P -Pos ([ref]$pos))
        $pos = $pos + 4
        if ($pos -gt $P.Length) { break }
        $i = $i + 1
    }
    $recs = New-Object System.Collections.Generic.List[object]
    $total = $an + $ns + $ar
    $i = 0
    while ($i -lt $total) {
        if ($pos -ge $P.Length) { break }
        $nome = Get-DnsNome -P $P -Pos ([ref]$pos)
        if ($pos + 10 -gt $P.Length) { break }
        $tipo = ([int]$P[$pos] -shl 8) + [int]$P[$pos + 1]
        $pos = $pos + 8
        $rdlen = ([int]$P[$pos] -shl 8) + [int]$P[$pos + 1]
        $pos = $pos + 2
        $rstart = $pos
        $dados = $null
        if ($pos + $rdlen -le $P.Length) {
            if ($tipo -eq 1 -and $rdlen -ge 4) {
                $dados = "$([int]$P[$rstart]).$([int]$P[$rstart + 1]).$([int]$P[$rstart + 2]).$([int]$P[$rstart + 3])"
            } elseif ($tipo -eq 28 -and $rdlen -ge 16) {
                $b16 = New-Object byte[] 16
                [Array]::Copy($P, $rstart, $b16, 0, 16)
                $dados = (New-Object System.Net.IPAddress -ArgumentList (, $b16)).IPAddressToString
            } elseif ($tipo -eq 12) {
                $rp = $rstart
                $dados = Get-DnsNome -P $P -Pos ([ref]$rp)
            } elseif ($tipo -eq 33 -and $rdlen -ge 7) {
                $porta = ([int]$P[$rstart + 4] -shl 8) + [int]$P[$rstart + 5]
                $rp = $rstart + 6
                $alvo = Get-DnsNome -P $P -Pos ([ref]$rp)
                $dados = [PSCustomObject]@{ Porta = $porta; Target = $alvo }
            } elseif ($tipo -eq 16) {
                $txt = @{}
                $tp = $rstart
                while ($tp -lt $rstart + $rdlen) {
                    $l = [int]$P[$tp]; $tp = $tp + 1
                    if ($l -le 0 -or $tp + $l -gt $rstart + $rdlen) { break }
                    $s = [System.Text.Encoding]::UTF8.GetString($P, $tp, $l)
                    $tp = $tp + $l
                    if ($s -match '^([^=]+)=(.*)$') { $txt[$Matches[1]] = $Matches[2] }
                    elseif ($s) { $txt[$s] = '' }
                }
                $dados = $txt
            }
        }
        $recs.Add([PSCustomObject]@{ Nome = $nome; Tipo = $tipo; Dados = $dados })
        $pos = $rstart + $rdlen
        $i = $i + 1
    }
    return $recs.ToArray()
}

function New-MdnsPergunta {
    param([string[]]$Servicos, [bool]$Unicast = $false)
    $b = New-Object System.Collections.Generic.List[byte]
    $b.AddRange([byte[]]@(0, 0, 0, 0, 0, [byte](@($Servicos).Count), 0, 0, 0, 0, 0, 0))
    foreach ($s in $Servicos) {
        foreach ($p in $s.Split('.')) {
            $x = [System.Text.Encoding]::ASCII.GetBytes($p)
            $b.Add([byte]$x.Length)
            $b.AddRange($x)
        }
        $b.Add(0)
        $b.AddRange([byte[]]@(0, 12))
        $cls = if ($Unicast) { 0x8001 } else { 0x0001 }
        $b.Add([byte](($cls -shr 8) -band 0xFF))
        $b.Add([byte]($cls -band 0xFF))
    }
    return , $b.ToArray()
}

function Get-MdnsPrinters {
    param([int]$TimeoutMs = 2500)
    if (-not $UsarMdns) { return @() }
    $servicos = @('_ipp._tcp.local', '_ipps._tcp.local', '_printer._tcp.local', '_pdl-datastream._tcp.local')
    $u = New-Object System.Net.Sockets.UdpClient
    try {
        $u.Client.ExclusiveAddressUse = $false
        $u.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]::Socket, [System.Net.Sockets.SocketOptionName]::ReuseAddress, $true)
        $portaLocal = 5353
        try { $u.Client.Bind((New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, $portaLocal))) }
        catch {
            $portaLocal = 0
            $u.Client.Bind((New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, $portaLocal)))
        }
        try { $u.JoinMulticastGroup([System.Net.IPAddress]::Parse('224.0.0.251')) } catch { }
        $u.Client.ReceiveTimeout = 500
        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse('224.0.0.251'), 5353)
        $q1 = New-MdnsPergunta -Servicos $servicos -Unicast $true
        [void]$u.Send($q1, $q1.Length, $ep)
        if ($portaLocal -eq 5353) {
            $q2 = New-MdnsPergunta -Servicos $servicos -Unicast $false
            [void]$u.Send($q2, $q2.Length, $ep)
        }

        $pacotes = @{}
        $fim = (Get-Date).AddMilliseconds($TimeoutMs)
        while ((Get-Date) -lt $fim) {
            try {
                $rem = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
                $r = $u.Receive([ref]$rem)
                $pacotes[[Convert]::ToBase64String($r)] = $r
            } catch { }
        }
    } catch { return @() }
    finally { $u.Close() }

    $recs = New-Object System.Collections.Generic.List[object]
    foreach ($p in $pacotes.Values) {
        try { foreach ($rec in (Read-MdnsPacote -P $p)) { $recs.Add($rec) } } catch { }
    }

    $resultado = New-Object System.Collections.Generic.List[object]
    $vistos = @{}
    foreach ($ptr in @($recs | Where-Object { $_.Tipo -eq 12 -and $_.Nome -match '^_(ipp|ipps|printer|pdl-datastream)\._tcp\.local$' })) {
        $inst = $ptr.Dados
        if (-not $inst -or $vistos.ContainsKey($inst)) { continue }
        $vistos[$inst] = $true
        $srv = @($recs | Where-Object { $_.Tipo -eq 33 -and $_.Nome -eq $inst }) | Select-Object -First 1
        $txt = @($recs | Where-Object { $_.Tipo -eq 16 -and $_.Nome -eq $inst }) | Select-Object -First 1
        $hostNome = $null; $porta = $null
        if ($srv -and $srv.Dados) { $hostNome = $srv.Dados.Target; $porta = $srv.Dados.Porta }
        $campos = @{}
        if ($txt -and $txt.Dados) { foreach ($k in $txt.Dados.Keys) { $campos[$k] = $txt.Dados[$k] } }

        $ip = $null
        if ($hostNome) {
            $a = @($recs | Where-Object { $_.Tipo -eq 1 -and $_.Nome -eq $hostNome }) | Select-Object -First 1
            if ($a) { $ip = $a.Dados }
            if (-not $ip) { $ip = Get-IPdoNome -Nome $hostNome }
        }
        if (-not $ip) { continue }

        $modelo = $null
        foreach ($k in 'ty', 'product', 'usb_MDL') { if ($campos[$k]) { $modelo = $campos[$k]; break } }
        $marca = Get-MarcaPorTexto -Texto "$($campos['usb_MFG']) $modelo $($campos['ty'])"
        $uuid = Normalize-Uuid ($(if ($campos['UUID']) { $campos['UUID'] } else { $campos['DUUID'] }))

        $resultado.Add([PSCustomObject]@{
            IP = $ip; Nome = $hostNome; Modelo = $modelo; Marca = $marca
            IppUuid = $uuid; HostNames = @($hostNome, $inst)
            Porta = $porta; Servico = $ptr.Nome
        })
    }
    return $resultado.ToArray()
}

# --------------------------------------------------------- SSDP / UPnP ----

function Get-SsdpPrinters {
    param([int]$TimeoutMs = 2000)
    if (-not $UsarUpnp) { return @() }
    $u = New-Object System.Net.Sockets.UdpClient
    try {
        $u.Client.ReceiveTimeout = 500
        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse('239.255.255.250'), 1900)
        $buscas = @(
            "M-SEARCH * HTTP/1.1`r`nHOST: 239.255.255.250:1900`r`nMAN: `"ssdp:discover`"`r`nMX: 2`r`nST: urn:schemas-upnp-org:device:Printer:1`r`n`r`n",
            "M-SEARCH * HTTP/1.1`r`nHOST: 239.255.255.250:1900`r`nMAN: `"ssdp:discover`"`r`nMX: 2`r`nST: ssdp:all`r`n`r`n"
        )
        foreach ($m in $buscas) {
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($m)
            try { [void]$u.Send($bytes, $bytes.Length, $ep) } catch { }
        }
        $locais = @{}
        $fim = (Get-Date).AddMilliseconds($TimeoutMs)
        while ((Get-Date) -lt $fim) {
            try {
                $rem = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
                $r = $u.Receive([ref]$rem)
                $txt = [System.Text.Encoding]::ASCII.GetString($r)
                $loc = [regex]::Match($txt, '(?im)^LOCATION:\s*(.+?)\s*$')
                if ($loc.Success) { $locais[$loc.Groups[1].Value.Trim()] = $true }
            } catch { }
        }
    } catch { return @() }
    finally { $u.Close() }

    $resultado = New-Object System.Collections.Generic.List[object]
    $vistos = @{}
    foreach ($loc in $locais.Keys) {
        if ($vistos.ContainsKey($loc)) { continue }
        $vistos[$loc] = $true
        try {
            $hostIp = ([uri]$loc).Host
            if (-not $hostIp) { continue }
            if (-not (Test-EhIPv4 $hostIp)) { $ipLoc = Get-IPdoNome -Nome $hostIp; if ($ipLoc) { $hostIp = $ipLoc } }
            $xml = [string](Invoke-WebRequest -Uri $loc -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop).Content
            if ($xml -notmatch '(?i)Printer') { continue }
            $nome = [regex]::Match($xml, '(?is)<friendlyName>\s*([^<]+?)\s*</friendlyName>').Groups[1].Value
            $fabricante = [regex]::Match($xml, '(?is)<manufacturer>\s*([^<]+?)\s*</manufacturer>').Groups[1].Value
            $modelo = [regex]::Match($xml, '(?is)<modelName>\s*([^<]+?)\s*</modelName>').Groups[1].Value
            $serie = [regex]::Match($xml, '(?is)<serialNumber>\s*([^<]+?)\s*</serialNumber>').Groups[1].Value
            $udn = [regex]::Match($xml, '(?is)<UDN>\s*([^<]+?)\s*</UDN>').Groups[1].Value
            $resultado.Add([PSCustomObject]@{
                IP = $hostIp; Nome = $nome
                Marca = (Get-MarcaPorTexto -Texto "$fabricante $nome $modelo")
                Modelo = $modelo; Serie = (Normalize-Serie $serie); UpnpUuid = (Normalize-Uuid $udn)
                HostNames = @($nome)
            })
        } catch { }
    }
    return $resultado.ToArray()
}

# ------------------------------------------------ Identificadores comuns ----

function Normalize-Uuid {
    param([string]$Valor)
    if (-not $Valor) { return $null }
    $v = ([string]$Valor).Trim().ToLowerInvariant()
    $v = $v -replace '^\s*(urn:)?uuid:', ''
    $v = $v -replace '[^0-9a-f]', ''
    if ($v.Length -eq 32) {
        return "$($v.Substring(0,8))-$($v.Substring(8,4))-$($v.Substring(12,4))-$($v.Substring(16,4))-$($v.Substring(20,12))"
    }
    if ($v.Length -eq 0) { return $null }
    return $v
}

function Normalize-Mac {
    param([string]$Valor)
    if (-not $Valor) { return $null }
    $v = ($Valor -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    if ($v.Length -eq 12) { return $v }
    return $null
}

function Normalize-Nome {
    param([string]$Valor)
    if (-not $Valor) { return $null }
    return (($Valor.Trim().ToUpperInvariant()) -replace '\.LOCAL\.?$', '')
}

function Normalize-Serie {
    param([string]$Valor)
    if (-not $Valor) { return $null }
    $v = ([string]$Valor).Trim().ToUpperInvariant()
    $v = $v -replace '(?i)^(S\s*/\s*N|SERIAL(?:\s*NUMBER)?|SER(?:IAL)?\.?\s*NO\.?)[\s:.#_-]*', ''
    $v = $v -replace '[^A-Z0-9]', ''
    if ($v.Length -lt 4) { return $null }
    if ($v -match '^(N/?A|UNKNOWN|NONE|NOTAVAILABLE|0+|-+)$') { return $null }
    return $v
}

function Get-UuidDoTexto {
    param([string]$Texto)
    if (-not $Texto) { return $null }
    $m = [regex]::Match($Texto, '(?i)(?:uuid[^0-9a-f]{0,12})?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})')
    if ($m.Success) { return Normalize-Uuid $m.Groups[1].Value }
    return $null
}

function Test-MesmoAparelho {
    param($A, $B)
    if (-not $A -or -not $B) { return $false }

    $sa = Normalize-Serie $A.Serie
    $sb = Normalize-Serie $B.Serie
    if ($sa -and $sb) { return ($sa -eq $sb) }

    $uas = @(); $ubs = @()
    foreach ($campo in 'IppUuid', 'WsdUuid', 'DeviceGuid', 'UpnpUuid') {
        $v = Normalize-Uuid $A.$campo
        if ($v) { $uas += $v }
        $v = Normalize-Uuid $B.$campo
        if ($v) { $ubs += $v }
    }
    # UUIDs podem aparecer em campos diferentes (IPP/WSD/UPnP); interseção é positiva
    if (@($uas | Where-Object { $ubs -contains $_ }).Count -gt 0) { return $true }

    $ma = Normalize-Mac $A.Mac
    $mb = Normalize-Mac $B.Mac
    if ($ma -and $mb) { return ($ma -eq $mb) }
    return $false
}

function Test-IdentidadeConflita {
    param($A, $B)
    if (-not $A -or -not $B) { return $false }
    $sa = Normalize-Serie $A.Serie
    $sb = Normalize-Serie $B.Serie
    if ($sa -and $sb) { return ($sa -ne $sb) }

    $uas = @(); $ubs = @()
    foreach ($campo in 'IppUuid', 'WsdUuid', 'DeviceGuid', 'UpnpUuid') {
        $v = Normalize-Uuid $A.$campo
        if ($v) { $uas += $v }
        $v = Normalize-Uuid $B.$campo
        if ($v) { $ubs += $v }
    }
    $uuidComparado = ($uas.Count -gt 0 -and $ubs.Count -gt 0)
    if ($uuidComparado -and @($uas | Where-Object { $ubs -contains $_ }).Count -gt 0) { return $false }

    $ma = Normalize-Mac $A.Mac
    $mb = Normalize-Mac $B.Mac
    if ($ma -and $mb) { return ($ma -ne $mb) }
    if ($uuidComparado) { return $true }
    return $false
}

function Merge-Descoberta {
    param($Base, $Extra)
    if (-not $Base) { return $Extra }
    if (-not $Extra) { return $Base }
    foreach ($campo in 'Serie', 'Mac', 'Modelo', 'Nome', 'Marca', 'SysName', 'Descr', 'IppUuid', 'WsdUuid', 'UpnpUuid', 'DeviceGuid') {
        if (-not $Base.$campo -and $Extra.$campo) { $Base.$campo = $Extra.$campo }
    }
    if (-not $Base.Vivo -and $Extra.Vivo) { $Base.Vivo = $true }
    if (-not $Base.Snmp -and $Extra.Snmp) { $Base.Snmp = $true }
    $hosts = @()
    foreach ($h in @($Base.HostNames) + @($Extra.HostNames) + @($Extra.Nome)) { if ($h) { $hosts += [string]$h } }
    $Base.HostNames = @($hosts | Select-Object -Unique)
    $extraTxt = "$($Extra.Modelo) $($Extra.Marca) $($Extra.Nome) $($Extra.Texto)"
    $t = ([string]$Base.Texto + ' ' + $extraTxt) -replace '[^A-Za-z0-9]', ''
    $t = $t.ToUpperInvariant()
    if ($t.Length -gt 40000) { $t = $t.Substring(0, 40000) }
    $Base.Texto = $t
    if (-not $Base.Marca) { $Base.Marca = Get-MarcaPorTexto -Texto "$($Base.Modelo) $($Base.Descr) $($Base.Texto) $extraTxt" }
    return $Base
}

function Test-CandidatoAtribuido {
    param($Cand, $Atribuidos, [string]$FilaAtual)
    foreach ($at in @($Atribuidos)) {
        if (-not $at -or $at.Fila -eq $FilaAtual) { continue }
        if (Test-MesmoAparelho -A $Cand -B $at) { continue }
        $nc = Normalize-Nome $Cand.Nome
        $na = Normalize-Nome $at.Nome
        if ($nc -and $na -and $nc -eq $na) { return $at }
        if ($Cand.IP -and $at.IP -and [string]$Cand.IP -eq [string]$at.IP) { return $at }
    }
    return $null
}

# ------------------------------------------------------------ Identidade ----

# Lê tudo o que identifica o aparelho que responde em $IP (SNMP, IPP, PJL, HTTP, ARP, nome, marca)
function Get-Identidade {
    param([string]$IP, [string[]]$Coms)
    $id = [PSCustomObject]@{
        IP = $IP; Vivo = $false; Snmp = $false; Serie = $null; Modelo = $null; SysName = $null
        Descr = $null; Mac = $null; Nome = $null; Marca = $null; Texto = $null
        Ipp = @{}; IppNome = $null; Pjl = $null; Http = $null
        IppUuid = $null; WsdUuid = $null; UpnpUuid = $null; DeviceGuid = $null
        HostNames = @(); SerieFonte = $null; ModeloFonte = $null; UuidFonte = $null
    }
    $com = $null
    foreach ($c in $Coms) {
        $d = Get-Snmp -IP $IP -Oid '1.3.6.1.2.1.1.1.0' -Comunidade $c
        if ($d) { $id.Descr = [string]$d; $id.Snmp = $true; $com = $c; break }
    }
    if ($id.Snmp) {
        $id.Vivo = $true
        $s = Get-Snmp -IP $IP -Oid '1.3.6.1.2.1.43.5.1.1.17.1' -Comunidade $com
        if (-not $s -and $id.Descr -match 'S/N\s+([A-Za-z0-9\-]+)') { $s = $Matches[1] }
        if ($s -and ([string]$s).Trim().Length -ge 4) { $id.Serie = Normalize-Serie ([string]$s); if ($id.Serie) { $id.SerieFonte = 'SNMP' } }
        $id.Modelo  = Get-Snmp -IP $IP -Oid '1.3.6.1.2.1.25.3.2.1.3.1' -Comunidade $com
        if ($id.Modelo) { $id.ModeloFonte = 'SNMP' }
        $id.SysName = Get-Snmp -IP $IP -Oid '1.3.6.1.2.1.1.5.0' -Comunidade $com
    }
    else {
        # Sem SNMP: considera "vivo" se alguma porta de impressão abrir
        foreach ($pt in 9100, 515, 631) { if (Test-Porta -Alvo $IP -Porta $pt -TimeoutMs 800) { $id.Vivo = $true; break } }
    }
    if (-not $id.Vivo) { return $id }

    # IPP só quando o SNMP falhou ou não trouxe número de série
    if (-not $id.Serie -and $UsarIpp) {
        $ipp = Get-IppAtributos -IP $IP -TimeoutMs $TimeoutSondaMs
        if ($ipp) {
            $id.Ipp = $ipp
            $u = Normalize-Uuid $ipp['printer-uuid']
            if ($u) { $id.IppUuid = $u; $id.UuidFonte = 'IPP' }
            if ($ipp['printer-name']) { $id.IppNome = [string]$ipp['printer-name'] }
            if (-not $id.Modelo -and $ipp['printer-make-and-model']) {
                $id.Modelo = [string]$ipp['printer-make-and-model']; $id.ModeloFonte = 'IPP'
            }
        }
    }

    # PJL: série/MAC/modelo e texto de identificação (só se ainda faltar algo)
    if ($UsarPjl -and (-not $id.Serie -or -not $id.Modelo)) {
        $pjl = Get-PjlInfo -IP $IP -TimeoutMs $TimeoutSondaMs
        if ($pjl) {
            $id.Pjl = $pjl
            if ($pjl.Serie -and -not $id.Serie) { $id.Serie = $pjl.Serie; $id.SerieFonte = 'PJL' }
            if ($pjl.Mac -and -not $id.Mac) { $id.Mac = $pjl.Mac }
            if ($pjl.Modelo -and -not $id.Modelo) { $id.Modelo = $pjl.Modelo; $id.ModeloFonte = 'PJL' }
        }
    }

    $id.Mac = if ($id.Mac) { $id.Mac } else { Get-MacPorIP -IP $IP }
    $id.Nome = Get-NomeDoNoPorIP -IP $IP
    if ($id.Nome) { $id.HostNames += [string]$id.Nome }

    # HTTP (LEDM/páginas) quando ainda faltar identidade forte
    $marcaProvisoria = Get-MarcaPorTexto -Texto "$($id.Modelo) $($id.Descr) $($id.SysName) $($id.Nome) $($id.IppNome)"
    if ($UsarHttpIdentidade -and (-not $id.Serie -or -not $id.Modelo -or -not $marcaProvisoria)) {
        $http = Get-HttpIdentidade -IP $IP -Marca $marcaProvisoria -TimeoutMs $TimeoutSondaMs
        if ($http) {
            $id.Http = $http
            if ($http.Serie -and -not $id.Serie) { $id.Serie = $http.Serie; $id.SerieFonte = 'HTTP' }
            if ($http.Modelo -and -not $id.Modelo) { $id.Modelo = $http.Modelo; $id.ModeloFonte = 'HTTP' }
            if ($http.Uuid -and -not $id.DeviceGuid) { $id.DeviceGuid = $http.Uuid; $id.UuidFonte = 'HTTP' }
            if ($http.Mac -and -not $id.Mac) { $id.Mac = $http.Mac }
        }
    }

    $partes = @($id.Modelo, $id.Descr, $id.SysName, $id.Nome, $id.IppNome)
    if ($id.Ipp) { $partes += @($id.Ipp['printer-info'], $id.Ipp['printer-make-and-model'], $id.Ipp['printer-device-id'], $id.Ipp['printer-dns-sd-name']) }
    if ($id.Pjl -and $id.Pjl.Texto) { $partes += [string]$id.Pjl.Texto }
    if ($id.Http -and $id.Http.Texto) { $partes += [string]$id.Http.Texto }
    $txt = (@($partes) | Where-Object { $_ }) -join ' '
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
    foreach ($c in (@($Dev.Nome, $Dev.SysName, $Dev.IppNome) + @($Dev.HostNames) + @($Extras))) {
        if ($c) {
            $nome = ([string]$c).Trim()
            if (-not $nome -or (Test-EhIPv4 $nome)) { continue }
            $lista += $nome
            if ($nome -notmatch '\.') { $lista += "$nome.local" }
        }
    }
    $lista = @($lista | Where-Object { $_ } | Select-Object -Unique)
    foreach ($c in $lista) {
        $ipc = Get-IPdoNome -Nome $c
        if (-not $ipc) { continue }
        if ($ipc -eq $Dev.IP) { return $c }
        $x = Get-IdentidadeCache -IP $ipc
        if (Test-MesmoAparelho -A $Dev -B $x) { return $c }
    }
    if ($Forcar -and $lista.Count -gt 0) { return $lista[0] }
    return $null
}

# Acha na lista de aparelhos o que corresponde ao registro do inventário
# (série > UUID IPP/WSD/DeviceGUID/UPnP > MAC > hostname)
function Find-NoInventario {
    param($Inv, $Lista)
    $serieInv = Normalize-Serie $Inv.Serie
    if ($serieInv) {
        $d = @($Lista | Where-Object { (Normalize-Serie $_.Serie) -eq $serieInv }) | Select-Object -First 1
        if ($d) { return [PSCustomObject]@{ Dev = $d; Criterio = 'número de série' } }
    }
    foreach ($par in @(
        @{ Campo = 'IppUuid'; Criterio = 'UUID IPP' },
        @{ Campo = 'WsdUuid'; Criterio = 'UUID WSD' },
        @{ Campo = 'DeviceGuid'; Criterio = 'DeviceGUID' },
        @{ Campo = 'UpnpUuid'; Criterio = 'UUID UPnP' }
    )) {
        $v = Normalize-Uuid $Inv.($par.Campo)
        if ($v) {
            $d = @($Lista | Where-Object { (Normalize-Uuid $_.($par.Campo)) -eq $v }) | Select-Object -First 1
            if ($d) { return [PSCustomObject]@{ Dev = $d; Criterio = $par.Criterio } }
        }
    }
    $macInv = Normalize-Mac $Inv.Mac
    if ($macInv) {
        $d = @($Lista | Where-Object {
            (Normalize-Mac $_.Mac) -eq $macInv -and
            -not ((Normalize-Serie $_.Serie) -and $serieInv -and (Normalize-Serie $_.Serie) -ne $serieInv)
        }) | Select-Object -First 1
        if ($d) { return [PSCustomObject]@{ Dev = $d; Criterio = 'MAC' } }
    }
    $nomesInv = @(@($Inv.HostNames) + @($Inv.Nome) | Where-Object { $_ } | ForEach-Object { Normalize-Nome $_ } | Where-Object { $_ })
    if ($nomesInv.Count -gt 0) {
        $d = @($Lista | Where-Object {
            $nomesDev = @(@($_.Nome, $_.SysName) + @($_.HostNames) + @($_.IppNome) |
                Where-Object { $_ } | ForEach-Object { Normalize-Nome $_ } | Where-Object { $_ })
            @($nomesDev | Where-Object { $nomesInv -contains $_ }).Count -gt 0 -and
            -not ((Normalize-Serie $_.Serie) -and $serieInv -and (Normalize-Serie $_.Serie) -ne $serieInv)
        }) | Select-Object -First 1
        if ($d) { return [PSCustomObject]@{ Dev = $d; Criterio = 'nome do nó' } }
    }
    return $null
}

# Devolve o nome da fila que já usa este nó com OUTRO aparelho
function Test-ConflitoNo {
    param([string]$Nome, [string]$IdIP, [string]$IdMac, [string]$IdSerie,
          [string]$IdIppUuid, [string]$IdWsdUuid, [string]$IdUpnpUuid, [string]$IdDeviceGuid)
    $chaveNome = Normalize-Nome $Nome
    if (-not $chaveNome) { return $null }
    $r = $nomesAtribuidos[$chaveNome]
    if (-not $r) { return $null }
    $sa = Normalize-Serie $IdSerie
    $sb = Normalize-Serie $r.Serie
    if ($sa -and $sb) { if ($sa -ne $sb) { return $r.Fila } else { return $null } }
    $uas = @(); $ubs = @()
    foreach ($par in @(
        @{ A = $IdIppUuid; B = $r.IppUuid },
        @{ A = $IdWsdUuid; B = $r.WsdUuid },
        @{ A = $IdUpnpUuid; B = $r.UpnpUuid },
        @{ A = $IdDeviceGuid; B = $r.DeviceGuid }
    )) {
        $va = Normalize-Uuid $par.A
        $vb = Normalize-Uuid $par.B
        if ($va) { $uas += $va }
        if ($vb) { $ubs += $vb }
    }
    $uuidComparado = ($uas.Count -gt 0 -and $ubs.Count -gt 0)
    if ($uuidComparado -and @($uas | Where-Object { $ubs -contains $_ }).Count -gt 0) { return $null }

    $ma = Normalize-Mac $IdMac
    $mb = Normalize-Mac $r.Mac
    if ($ma -and $mb) { if ($ma -ne $mb) { return $r.Fila } else { return $null } }
    if ($uuidComparado) { return $r.Fila }
    if ($IdIP -and $r.IP -and $IdIP -ne $r.IP) {
        if ($ma -and $mb -and $ma -eq $mb) { return $null }
        return $r.Fila
    }
    return $null
}

# Procura impressoras por multicast (WSD/mDNS/SSDP) e por varrimento TCP nas sub-redes locais
function Get-ImpressorasNaRede {
    Write-Host "Procurando impressoras na rede local (pode levar alguns segundos)..." -ForegroundColor Cyan

    $achadas = New-Object System.Collections.Generic.List[object]
    $porIp = @{}

    # --- Multicast (mesma L2): WSD, mDNS e SSDP ---
    $extras = @()
    if ($UsarWsd)  { try { $extras += @(Get-WsdDevices -TimeoutMs 2500) } catch { } }
    if ($UsarMdns) { try { $extras += @(Get-MdnsPrinters -TimeoutMs 2500) } catch { } }
    if ($UsarUpnp) { try { $extras += @(Get-SsdpPrinters -TimeoutMs 2000) } catch { } }

    foreach ($e in $extras) {
        if (-not $e) { continue }
        $eip = [string]$e.IP
        if (-not $eip -or $eip -match ':') { continue }
        if (-not $porIp.ContainsKey($eip)) {
            $id = Get-IdentidadeCache -IP $eip
            $porIp[$eip] = $id
            $achadas.Add($id)
        }
        $id = $porIp[$eip]
        if ($e.PSObject.Properties.Name -contains 'Uuid' -and $e.Uuid) { $id.WsdUuid = $e.Uuid }
        if ($e.PSObject.Properties.Name -contains 'IppUuid' -and $e.IppUuid) { $id.IppUuid = $e.IppUuid }
        if ($e.PSObject.Properties.Name -contains 'UpnpUuid' -and $e.UpnpUuid) { $id.UpnpUuid = $e.UpnpUuid }
        $id = Merge-Descoberta -Base $id -Extra $e
        $id.Vivo = $true
        if ($id.HostNames -and $id.HostNames.Count -gt 0 -and -not $id.Nome) { $id.Nome = @($id.HostNames)[0] }
    }

    # --- Varrimento TCP nas sub-redes /24 ---
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

    foreach ($ip in $abertos) {
        if ($porIp.ContainsKey($ip)) { continue }
        $id = Get-IdentidadeCache -IP $ip
        $porIp[$ip] = $id
        $achadas.Add($id)
    }

    if ($achadas.Count -gt 0) {
        Write-Host "Dispositivos de impressão encontrados:" -ForegroundColor Cyan
        $achadas | Format-Table IP, Nome, Marca, Modelo, Serie, Mac, IppUuid, WsdUuid -AutoSize | Out-String | Write-Host
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
    param([string]$Fila, [string]$Driver, [string]$Marca, [string]$Nome, [string]$IP, [string]$Mac,
          [string]$Serie, [string]$Modelo, [string]$IppUuid, [string]$WsdUuid, [string]$UpnpUuid,
          [string]$DeviceGuid, [string[]]$HostNames, [string]$Fonte)
    if (-not $Serie -and -not $Mac -and -not $IP -and -not $IppUuid -and -not $WsdUuid -and -not $UpnpUuid -and -not $DeviceGuid) { return }
    $velho = $inventario[$Fila]
    $novo = [PSCustomObject]@{
        Fila = $Fila; Driver = $Driver; Marca = $Marca; Nome = $Nome; IP = $IP; Mac = Normalize-Mac $Mac
        Serie = Normalize-Serie $Serie; Modelo = $Modelo; IppUuid = Normalize-Uuid $IppUuid
        WsdUuid = Normalize-Uuid $WsdUuid; UpnpUuid = Normalize-Uuid $UpnpUuid
        DeviceGuid = Normalize-Uuid $DeviceGuid; HostNames = @()
        Fonte = $Fonte; UltimoIP = $IP; UltimaVezVisto = (Get-Date -Format 's'); Atualizado = (Get-Date -Format 's')
    }
    $hosts = @()
    foreach ($h in @($HostNames) + @($Nome)) { if ($h) { $hosts += [string]$h } }
    if ($velho) {
        foreach ($campo in 'Driver', 'Marca', 'Nome', 'IP', 'Mac', 'Serie', 'Modelo', 'IppUuid', 'WsdUuid', 'UpnpUuid', 'DeviceGuid', 'Fonte') {
            if (-not $novo.$campo -and $velho.$campo) { $novo.$campo = $velho.$campo }
        }
        foreach ($h in @($velho.HostNames) + @($velho.Nome)) { if ($h) { $hosts += [string]$h } }
        if (-not $novo.UltimoIP -and $velho.UltimoIP) { $novo.UltimoIP = $velho.UltimoIP }
    }
    $novo.HostNames = @($hosts | Where-Object { $_ } | Select-Object -Unique)
    $inventario[$Fila] = $novo
}

function Write-Inventario {
    if (-not $GravarInventario -or $inventario.Count -eq 0) { return }
    try {
        if (-not (Test-Path $pastaInv)) { New-Item -ItemType Directory -Path $pastaInv -Force | Out-Null }
        ConvertTo-Json -InputObject @($inventario.Values) -Depth 4 | Set-Content -Path $arqInv -Encoding UTF8
        Write-Host "Inventário salvo em: $arqInv" -ForegroundColor Cyan
    } catch { Write-Host "Aviso: não consegui gravar o inventário: $($_.Exception.Message)" -ForegroundColor Yellow }
}

# --------------------------------------------------- Coleta das filas (1) ----

$wsdPorts = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\WSD Port\Ports\*" -ErrorAction SilentlyContinue |
    ForEach-Object { $wsdPorts[$_.PSChildName] = $_ }
$hpPorts = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\HP Standard TCP/IP Port\Ports\*" -ErrorAction SilentlyContinue |
    ForEach-Object { $hpPorts[$_.PSChildName] = $_ }
$ippPorts = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\IPP Monitor\Ports\*" -ErrorAction SilentlyContinue |
    ForEach-Object { $ippPorts[$_.PSChildName] = $_ }

function Get-InfoFila {
    param($Imp)
    $porta = Get-PrinterPort -Name $Imp.PortName -ErrorAction SilentlyContinue
    $regStd = Get-ItemProperty -LiteralPath "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports\$($Imp.PortName)" -ErrorAction SilentlyContinue
    $regHp  = $hpPorts[$Imp.PortName]
    $regIpp = $ippPorts[$Imp.PortName]
    $wsd    = $wsdPorts[$Imp.PortName]

    $f = [PSCustomObject]@{
        Imp = $Imp; Porta = $porta; Reg = $null; RegTipo = $null; Tipo = $null
        HostAddr = $porta.PrinterHostAddress; Ip = $null; Nome = $null
        Monitor = $(if ($porta) { [string]$porta.PortMonitor } else { $null })
        Marca = (Get-MarcaImpressora -Imp $Imp)
        Fax = ($IgnorarFax -and ($Imp.Name -match 'fax' -or $Imp.DriverName -match 'fax'))
        Ident = $null; Diverge = $false
        WsdUuid = $null; Ipp = $false; MonitorHP = $false
    }

    if ($Imp.PortName -match '^WSD' -or ($porta -and $porta.Description -match 'WSD')) {
        $f.Tipo = 'WSD'
        if ($wsd) {
            $f.WsdUuid = Normalize-Uuid $wsd.'Printer UUID'
            if ($wsd.IPAddress) { $f.Ip = [string]$wsd.IPAddress }
            if ($wsd.HostName)  { $f.Nome = [string]$wsd.HostName }
        }
        if (-not $f.Ip -or -not $f.Nome) {
            $pnp = Get-PnpDeviceProperty -InstanceId "*$($Imp.PortName)*" -ErrorAction SilentlyContinue
            if (-not $f.Ip)   { $f.Ip   = ($pnp | Where-Object KeyName -match 'IPAddress').Data }
            if (-not $f.Nome) { $f.Nome = ($pnp | Where-Object KeyName -match 'HostName').Data }
        }
        if (-not $f.Ip -and $f.WsdUuid -and $UsarWsd) {
            $r = Resolve-WsdDevice -Uuid $f.WsdUuid -TimeoutMs ([math]::Max(1500, $TimeoutSondaMs * 2))
            if ($r) {
                $f.Ip = $r.Ip
                if (-not $f.Nome) { $f.Nome = $r.Ip }
            }
        }
        if ($f.Nome) { $f.HostAddr = $f.Nome }
        elseif ($f.Ip) { $f.HostAddr = $f.Ip }
    }
    elseif (($porta -and ($porta.PortMonitor -match 'IPP' -or $porta.Description -match 'IPP')) -or $regIpp) {
        $f.Tipo = 'IPP'; $f.Ipp = $true
        $hostAddr = $porta.PrinterHostAddress
        if (-not $hostAddr -and $regIpp) {
            if ($regIpp.HostName) { $hostAddr = $regIpp.HostName }
            elseif ($regIpp.IPAddress) { $hostAddr = $regIpp.IPAddress }
        }
        if ($hostAddr) {
            $f.HostAddr = $hostAddr
            if (Test-EhIPv4 $hostAddr) { $f.Ip = $hostAddr }
            else { $f.Nome = $hostAddr; $f.Ip = Get-IPdoNome -Nome $hostAddr }
        }
    }
    else {
        $hostAddr = $porta.PrinterHostAddress
        if (-not $hostAddr -and $regHp) {
            if ($regHp.HostName) { $hostAddr = $regHp.HostName }
            elseif ($regHp.IPAddress) { $hostAddr = $regHp.IPAddress }
        }
        if (-not $hostAddr -and $regStd) {
            if ($regStd.HostName) { $hostAddr = $regStd.HostName }
            elseif ($regStd.IPAddress) { $hostAddr = $regStd.IPAddress }
        }
        if ($hostAddr) {
            $f.HostAddr = $hostAddr
            $f.MonitorHP = [bool]$regHp
            if ($regHp) { $f.Reg = $regHp; $f.RegTipo = 'HP' }
            elseif ($regStd) { $f.Reg = $regStd; $f.RegTipo = 'Standard' }
            $monitorSuportado = (-not $f.Monitor -or $f.Monitor -in 'Standard TCP/IP Port', 'HP Standard TCP/IP Port')
            if (Test-EhIPv4 $hostAddr) { $f.Ip = $hostAddr }
            else { $f.Nome = $hostAddr; $f.Ip = Get-IPdoNome -Nome $hostAddr }
            if (-not $monitorSuportado) { $f.Tipo = 'Monitor proprietário' }
            elseif (Test-EhIPv4 $hostAddr) { $f.Tipo = 'TCP/IP (IP estático)' }
            else { $f.Tipo = 'TCP/IP (já usa nome)' }
        }
        elseif (-not $porta -or $Imp.PortName -match '^(PORTPROMPT|FILE|nul|SHRFAX|XPS|PDF|USB|LPT\d|COM\d|TS\d|Microsoft|\\\\\.\\pipe)' -or $porta.Description -match 'Local|Fax|Arquivo|File') {
            $f.Tipo = 'Local/Virtual'
        }
        else { $f.Tipo = 'Monitor proprietário' }
    }
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

$comunidadesPortas = @(Get-PrinterPort -ErrorAction SilentlyContinue |
    Where-Object { $_.SNMPCommunity } | ForEach-Object { $_.SNMPCommunity })
$comunidadesHp = @($hpPorts.Values | ForEach-Object { $_.'SNMP Community' })
$comunidadesTodas = @(@($Comunidades) + $comunidadesPortas + $comunidadesHp | Where-Object { $_ } | Select-Object -Unique)
if ($comunidadesTodas.Count -eq 0) { $comunidadesTodas = @('public') }
$cacheIdent = @{}

$filas = @(foreach ($imp in $impressoras) { Get-InfoFila -Imp $imp })

# --- Identidade do aparelho que responde no IP de cada fila ---
foreach ($f in $filas) {
    if ($f.Fax -or $f.Tipo -notin 'TCP/IP (IP estático)', 'TCP/IP (já usa nome)', 'WSD', 'IPP', 'Monitor proprietário' -or -not $f.Ip) { continue }
    $f.Ident = Get-IdentidadeCache -IP $f.Ip
    if ($f.WsdUuid) { $f.Ident.WsdUuid = $f.WsdUuid }
    if ($f.Reg -and $f.Reg.DeviceGUID) { $f.Ident.DeviceGuid = Normalize-Uuid $f.Reg.DeviceGUID }
    $inv = $inventario[$f.Imp.Name]
    if ($f.Ident.Vivo) {
        $motivo = $null
        if ($inv -and $inv.Serie -and $f.Ident.Serie -and (Normalize-Serie $inv.Serie) -ne (Normalize-Serie $f.Ident.Serie)) {
            $motivo = "S/N $($f.Ident.Serie) difere do inventário ($($inv.Serie))"
        }
        elseif ($inv -and $inv.IppUuid -and $f.Ident.IppUuid -and (Normalize-Uuid $inv.IppUuid) -ne (Normalize-Uuid $f.Ident.IppUuid)) {
            $motivo = "UUID IPP difere do inventário"
        }
        elseif ($inv -and $inv.WsdUuid -and $f.Ident.WsdUuid -and (Normalize-Uuid $inv.WsdUuid) -ne (Normalize-Uuid $f.Ident.WsdUuid)) {
            $motivo = "UUID WSD difere do inventário"
        }
        elseif ($f.Marca -and $f.Ident.Marca -and $f.Marca -ne $f.Ident.Marca) {
            $motivo = "marca $($f.Ident.Marca) difere da fila ($($f.Marca))"
        }
        elseif ($f.Ident.Modelo -and (Test-ModeloCombina -Driver $f.Imp.DriverName -TextoNorm $f.Ident.Texto) -eq $false) {
            $motivo = "modelo '$($f.Ident.Modelo)' não combina com o driver '$($f.Imp.DriverName)'"
        }
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
        $chaveHost = Normalize-Nome $f.HostAddr
        if ($chaveHost) {
            $nomesAtribuidos[$chaveHost] = @{
                IP = $f.Ip; Mac = $f.Ident.Mac; Serie = $f.Ident.Serie; IppUuid = $f.Ident.IppUuid
                WsdUuid = $f.Ident.WsdUuid; UpnpUuid = $f.Ident.UpnpUuid; DeviceGuid = $f.Ident.DeviceGuid
                Fila = $f.Imp.Name
            }
        }
        if ($f.Marca) {
            if (-not $conhecidos.ContainsKey($f.Marca)) { $conhecidos[$f.Marca] = @() }
            if ($conhecidos[$f.Marca] -notcontains $f.HostAddr) { $conhecidos[$f.Marca] += $f.HostAddr }
        }
    }
    if ($f.Marca -and -not $f.Fax -and $f.Tipo -in 'TCP/IP (IP estático)', 'TCP/IP (já usa nome)', 'WSD', 'IPP', 'Monitor proprietário') {
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

# Aparelhos já ligados a uma fila, para não os entregar a outra fila por engano
$atribuidos = New-Object System.Collections.ArrayList
foreach ($f in $filas) {
    if ($f.Fax -or $f.Diverge) { continue }
    if (-not ($f.Ident -and $f.Ident.Vivo)) { continue }
    $nomeAt = if ($f.Nome) { $f.Nome } elseif ($f.HostAddr) { $f.HostAddr } else { $null }
    [void]$atribuidos.Add([PSCustomObject]@{
        Fila = $f.Imp.Name; IP = $(if ($f.Ip) { $f.Ip } else { $f.Ident.IP }); Nome = $nomeAt
        Serie = $f.Ident.Serie; Mac = $f.Ident.Mac; IppUuid = $f.Ident.IppUuid
        WsdUuid = $f.Ident.WsdUuid; UpnpUuid = $f.Ident.UpnpUuid; DeviceGuid = $f.Ident.DeviceGuid
    })
}

$resultado = New-Object System.Collections.Generic.List[object]
$portasAntigas = New-Object System.Collections.Generic.HashSet[string]
$descobertas = $null

foreach ($f in $filas) {

    $imp = $f.Imp; $m = $f.Marca; $tipo = $f.Tipo
    $ip = $f.Ip; $nome = $null; $nomePorta = $null; $dev = $null; $ipAtual = $null
    $status = $null; $novaPorta = $null; $origemNome = $null; $porIP = $false
    $serieDev = $null; $macDev = $null; $ippUuidDev = $null; $wsdUuidDev = $null; $upnpUuidDev = $null; $deviceGuidDev = $null

    # Captura de identidade (inventário) mesmo para filas que não serão migradas
    if (-not $f.Fax -and $f.Ident -and $f.Ident.Vivo -and -not $f.Diverge -and
        ($f.Ident.Serie -or $f.Ident.Mac -or $f.Ident.IppUuid -or $f.Ident.WsdUuid -or $f.Ident.DeviceGuid)) {
        $nomeCap = if ($f.Nome -and -not (Test-EhIPv4 $f.Nome)) { $f.Nome } elseif ($f.Ident.Nome) { $f.Ident.Nome } else { $f.Ident.SysName }
        $hostsCap = @(@($f.Ident.HostNames) + @($f.Nome) + @($f.Ident.Nome) + @($f.Ident.SysName) |
            Where-Object { $_ -and -not (Test-EhIPv4 $_) } | Select-Object -Unique)
        $marcaCap = if ($m) { $m } elseif ($f.Ident.Marca) { $f.Ident.Marca } else { $null }
        Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $marcaCap -Nome $nomeCap -IP $f.Ip `
            -Mac $f.Ident.Mac -Serie $f.Ident.Serie -Modelo $f.Ident.Modelo `
            -IppUuid $f.Ident.IppUuid -WsdUuid $f.Ident.WsdUuid -UpnpUuid $f.Ident.UpnpUuid `
            -DeviceGuid $f.Ident.DeviceGuid -HostNames $hostsCap -Fonte 'captura'
    }
    $inv = $inventario[$imp.Name]
    if (-not $m -and $inv -and $inv.Marca) { $m = $inv.Marca }

    if ($f.Fax) { $status = 'Ignorada: impressora de fax' }
    elseif ($tipo -eq 'Local/Virtual') { $status = 'Ignorada: porta local/virtual' }
    elseif ($tipo -eq 'Monitor proprietário') {
        $status = "Ignorada: porta de monitor proprietário ('$($f.Monitor)'), ainda não suportada"
    }
    elseif ($tipo -eq 'IPP') {
        $status = 'Ignorada: porta IPP (identidade registada; migração para nome não suportada)'
    }
    elseif ($f.MonitorHP -and -not $MigrarHPStandardTCPIP) {
        $status = 'Ignorada: monitor HP Standard TCP/IP (migração desativada)'
    }
    else {
        if ($tipo -eq 'TCP/IP (já usa nome)') { $status = 'Ignorada: já usa nome do nó' }
        else {
            $bloqueioMotivo = $null; $motivoFalha = $null

            # 1) Mapa manual
            if ($Manual.ContainsKey($imp.Name)) {
                $nome = $Manual[$imp.Name]; $origemNome = 'manual'
                $ipm = Get-IPdoNome -Nome $nome
                if ($ipm) { $dev = Get-IdentidadeCache -IP $ipm }
            }
            # 2) O aparelho ainda responde no IP da fila e é o registrado: usa um nome DELE que resolva para o IP dele
            elseif ($f.Ident -and $f.Ident.Vivo -and -not $f.Diverge) {
                $dev = $f.Ident
                $nome = Select-NomeValido -Dev $dev -Extras (@($inv.Nome) + @($inv.HostNames) + @($f.Nome) + @($f.HostAddr))
                if ($nome) { $origemNome = 'aparelho no IP da fila' }
                else { $motivoFalha = "o aparelho em $ip responde, mas nenhum nome dele resolve para o IP dele" }
            }

            # 3) Inventário: acha o aparelho na rede (série > UUID > MAC > nome)
            $temInv = $inv -and ($inv.Serie -or $inv.Mac -or $inv.IppUuid -or $inv.WsdUuid -or $inv.DeviceGuid -or $inv.UpnpUuid -or $inv.Nome -or @($inv.HostNames).Count -gt 0)
            if (-not $nome -and -not $bloqueioMotivo -and $BuscarNaRede -and $temInv) {
                if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
                $ach = Find-NoInventario -Inv $inv -Lista $descobertas
                if ($ach) {
                    $dev = $ach.Dev
                    $n3 = Select-NomeValido -Dev $ach.Dev -Extras (@($inv.Nome) + @($inv.HostNames))
                    if ($n3) { $nome = $n3; $origemNome = "inventário: achado na rede por $($ach.Criterio), IP atual $($ach.Dev.IP)" }
                    else { $motivoFalha = "achei o aparelho do inventário em $($ach.Dev.IP), mas nenhum nome dele resolve para esse IP" }
                }
            }

            # 4) Candidatos por marca e modelo (sem o antigo bloqueio por quantidade da marca)
            if (-not $nome -and -not $bloqueioMotivo -and $BuscarNaRede -and $m) {
                if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
                $serieInv = if ($inv) { Normalize-Serie $inv.Serie } else { $null }
                $cands = @($descobertas | Where-Object {
                    (-not $_.Marca -or $_.Marca -eq $m) -and -not ($serieInv -and (Normalize-Serie $_.Serie) -and (Normalize-Serie $_.Serie) -ne $serieInv)
                })
                $cands = @($cands | Where-Object {
                    -not (Test-CandidatoAtribuido -Cand $_ -Atribuidos $atribuidos -FilaAtual $imp.Name)
                })

                # Remove duplicados do mesmo aparelho (mesma série/UUID/MAC), mantendo IPs distintos
                $unicos = @{}
                $lista = @()
                foreach ($c in $cands) {
                    $chave = if (Normalize-Serie $c.Serie) { 'S:' + (Normalize-Serie $c.Serie) }
                             elseif (Normalize-Uuid $c.IppUuid) { 'I:' + (Normalize-Uuid $c.IppUuid) }
                             elseif (Normalize-Uuid $c.WsdUuid) { 'W:' + (Normalize-Uuid $c.WsdUuid) }
                             elseif (Normalize-Uuid $c.DeviceGuid) { 'G:' + (Normalize-Uuid $c.DeviceGuid) }
                             elseif (Normalize-Mac $c.Mac) { 'M:' + (Normalize-Mac $c.Mac) }
                             else { 'P:' + $c.IP }
                    if (-not $unicos.ContainsKey($chave)) { $unicos[$chave] = $true; $lista += $c }
                }
                $cands = $lista

                $verificados = @()
                $naoVerificados = @()
                $descartados = @()
                foreach ($c in $cands) {
                    $mod = Test-ModeloCombina -Driver $imp.DriverName -TextoNorm $c.Texto
                    if ($mod -eq $true) { $verificados += $c }
                    elseif ($null -eq $mod) { $naoVerificados += $c }
                    else { $descartados += $c }
                }
                $modeloStatus = $null
                if ($verificados.Count -gt 0) { $cands = $verificados; $modeloStatus = 'confirmado' }
                elseif ($naoVerificados.Count -gt 0) { $cands = $naoVerificados; $modeloStatus = 'nao_verificavel' }
                elseif ($descartados.Count -gt 0) { $cands = @(); $modeloStatus = 'nao_confere' }

                if ($cands.Count -eq 1 -and $modeloStatus -ne 'nao_confere') {
                    $dev = $cands[0]
                    $obs = if ($modeloStatus -eq 'confirmado') { 'modelo confirmado' } else { 'modelo não verificável' }
                    $n4 = Select-NomeValido -Dev $dev -Extras $(if ($inv) { @($inv.Nome) + @($inv.HostNames) } else { @() })
                    if ($n4) { $nome = $n4; $origemNome = "achada na rede por marca ($obs)" }
                    else { $motivoFalha = "achei uma $m em $($dev.IP), mas nenhum nome dela resolve para esse IP" }
                }
                elseif ($cands.Count -gt 1) {
                    $motivoFalha = "$($cands.Count) impressoras $m candidatas ($(($cands.IP) -join ', ')); sem como escolher sozinho"
                }
                elseif ($modeloStatus -eq 'nao_confere') {
                    $motivoFalha = "nenhuma impressora $m com modelo compatível com o driver '$($imp.DriverName)'"
                }
                else { $motivoFalha = "nenhuma impressora $m encontrada na rede" }
            }

            # 5) Modo assistido: pergunta no terminal (só interativo e fora da simulação)
            if (-not $nome -and -not $bloqueioMotivo -and $ModoAssistido -and $BuscarNaRede -and -not $Simular -and $Host.Name -match 'ConsoleHost|ISE') {
                if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
                $cands5 = @($descobertas | Where-Object { ($_.Snmp -or $_.Marca) -and (-not $m -or -not $_.Marca -or $_.Marca -eq $m) })
                $esc = Invoke-Assistido -Fila $f -Cands $cands5
                if ($esc) {
                    $dev = $esc
                    $n5 = Select-NomeValido -Dev $esc
                    if ($n5) { $nome = $n5; $origemNome = 'escolha assistida'; $motivoFalha = $null }
                    else { $motivoFalha = "o aparelho escolhido ($($esc.IP)) não tem nome que resolva para o IP dele" }
                }
            }

            # Sem hostname: porta por IP atual (com aviso), se já houver um aparelho identificado
            if (-not $nome -and -not $bloqueioMotivo -and $PermitirPortaPorIP -and $dev -and (Test-EhIPv4 $dev.IP) -and -not $f.Diverge) {
                $nome = [string]$dev.IP
                $porIP = $true
                $sufixo = Normalize-Serie $dev.Serie
                if (-not $sufixo) { $sufixo = $nome -replace '[^\w.-]', '_' }
                $nomePorta = "IP_$sufixo"
                $origemNome = 'porta por IP (sem hostname; reexecute se o IP mudar)'
                $motivoFalha = $null
            }

            # Sem nome: explica por quê
            if (-not $nome) {
                if ($bloqueioMotivo) { $status = "Ignorada: $bloqueioMotivo" }
                elseif ($motivoFalha) { $status = "Ignorada: $motivoFalha" }
                else { $status = 'Ignorada: não consegui identificar o aparelho desta fila (sem inventário e sem resposta no IP antigo)' }
            }
            else {
                $serieDev = Normalize-Serie $(if ($dev) { $dev.Serie } else { $null })
                if (-not $serieDev -and $inv) { $serieDev = Normalize-Serie $inv.Serie }
                $macDev = Normalize-Mac $(if ($dev) { $dev.Mac } else { $null })
                if (-not $macDev -and $inv) { $macDev = Normalize-Mac $inv.Mac }
                $ippUuidDev = Normalize-Uuid $(if ($dev) { $dev.IppUuid } else { $null })
                if (-not $ippUuidDev -and $inv) { $ippUuidDev = Normalize-Uuid $inv.IppUuid }
                $wsdUuidDev = Normalize-Uuid $(if ($dev) { $dev.WsdUuid } else { $null })
                if (-not $wsdUuidDev -and $inv) { $wsdUuidDev = Normalize-Uuid $inv.WsdUuid }
                $upnpUuidDev = Normalize-Uuid $(if ($dev) { $dev.UpnpUuid } else { $null })
                if (-not $upnpUuidDev -and $inv) { $upnpUuidDev = Normalize-Uuid $inv.UpnpUuid }
                $deviceGuidDev = Normalize-Uuid $(if ($dev) { $dev.DeviceGuid } else { $null })
                if (-not $deviceGuidDev -and $inv) { $deviceGuidDev = Normalize-Uuid $inv.DeviceGuid }
                $modeloDev = if ($dev -and $dev.Modelo) { $dev.Modelo } elseif ($inv) { $inv.Modelo } else { $null }
                $hostsDev  = @(@(if ($dev) { $dev.HostNames }) + @(if ($inv) { $inv.HostNames }) | Where-Object { $_ } | Select-Object -Unique)
                $ipAtual  = if ($porIP) { $nome } elseif ($dev -and $dev.IP) { $dev.IP } else { Get-IPdoNome -Nome $nome }
                if ($porIP -and $serieDev) { $nomePorta = "IP_$($serieDev -replace '[^\w.-]', '_')" }

                # Confere o aparelho contra o inventário e contra outras filas
                if ($origemNome -ne 'manual' -and $dev -and $inv -and (Test-IdentidadeConflita -A $dev -B $inv)) {
                    $status = "Ignorada: o aparelho encontrado não é o registrado para esta fila (identidade diverge do inventário)"
                }
                $outraFila = $null
                if (-not $status -and $origemNome -ne 'manual') {
                    $outraFila = Test-ConflitoNo -Nome $nome -IdIP $ipAtual -IdMac $macDev -IdSerie $serieDev `
                        -IdIppUuid $ippUuidDev -IdWsdUuid $wsdUuidDev -IdUpnpUuid $upnpUuidDev -IdDeviceGuid $deviceGuidDev
                }
                if (-not $status) {
                    if (-not $porIP -and -not (Get-IPdoNome -Nome $nome) -and -not $Forcar) { $status = "Ignorada: '$nome' não resolve neste PC (use `$Forcar = `$true para ignorar)" }
                    elseif ($outraFila) { $status = "Ignorada: o nó '$nome' já é usado pela fila '$outraFila' com outro aparelho; use `$Manual se forem o mesmo" }
                }

                if (-not $status) {
                    $novaPorta = if ($nomePorta) { $nomePorta } else { $nome }
                    $extra = if ($origemNome) { " [$origemNome]" } else { '' }
                    $nomeInv = if ($porIP) { $null } else { $nome }
                    if ($porIP) { $extra = " [PORTA POR IP $nome; aviso: se o IP mudar, reexecute o script]" }

                    if ($Simular) {
                        $status = "SIMULAÇÃO: trocaria '$($imp.PortName)' por '$novaPorta'$extra"
                        Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $m -Nome $nomeInv -IP $ipAtual `
                            -Mac $macDev -Serie $serieDev -Modelo $modeloDev -IppUuid $ippUuidDev `
                            -WsdUuid $wsdUuidDev -UpnpUuid $upnpUuidDev -DeviceGuid $deviceGuidDev `
                            -HostNames $hostsDev -Fonte $(if ($porIP) { 'porta-ip' } else { 'porta-nome' })
                        $chaveNome = Normalize-Nome $nome
                        if ($chaveNome) {
                            $nomesAtribuidos[$chaveNome] = @{
                                IP = $ipAtual; Mac = $macDev; Serie = $serieDev; IppUuid = $ippUuidDev
                                WsdUuid = $wsdUuidDev; UpnpUuid = $upnpUuidDev; DeviceGuid = $deviceGuidDev; Fila = $imp.Name
                            }
                        }
                        [void]$atribuidos.Add([PSCustomObject]@{
                            Fila = $imp.Name; IP = $ipAtual; Nome = $nomeInv; Serie = $serieDev; Mac = $macDev
                            IppUuid = $ippUuidDev; WsdUuid = $wsdUuidDev; UpnpUuid = $upnpUuidDev; DeviceGuid = $deviceGuidDev
                        })
                    }
                    else {
                        try {
                            $portaAntiga = $imp.PortName
                            # Protocolo da porta antiga (Registro): 1 = RAW, 2 = LPR
                            $protReg = if ($f.Reg -and ([string]$f.Reg.Protocol -match '^(2|LPR)$')) { 2 } else { 1 }
                            $filaLpr = if ($f.Reg -and $f.Reg.Queue) { [string]$f.Reg.Queue } else { [string]$f.Porta.LprQueueName }
                            $usaLpr = ($protReg -eq 2) -or (-not $f.Reg -and $f.Porta -and $f.Porta.LprQueueName)
                            $tcp = if ($usaLpr) { 515 } elseif ($f.Reg -and $f.Reg.PortNumber) { [int]$f.Reg.PortNumber } elseif ($f.Porta.PortNumber) { [int]$f.Porta.PortNumber } else { 9100 }
                            if ($usaLpr -and -not $filaLpr) { throw "a porta antiga é LPR mas não achei o nome da fila LPR" }
                            $criou = $false
                            $aplicouPorta = $false

                            $portaNova = Get-PrinterPort -Name $novaPorta -ErrorAction SilentlyContinue
                            if ($portaNova) {
                                $addrNova = if ($usaLpr) { [string]$portaNova.LprHostAddress } else { [string]$portaNova.PrinterHostAddress }
                                $filaNova = if ($usaLpr) { [string]$portaNova.LprQueueName } else { $null }
                                $portaOk = ($addrNova -and $addrNova.Trim() -eq [string]$nome -and (-not $usaLpr -or -not $filaLpr -or $filaNova -eq $filaLpr))
                                if ($portaOk -and $usaLpr -and -not $portaNova.LprQueueName) { $portaOk = $false }
                                if ($portaOk -and -not $usaLpr -and $portaNova.LprQueueName) { $portaOk = $false }
                                if ($portaOk -and -not $usaLpr) {
                                    $portaOk = $(if ($portaNova.PortNumber) { [int]$portaNova.PortNumber -eq $tcp } else { $tcp -eq 9100 })
                                }
                                if (-not $portaOk) {
                                    $emUso = Get-Printer -ErrorAction SilentlyContinue | Where-Object PortName -eq $novaPorta | Select-Object -First 1
                                    if ($emUso -and $emUso.Name -eq $imp.Name) {
                                        $basePorta = $novaPorta
                                        $k = 2
                                        while (Get-PrinterPort -Name "$basePorta`_$k" -ErrorAction SilentlyContinue) { $k = $k + 1 }
                                        $novaPorta = "$basePorta`_$k"
                                        $portaNova = $null
                                    } elseif ($emUso) {
                                        throw "a porta '$novaPorta' já existe com endereço '$addrNova' e está em uso pela fila '$($emUso.Name)'"
                                    } else {
                                        Remove-PrinterPort -Name $novaPorta
                                        $portaNova = $null
                                    }
                                }
                            }
                            if (-not $portaNova) {
                                $p = @{ Name = $novaPorta }
                                if ($usaLpr) {
                                    $p.LprHostAddress = $nome
                                    $p.LprQueueName = $filaLpr
                                    if (($f.Reg -and [string]$f.Reg.'Byte Count' -eq '1') -or ($f.Porta -and $f.Porta.LprByteCounting)) { $p.LprByteCounting = $true }
                                } else {
                                    $p.PrinterHostAddress = $nome
                                    $p.PortNumber = $tcp
                                }
                                $snmpEnabled = ($f.Porta -and $f.Porta.SNMPEnabled) -or ($f.Reg -and ([string]$f.Reg.'SNMP Enabled' -eq '1'))
                                $snmpIndex = if ($f.Porta -and $f.Porta.SNMPIndex) { $f.Porta.SNMPIndex } elseif ($f.Reg) { $f.Reg.'SNMP Index' } else { $null }
                                $snmpCommunity = if ($f.Porta -and $f.Porta.SNMPCommunity) { $f.Porta.SNMPCommunity } elseif ($f.Reg) { $f.Reg.'SNMP Community' } else { $null }
                                if ($snmpEnabled -and $snmpIndex) {
                                    $p.SNMP = $snmpIndex
                                    if ($snmpCommunity) { $p.SNMPCommunity = $snmpCommunity }
                                }
                                Add-PrinterPort @p
                                $criou = $true
                            }
                            Set-Printer -Name $imp.Name -PortName $novaPorta
                            $aplicouPorta = $true

                            # Verifica: responde E é o mesmo aparelho (série > UUID > MAC)
                            $motivoRev = $null
                            if ($VerificarDepois -and -not $Forcar) {
                                if (-not (Test-ImpressoraResponde -Nome $nome -PortaTcp $tcp)) { $motivoRev = "'$nome' não respondeu na porta $tcp" }
                                else {
                                    $ipN = if ($porIP) { $nome } else { Get-IPdoNome -Nome $nome }
                                    if ($ipN) {
                                        $x = Get-Identidade -IP $ipN -Coms $comunidadesTodas
                                        if ($x.Serie -and $serieDev -and (Normalize-Serie $x.Serie) -ne $serieDev) { $motivoRev = "'$nome' respondeu, mas é OUTRO aparelho (S/N $($x.Serie); esperado $serieDev)" }
                                        elseif ($x.IppUuid -and $ippUuidDev -and (Normalize-Uuid $x.IppUuid) -ne $ippUuidDev) { $motivoRev = "'$nome' respondeu, mas é OUTRO aparelho (UUID IPP $($x.IppUuid); esperado $ippUuidDev)" }
                                        elseif (-not ($x.Serie -and $serieDev) -and $x.Mac -and $macDev -and (Normalize-Mac $x.Mac) -ne $macDev) { $motivoRev = "'$nome' respondeu, mas é OUTRO aparelho (MAC $($x.Mac); esperado $macDev)" }
                                    }
                                }
                            }
                            if ($motivoRev) {
                                Set-Printer -Name $imp.Name -PortName $portaAntiga
                                if ($criou -and -not (Get-Printer | Where-Object PortName -eq $novaPorta)) {
                                    Remove-PrinterPort -Name $novaPorta -ErrorAction SilentlyContinue
                                }
                                $status = "REVERTIDA: $motivoRev; voltou para '$portaAntiga'$extra"
                                $novaPorta = $null
                            }
                            else {
                                [void]$portasAntigas.Add($portaAntiga)
                                $status = "OK: '$portaAntiga' -> '$novaPorta'$extra"
                                Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $m -Nome $nomeInv -IP $ipAtual `
                                    -Mac $macDev -Serie $serieDev -Modelo $modeloDev -IppUuid $ippUuidDev `
                                    -WsdUuid $wsdUuidDev -UpnpUuid $upnpUuidDev -DeviceGuid $deviceGuidDev `
                                    -HostNames $hostsDev -Fonte $(if ($porIP) { 'porta-ip' } else { 'porta-nome' })
                                $nomesAtribuidos[$nome.ToUpperInvariant()] = @{
                                    IP = $ipAtual; Mac = $macDev; Serie = $serieDev; IppUuid = $ippUuidDev
                                    WsdUuid = $wsdUuidDev; UpnpUuid = $upnpUuidDev; DeviceGuid = $deviceGuidDev; Fila = $imp.Name
                                }
                                [void]$atribuidos.Add([PSCustomObject]@{
                                    Fila = $imp.Name; IP = $ipAtual; Nome = $nomeInv; Serie = $serieDev; Mac = $macDev
                                    IppUuid = $ippUuidDev; WsdUuid = $wsdUuidDev; UpnpUuid = $upnpUuidDev; DeviceGuid = $deviceGuidDev
                                })
                            }
                        }
                        catch {
                            $erroMsg = $_.Exception.Message
                            if ($aplicouPorta) {
                                try { Set-Printer -Name $imp.Name -PortName $portaAntiga; $erroMsg = "$erroMsg; revertida para '$portaAntiga'" }
                                catch { $erroMsg = "$erroMsg; FALHOU a reversão para '$portaAntiga'" }
                                if ($criou -and -not (Get-Printer | Where-Object PortName -eq $novaPorta)) {
                                    try { Remove-PrinterPort -Name $novaPorta -ErrorAction SilentlyContinue } catch { }
                                }
                            }
                            $status = "ERRO: $erroMsg"; $novaPorta = $null
                        }
                    }
                }
            }
        }
    }

    $cor = switch -Wildcard ($status) { 'OK*' {'Green'} 'SIMULA*' {'Cyan'} 'ERRO*' {'Red'} 'REVERT*' {'Yellow'} default {'DarkGray'} }
    Write-Host ("{0,-40} {1}" -f $imp.Name, $status) -ForegroundColor $cor

    $resultado.Add([PSCustomObject]@{
        Impressora = $imp.Name; TipoPorta = $tipo; IPAntigo = $ip; IPAtual = $ipAtual
        NomeDoNo = $(if ($porIP) { $null } else { $nome }); PorIP = $porIP
        Serie = $serieDev; IppUuid = $ippUuidDev
        PortaAntiga = $(if ($portaAntiga) { $portaAntiga } else { $imp.PortName }); PortaNova = $novaPorta
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

