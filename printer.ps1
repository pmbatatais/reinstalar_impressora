# ===================== CONFIGURAÇÃO (edite se quiser) =====================
$Simular            = $false   # $true = só mostra o que faria, sem alterar nada
$RemoverPortaAntiga = $true    # remove a porta antiga se nenhuma impressora a usar mais
$FiltroNome         = '*'      # ex.: '*Brother*' para migrar só algumas impressoras
$IgnorarFax         = $true    # ignora impressoras/drivers de fax (ex.: Brother PC-FAX)
$Forcar             = $false   # $true = migra mesmo se o nome do nó não resolver neste PC
$BuscarNaRede       = $true    # se o nome não for achado, procura a impressora na rede local

# Mapa manual (opcional): nome da impressora no Windows = nome do nó.
# Use quando houver mais de uma impressora da mesma marca na rede. Exemplo:
# $Manual = @{ 'Tributos 1' = 'BRWBCF4D443678E'; 'Tributos 2' = 'SEC842519AAD2D3' }
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

function Test-NomeResolve {
    param([string]$Nome)
    try { $r = Resolve-DnsName -Name $Nome -NetbiosFallback -ErrorAction Stop | Where-Object { $_.IPAddress }; if ($r) { return $true } } catch { }
    try { [void][System.Net.Dns]::GetHostAddresses($Nome); return $true } catch { return $false }
}

# Marca pela impressora (driver + nome), mesmo que o nome tenha sido trocado
function Get-MarcaImpressora {
    param($Imp)
    $txt = "$($Imp.DriverName) $($Imp.Name)"
    if ($txt -match 'brother') { return [PSCustomObject]@{ Marca = 'Brother'; Prefixo = 'BRW' } }
    if ($txt -match 'samsung') { return [PSCustomObject]@{ Marca = 'Samsung'; Prefixo = 'SEC' } }
    return $null
}

function Get-MarcaPorHttp {
    param([string]$IP)
    try {
        $r = Invoke-WebRequest -Uri "http://$IP" -UseBasicParsing -TimeoutSec 3 -ErrorAction Stop
        if ($r.Content -match 'Brother') { return 'Brother' }
        if ($r.Content -match 'Samsung|SyncThru') { return 'Samsung' }
    } catch { }
    return $null
}

# Procura na sub-rede (/24 de cada placa ativa) dispositivos com a porta de impressão 9100 aberta
function Get-ImpressorasNaRede {
    Write-Host "Procurando impressoras na rede local (pode levar alguns segundos)..." -ForegroundColor Cyan

    $bases = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.IPAddress -notmatch '^(127\.|169\.254\.)' -and
            $_.InterfaceAlias -notmatch 'vEthernet|VirtualBox|VMware|Loopback|Bluetooth|Hyper-V|VPN|WSL' -and
            $_.AddressState -eq 'Preferred'
        } | ForEach-Object { ($_.IPAddress -split '\.')[0..2] -join '.' } | Select-Object -Unique)

    $tentativas = New-Object System.Collections.Generic.List[object]
    foreach ($b in $bases) {
        foreach ($i in 1..254) {
            $ip = "$b.$i"
            $c  = New-Object System.Net.Sockets.TcpClient
            $tentativas.Add([PSCustomObject]@{ IP = $ip; Client = $c; Task = $c.ConnectAsync($ip, 9100) })
        }
    }
    try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]@($tentativas | ForEach-Object { $_.Task }), 3000) } catch { }

    $abertos = @($tentativas | Where-Object { $_.Task.Status -eq 'RanToCompletion' -and $_.Client.Connected } | ForEach-Object { $_.IP })
    foreach ($t in $tentativas) { try { $t.Client.Dispose() } catch { } }

    $achadas = New-Object System.Collections.Generic.List[object]
    foreach ($ip in $abertos) {
        $nome = Get-NomeNetBios -IP $ip
        if (-not $nome) { $nome = Get-NomeDoNoPorIP -IP $ip }
        $marca = $null; $origem = 'nome na rede'

        if ($nome -match '^BRW') { $marca = 'Brother' }
        elseif ($nome -match '^SEC') { $marca = 'Samsung' }
        else {
            # Sem nome útil: descobre a marca pela página web e monta o nome pelo MAC
            $marca = Get-MarcaPorHttp -IP $ip
            $mac = (Get-NetNeighbor -IPAddress $ip -ErrorAction SilentlyContinue | Where-Object LinkLayerAddress | Select-Object -First 1).LinkLayerAddress
            if ($marca -and $mac) {
                $pre = if ($marca -eq 'Brother') { 'BRW' } else { 'SEC' }
                $nome = $pre + ($mac -replace '[-:]', '').ToUpper()
                $origem = 'montado pelo MAC'
            } else { $nome = $null }
        }
        $achadas.Add([PSCustomObject]@{ IP = $ip; Nome = $nome; Marca = $marca; Origem = $origem })
    }

    if ($achadas.Count -gt 0) {
        Write-Host "Dispositivos com porta 9100 encontrados:" -ForegroundColor Cyan
        $achadas | Format-Table IP, Nome, Marca, Origem -AutoSize | Out-String | Write-Host
    } else { Write-Host "Nenhum dispositivo com porta 9100 encontrado." -ForegroundColor Yellow }
    return $achadas
}

# ----------------------------------------------------- Portas WSD (registro) ----

$wsdPorts = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\WSD Port\Ports\*" -ErrorAction SilentlyContinue | ForEach-Object { $wsdPorts[$_.PSChildName] = $_ }

# --------------------------------------------------------------- Principal ----

$resultado = New-Object System.Collections.Generic.List[object]
$portasAntigas = New-Object System.Collections.Generic.HashSet[string]
$descobertas = $null

if ($Simular) { Write-Host "`n*** MODO SIMULAÇÃO: nada será alterado ***`n" -ForegroundColor Yellow }

foreach ($imp in (Get-Printer -Name $FiltroNome)) {

    $porta = Get-PrinterPort -Name $imp.PortName -ErrorAction SilentlyContinue
    $hostAddr = $porta.PrinterHostAddress
    $ip = $null; $nome = $null; $tipo = $null; $status = $null; $novaPorta = $null; $origemNome = $null

    if ($hostAddr) {
        if (Test-EhIPv4 $hostAddr) { $tipo = 'TCP/IP (IP estático)'; $ip = $hostAddr; $nome = Get-NomeDoNoPorIP -IP $ip }
        else { $tipo = 'TCP/IP (já usa nome)'; $nome = $hostAddr }
    }
    elseif ($imp.PortName -match '^WSD' -or $porta.Description -match 'WSD') {
        $tipo = 'WSD'
        $wsd = $wsdPorts[$imp.PortName]
        if ($wsd) { $ip = $wsd.IPAddress; $nome = $wsd.HostName }
        if (-not $ip -or -not $nome) {
            $pnp = Get-PnpDeviceProperty -InstanceId "*$($imp.PortName)*" -ErrorAction SilentlyContinue
            if (-not $ip)   { $ip   = ($pnp | Where-Object KeyName -match 'IPAddress').Data }
            if (-not $nome) { $nome = ($pnp | Where-Object KeyName -match 'HostName').Data }
        }
        if (-not $nome -and $ip) { $nome = Get-NomeDoNoPorIP -IP $ip }
    }
    else { $tipo = if ($porta.Description) { $porta.Description } else { 'Local/Virtual' } }

    $precisaMigrar = $tipo -in 'TCP/IP (IP estático)', 'WSD'
    $ehFax = $IgnorarFax -and ($imp.Name -match 'fax' -or $imp.DriverName -match 'fax')

    if ($ehFax) { $status = 'Ignorada: impressora de fax' }
    elseif ($tipo -eq 'TCP/IP (já usa nome)') { $status = 'Ignorada: já usa nome do nó' }
    elseif (-not $precisaMigrar) { $status = 'Ignorada: porta local/virtual' }
    else {
        $m = Get-MarcaImpressora -Imp $imp

        # 1) Mapa manual tem prioridade
        if ($Manual.ContainsKey($imp.Name)) { $nome = $Manual[$imp.Name]; $origemNome = 'manual' }
        # 2) Descarta nome que não combina com a marca (ex.: IP antigo agora é de outro aparelho)
        elseif ($nome -and $m -and $nome -notmatch "^$($m.Prefixo)") {
            Write-Host "  [$($imp.Name)] nome '$nome' no IP antigo não parece ser da impressora; ignorando." -ForegroundColor DarkYellow
            $nome = $null
        }

        # 3) Procura na rede pela marca
        if (-not $nome -and $BuscarNaRede -and $m) {
            if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
            $cands = @($descobertas | Where-Object { $_.Marca -eq $m.Marca -and $_.Nome })
            if ($cands.Count -eq 1) { $nome = $cands[0].Nome; $ip = $cands[0].IP; $origemNome = "achada na rede ($($cands[0].Origem))" }
            elseif ($cands.Count -gt 1) {
                $status = "Ignorada: $($cands.Count) impressoras $($m.Marca) na rede ($(($cands.Nome) -join ', ')); use o mapa `$Manual"
            }
            else { $status = "Ignorada: nenhuma impressora $($m.Marca) encontrada na rede" }
        }

        if (-not $status) {
            if (-not $nome) { $status = 'Ignorada: nome do nó não encontrado' }
            elseif (-not (Test-NomeResolve $nome) -and -not $Forcar) { $status = "Ignorada: '$nome' não resolve neste PC (use `$Forcar = `$true para ignorar)" }
            else {
                $novaPorta = $nome
                $extra = if ($origemNome) { " [$origemNome]" } else { '' }
                if ($Simular) { $status = "SIMULAÇÃO: trocaria '$($imp.PortName)' por '$novaPorta'$extra" }
                else {
                    try {
                        if (-not (Get-PrinterPort -Name $novaPorta -ErrorAction SilentlyContinue)) {
                            $p = @{ Name = $novaPorta }
                            if ($tipo -eq 'TCP/IP (IP estático)' -and $porta.LprQueueName) {
                                $p.LprHostAddress = $nome
                                $p.LprQueueName = $porta.LprQueueName
                                if ($porta.LprByteCounting) { $p.LprByteCounting = $true }
                            } else {
                                $p.PrinterHostAddress = $nome
                                $p.PortNumber = $(if ($porta.PortNumber) { $porta.PortNumber } else { 9100 })
                            }
                            if ($porta.SNMPEnabled) {
                                $p.SNMP = $porta.SNMPIndex
                                if ($porta.SNMPCommunity) { $p.SNMPCommunity = $porta.SNMPCommunity }
                            }
                            Add-PrinterPort @p
                        }
                        Set-Printer -Name $imp.Name -PortName $novaPorta
                        [void]$portasAntigas.Add($imp.PortName)
                        $status = "OK: '$($imp.PortName)' -> '$novaPorta'$extra"
                    }
                    catch { $status = "ERRO: $($_.Exception.Message)" }
                }
            }
        }
    }

    $cor = switch -Wildcard ($status) { 'OK*' {'Green'} 'SIMULA*' {'Cyan'} 'ERRO*' {'Red'} default {'DarkGray'} }
    Write-Host ("{0,-40} {1}" -f $imp.Name, $status) -ForegroundColor $cor

    $resultado.Add([PSCustomObject]@{
        Impressora = $imp.Name; TipoPorta = $tipo; EnderecoIP = $ip
        NomeDoNo = $nome; PortaAntiga = $imp.PortName; PortaNova = $novaPorta
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
