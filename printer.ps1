# ===================== CONFIGURAÇÃO (edite se quiser) =====================
$Simular            = $false   # $true = só mostra o que faria, sem alterar nada
$RemoverPortaAntiga = $true    # remove a porta antiga se nenhuma impressora a usar mais
$FiltroNome         = '*'      # ex.: '*Brother*' para migrar só algumas impressoras
$IgnorarFax         = $true    # ignora impressoras/drivers de fax (ex.: Brother PC-FAX)
# ==========================================================================

$ErrorActionPreference = 'Stop'

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { Write-Host "Abra o PowerShell como ADMINISTRADOR e cole o script de novo." -ForegroundColor Red; return }

function Test-EhIPv4 { param([string]$Texto); return [bool]($Texto -match '^\d{1,3}(\.\d{1,3}){3}$') }

function Get-NomeDoNoPorIP {
    param([string]$IP)
    $dns = (Resolve-DnsName -Name $IP -Type PTR -ErrorAction SilentlyContinue).NameHost
    if ($dns) { return ($dns | Select-Object -First 1).Split('.')[0] }
    $linha = nbtstat -A $IP 2>$null | Select-String '<00>\s+UNIQUE' | Select-Object -First 1
    if ($linha -and $linha.ToString() -match '^\s*(.+?)\s*<00>\s+UNIQUE') { return $Matches[1].Trim() }
    return $null
}

function Test-NomeResolve {
    param([string]$Nome)
    try { $r = Resolve-DnsName -Name $Nome -NetbiosFallback -ErrorAction Stop | Where-Object { $_.IPAddress }; if ($r) { return $true } } catch { }
    try { [void][System.Net.Dns]::GetHostAddresses($Nome); return $true } catch { return $false }
}

$wsdPorts = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\WSD Port\Ports\*" -ErrorAction SilentlyContinue | ForEach-Object { $wsdPorts[$_.PSChildName] = $_ }

$pasta  = [Environment]::GetFolderPath('Desktop')
if (-not $pasta) { $pasta = $env:TEMP }
if (-not $pasta) { $pasta = (Get-Location).Path }
$logCsv = Join-Path $pasta ("MigracaoPortas_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$resultado = New-Object System.Collections.Generic.List[object]
$portasAntigas = New-Object System.Collections.Generic.HashSet[string]

if ($Simular) { Write-Host "`n*** MODO SIMULAÇÃO: nada será alterado ***`n" -ForegroundColor Yellow }

foreach ($imp in (Get-Printer -Name $FiltroNome)) {

    $porta = Get-PrinterPort -Name $imp.PortName -ErrorAction SilentlyContinue
    $hostAddr = $porta.PrinterHostAddress
    $ip = $null; $nome = $null; $tipo = $null; $status = $null; $novaPorta = $null

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

    if ($IgnorarFax -and ($imp.Name -match 'fax' -or $imp.DriverName -match 'fax')) {
        $status = 'Ignorada: impressora de fax'
    }
    elseif ($tipo -eq 'TCP/IP (já usa nome)') { $status = 'Ignorada: já usa nome do nó' }
    elseif ($tipo -notin 'TCP/IP (IP estático)', 'WSD') { $status = 'Ignorada: porta local/virtual' }
    elseif (-not $nome) { $status = 'Ignorada: nome do nó não encontrado' }
    elseif (-not (Test-NomeResolve $nome)) { $status = "Ignorada: '$nome' não resolve neste PC" }
    else {
        $novaPorta = $nome
        if ($Simular) { $status = "SIMULAÇÃO: trocaria '$($imp.PortName)' por '$novaPorta'" }
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
                $status = "OK: '$($imp.PortName)' -> '$novaPorta'"
            }
            catch { $status = "ERRO: $($_.Exception.Message)" }
        }
    }

    $cor = switch -Wildcard ($status) { 'OK*' {'Green'} 'SIMULA*' {'Cyan'} 'ERRO*' {'Red'} default {'DarkGray'} }
    Write-Host ("{0,-40} {1}" -f $imp.Name, $status) -ForegroundColor $cor

    $resultado.Add([PSCustomObject]@{
        Impressora = $imp.Name; Driver = $imp.DriverName; TipoPorta = $tipo
        PortaAntiga = $imp.PortName; EnderecoIP = $ip; NomeDoNo = $nome
        PortaNova = $novaPorta; Status = $status
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
$resultado | Format-Table Impressora, TipoPorta, EnderecoIP, NomeDoNo, PortaAntiga, PortaNova -AutoSize
$resultado | Export-Csv -Path $logCsv -NoTypeInformation -Encoding UTF8
Write-Host "Log salvo em: $logCsv" -ForegroundColor Cyan
