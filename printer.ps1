#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Migra as portas das impressoras instaladas (IP estático ou WSD) para uma porta
    TCP/IP Padrão que usa o NOME DO NÓ da impressora.

.DESCRIPTION
    Por padrão roda em modo SIMULAÇÃO (nada é alterado, só mostra o que seria feito).
    Use -Executar para aplicar as mudanças.

.PARAMETER Executar
    Aplica as alterações. Sem este parâmetro, apenas simula.

.PARAMETER Forcar
    Migra mesmo que o nome do nó não seja resolvido a partir deste computador agora.
    (Não recomendado: a impressora pode parar de imprimir.)

.PARAMETER RemoverPortaAntiga
    Remove a porta antiga (IP ou WSD) se nenhuma outra impressora a estiver usando.

.PARAMETER Impressora
    Filtra por nome (aceita curinga). Padrão: todas.

.EXAMPLE
    .\Migrar-PortasParaNomeDoNo.ps1                      # simulação
    .\Migrar-PortasParaNomeDoNo.ps1 -Executar            # aplica
    .\Migrar-PortasParaNomeDoNo.ps1 -Executar -RemoverPortaAntiga
    .\Migrar-PortasParaNomeDoNo.ps1 -Impressora "*Brother*" -Executar
#>
[CmdletBinding()]
param(
    [switch]$Executar,
    [switch]$Forcar,
    [switch]$RemoverPortaAntiga,
    [string]$Impressora = '*'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- Funções ----

function Test-EhIPv4 {
    param([string]$Texto)
    return [bool]($Texto -match '^\d{1,3}(\.\d{1,3}){3}$')
}

function Get-NomeDoNoPorIP {
    param([string]$IP)

    # 1) DNS reverso
    $dns = (Resolve-DnsName -Name $IP -Type PTR -ErrorAction SilentlyContinue).NameHost
    if ($dns) { return ($dns | Select-Object -First 1).Split('.')[0] }

    # 2) NetBIOS (nome <00> UNIQUE)
    $linha = nbtstat -A $IP 2>$null | Select-String '<00>\s+UNIQUE' | Select-Object -First 1
    if ($linha -and $linha.ToString() -match '^\s*(.+?)\s*<00>\s+UNIQUE') {
        return $Matches[1].Trim()
    }
    return $null
}

function Test-NomeResolve {
    param([string]$Nome)
    try {
        $r = Resolve-DnsName -Name $Nome -NetbiosFallback -ErrorAction Stop |
             Where-Object { $_.IPAddress }
        if ($r) { return $true }
    } catch { }
    try {
        [void][System.Net.Dns]::GetHostAddresses($Nome)
        return $true
    } catch { return $false }
}

# --------------------------------------------------- Portas WSD (registro) ----

$wsdPorts = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\WSD Port\Ports\*" -ErrorAction SilentlyContinue |
    ForEach-Object { $wsdPorts[$_.PSChildName] = $_ }

# --------------------------------------------------------------- Principal ----

$pasta = if ($PSScriptRoot) { $PSScriptRoot } else { $env:TEMP }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logCsv = Join-Path $pasta "MigracaoPortas_$stamp.csv"

if (-not $Executar) {
    Write-Host "`n*** MODO SIMULAÇÃO - nada será alterado. Use -Executar para aplicar. ***`n" -ForegroundColor Yellow
}

$resultado = New-Object System.Collections.Generic.List[object]
$portasAntigas = New-Object System.Collections.Generic.HashSet[string]

foreach ($imp in (Get-Printer -Name $Impressora)) {

    $porta     = Get-PrinterPort -Name $imp.PortName -ErrorAction SilentlyContinue
    $hostAddr  = $porta.PrinterHostAddress
    $ip        = $null
    $nome      = $null
    $tipo      = $null
    $status    = $null
    $novaPorta = $null

    # ---- Identifica tipo de porta, IP e nome do nó ----
    if ($hostAddr) {
        if (Test-EhIPv4 $hostAddr) {
            $tipo = 'TCP/IP (IP estático)'
            $ip   = $hostAddr
            $nome = Get-NomeDoNoPorIP -IP $ip
        } else {
            $tipo = 'TCP/IP (já usa nome)'
            $nome = $hostAddr
        }
    }
    elseif ($imp.PortName -match '^WSD' -or $porta.Description -match 'WSD') {
        $tipo = 'WSD'
        $wsd  = $wsdPorts[$imp.PortName]
        if ($wsd) {
            $ip   = $wsd.IPAddress
            $nome = $wsd.HostName
        }
        if (-not $ip -or -not $nome) {
            $pnp = Get-PnpDeviceProperty -InstanceId "*$($imp.PortName)*" -ErrorAction SilentlyContinue
            if (-not $ip)   { $ip   = ($pnp | Where-Object KeyName -match 'IPAddress').Data }
            if (-not $nome) { $nome = ($pnp | Where-Object KeyName -match 'HostName').Data }
        }
        if (-not $nome -and $ip) { $nome = Get-NomeDoNoPorIP -IP $ip }
    }
    else {
        $tipo = if ($porta.Description) { $porta.Description } else { 'Local/Virtual' }
    }

    # ---- Decide o que fazer ----
    if ($tipo -eq 'TCP/IP (já usa nome)') {
        $status = 'Ignorada: já usa nome do nó'
    }
    elseif ($tipo -notin 'TCP/IP (IP estático)', 'WSD') {
        $status = 'Ignorada: porta local/virtual'
    }
    elseif (-not $nome) {
        $status = 'Ignorada: nome do nó não encontrado'
    }
    elseif (-not (Test-NomeResolve $nome) -and -not $Forcar) {
        $status = "Ignorada: '$nome' não resolve neste PC (use -Forcar para ignorar)"
    }
    else {
        $novaPorta = $nome

        if (-not $Executar) {
            $status = "SIMULAÇÃO: trocaria '$($imp.PortName)' por '$novaPorta'"
        }
        else {
            try {
                # Cria a porta com o nome do nó (se ainda não existir)
                if (-not (Get-PrinterPort -Name $novaPorta -ErrorAction SilentlyContinue)) {
                    $args = @{ Name = $novaPorta }

                    if ($tipo -eq 'TCP/IP (IP estático)' -and $porta.LprQueueName) {
                        # Preserva protocolo LPR
                        $args.LprHostAddress = $nome
                        $args.LprQueueName   = $porta.LprQueueName
                        if ($porta.LprByteCounting) { $args.LprByteCounting = $true }
                    } else {
                        $args.PrinterHostAddress = $nome
                        $pn = if ($porta.PortNumber) { $porta.PortNumber } else { 9100 }
                        $args.PortNumber = $pn
                    }
                    if ($porta.SNMPEnabled) {
                        $args.SNMP = $porta.SNMPIndex
                        if ($porta.SNMPCommunity) { $args.SNMPCommunity = $porta.SNMPCommunity }
                    }
                    Add-PrinterPort @args
                }

                # Aponta a impressora para a nova porta (driver/preferências são mantidos)
                Set-Printer -Name $imp.Name -PortName $novaPorta
                [void]$portasAntigas.Add($imp.PortName)
                $status = "OK: '$($imp.PortName)' -> '$novaPorta'"
            }
            catch {
                $status = "ERRO: $($_.Exception.Message)"
            }
        }
    }

    $cor = switch -Wildcard ($status) {
        'OK*'        { 'Green' }
        'SIMULA*'    { 'Cyan' }
        'ERRO*'      { 'Red' }
        default      { 'DarkGray' }
    }
    Write-Host ("{0,-40} {1}" -f $imp.Name, $status) -ForegroundColor $cor

    $resultado.Add([PSCustomObject]@{
        Impressora = $imp.Name
        Driver     = $imp.DriverName
        TipoPorta  = $tipo
        PortaAntiga= $imp.PortName
        EnderecoIP = $ip
        NomeDoNo   = $nome
        PortaNova  = $novaPorta
        Status     = $status
    })
}

# ---- Remove portas antigas que ficaram sem uso ----
if ($Executar -and $RemoverPortaAntiga) {
    foreach ($p in $portasAntigas) {
        $emUso = Get-Printer | Where-Object PortName -eq $p
        if (-not $emUso) {
            try {
                Remove-PrinterPort -Name $p
                Write-Host "Porta antiga removida: $p" -ForegroundColor Green
            } catch {
                Write-Host "Não foi possível remover a porta '$p': $($_.Exception.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Porta '$p' mantida (ainda em uso por outra impressora)." -ForegroundColor Yellow
        }
    }
}

# ---- Resumo e log (serve também como backup para reverter) ----
Write-Host ""
$resultado | Format-Table Impressora, TipoPorta, EnderecoIP, NomeDoNo, PortaAntiga, PortaNova -AutoSize
$resultado | Export-Csv -Path $logCsv -NoTypeInformation -Encoding UTF8
Write-Host "Log salvo em: $logCsv" -ForegroundColor Cyan
