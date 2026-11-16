# ===================== CONFIGURAÇÃO (edite se quiser) =====================
$Simular            = $false   # $true = só mostra o que faria, sem alterar impressoras
$RemoverPortaAntiga = $true    # remove a porta antiga se nenhuma impressora a usar mais
$FiltroNome         = '*'      # ex.: '*Brother*' para migrar só algumas impressoras
$IgnorarFax         = $true    # ignora impressoras/drivers de fax (ex.: Brother PC-FAX)
$Forcar             = $false   # $true = migra mesmo se o nome do nó não resolver neste PC
$BuscarNaRede       = $true    # se o nome não for achado, procura a impressora na rede local
$VerificarDepois    = $true    # testa a impressora após a troca e REVERTE se não responder
$GravarInventario   = $true    # grava nome do nó / IP / MAC de cada impressora (ProgramData)
                               # (o inventário é gravado mesmo em simulação)

# Mapa manual (opcional): nome da impressora no Windows = nome do nó. Exemplo:
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

function Test-NomeResolve { param([string]$Nome); return [bool](Get-IPdoNome -Nome $Nome) }

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

# MAC do aparelho (sem separadores, maiúsculo), lido da tabela ARP
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

function Get-MarcaPorTexto {
    param([string]$Texto)
    if ($Texto -match 'Brother') { return 'Brother' }
    if ($Texto -match 'Samsung|SyncThru') { return 'Samsung' }
    return $null
}

# $true = modelo do driver aparece na página; $false = não aparece; $null = não dá para saber
function Test-ModeloCombina {
    param([string]$Driver, [string]$WebNorm)
    if (-not $Driver -or -not $WebNorm) { return $null }
    $tokens = @([regex]::Matches($Driver, '[A-Za-z0-9\-]*\d[A-Za-z0-9\-]*') | ForEach-Object { $_.Value } |
        Where-Object { $_.Length -ge 3 -and $_ -notmatch '^(PCL\d*|PS\d*|XPS|v\d+)$' })
    if ($tokens.Count -eq 0) { return $null }
    foreach ($tok in $tokens) {
        $t = ($tok -replace '[^A-Za-z0-9]', '')
        if ($tok -cmatch 'x$') { $re = [regex]::Escape($t.Substring(0, $t.Length - 1).ToUpper()) + '[A-Z0-9]' }   # "M337x" -> M337?
        else { $re = [regex]::Escape($t.ToUpper()) }
        if ($WebNorm -match $re) { return $true }
    }
    return $false
}

# Marca pela impressora (driver + nome), mesmo que a fila tenha sido renomeada
function Get-MarcaImpressora {
    param($Imp)
    $txt = "$($Imp.DriverName) $($Imp.Name)"
    if ($txt -match 'brother') { return [PSCustomObject]@{ Marca = 'Brother'; Prefixo = 'BRW' } }
    if ($txt -match 'samsung') { return [PSCustomObject]@{ Marca = 'Samsung'; Prefixo = 'SEC' } }
    return $null
}

# Devolve o nome da fila que já usa este nó com OUTRO aparelho (IP diferente e MAC não coincidente)
function Test-ConflitoNo {
    param([string]$Nome, [string]$IdIP, [string]$IdMac)
    $r = $nomesAtribuidos[$Nome.ToUpper()]
    if (-not $r) { return $null }
    if ($IdIP -and $r.IP -and $IdIP -ne $r.IP) {
        if ($IdMac -and $r.Mac -and $IdMac -eq $r.Mac) { return $null }
        return $r.Fila
    }
    return $null
}

# Procura na sub-rede (/24 de cada placa ativa) dispositivos com a porta 9100 aberta
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
        $mac  = Get-MacPorIP -IP $ip
        $web  = Get-PaginaWeb -IP $ip
        $webNorm = if ($web) { ($web -replace '[^A-Za-z0-9]', '').ToUpper() } else { $null }
        if ($webNorm -and $webNorm.Length -gt 40000) { $webNorm = $webNorm.Substring(0, 40000) }

        $marca = $null; $origem = 'nome na rede'
        if ($nome -match '^BRW') { $marca = 'Brother' }
        elseif ($nome -match '^SEC') { $marca = 'Samsung' }
        if (-not $marca) { $marca = Get-MarcaPorTexto -Texto $web }

        # Só monta o nome pelo MAC se o aparelho não tiver nenhum nome
        if (-not $nome -and $marca -and $mac) {
            $pre = if ($marca -eq 'Brother') { 'BRW' } else { 'SEC' }
            $nome = $pre + $mac
            $origem = 'montado pelo MAC'
        }
        $achadas.Add([PSCustomObject]@{ IP = $ip; Nome = $nome; Marca = $marca; Mac = $mac; Origem = $origem; Web = $webNorm })
    }

    if ($achadas.Count -gt 0) {
        Write-Host "Dispositivos com porta 9100 encontrados:" -ForegroundColor Cyan
        $achadas | Format-Table IP, Nome, Marca, Mac, Origem -AutoSize | Out-String | Write-Host
    } else { Write-Host "Nenhum dispositivo com porta 9100 encontrado." -ForegroundColor Yellow }
    return $achadas
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
    param([string]$Fila, [string]$Driver, [string]$Marca, [string]$Nome, [string]$IP, [string]$Mac)
    if (-not $Nome -or (-not $IP -and -not $Mac)) { return }
    $velho = $inventario[$Fila]
    $novo = [PSCustomObject]@{
        Fila = $Fila; Driver = $Driver; Marca = $Marca; Nome = $Nome; IP = $IP; Mac = $Mac
        Atualizado = (Get-Date -Format 's')
    }
    if ($velho) {
        foreach ($campo in 'Driver', 'Marca', 'IP', 'Mac') { if (-not $novo.$campo -and $velho.$campo) { $novo.$campo = $velho.$campo } }
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

# ----------------------------------------------------- Portas WSD (registro) ----

$wsdPorts = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\WSD Port\Ports\*" -ErrorAction SilentlyContinue | ForEach-Object { $wsdPorts[$_.PSChildName] = $_ }

# --------------------------------------------------------------- Principal ----

Read-Inventario

# Pré-análise: nós já em uso, nós por marca e quantos APARELHOS distintos há por marca
$conhecidos = @{}
$nomesAtribuidos = @{}
$grupos = New-Object System.Collections.ArrayList
foreach ($x in (Get-Printer)) {
    $pp = Get-PrinterPort -Name $x.PortName -ErrorAction SilentlyContinue
    $mx = Get-MarcaImpressora -Imp $x
    $hostX = $pp.PrinterHostAddress
    $ehNomeX = $hostX -and -not (Test-EhIPv4 $hostX)

    if ($ehNomeX) {
        $nomesAtribuidos[$hostX.ToUpper()] = @{ IP = (Get-IPdoNome -Nome $hostX); Mac = $null; Fila = $x.Name }
        if ($mx) {
            if (-not $conhecidos.ContainsKey($mx.Marca)) { $conhecidos[$mx.Marca] = @() }
            if ($conhecidos[$mx.Marca] -notcontains $hostX) { $conhecidos[$mx.Marca] += $hostX }
        }
    }

    # Agrupa filas que são do mesmo aparelho (compartilham nome, IP ou MAC); ignora fax e portas que não são de rede
    $ehFaxX = $IgnorarFax -and ($x.Name -match 'fax' -or $x.DriverName -match 'fax')
    if ($mx -and -not $ehFaxX) {
        $ids = New-Object System.Collections.Generic.HashSet[string]
        if ($hostX) { [void]$ids.Add($hostX.ToUpper()) }
        elseif ($x.PortName -match '^WSD' -or $pp.Description -match 'WSD') {
            $w = $wsdPorts[$x.PortName]
            if ($w.IPAddress) { [void]$ids.Add(([string]$w.IPAddress).ToUpper()) }
            if ($w.HostName)  { [void]$ids.Add(([string]$w.HostName).ToUpper()) }
            if ($ids.Count -eq 0) { [void]$ids.Add('WSD:' + $x.PortName.ToUpper()) }
        }
        if ($ids.Count -gt 0) {
            $e = $inventario[$x.Name]
            if ($e) { foreach ($v in $e.Nome, $e.IP, $e.Mac) { if ($v) { [void]$ids.Add(([string]$v).ToUpper()) } } }
            $alvo = $null
            foreach ($g in @($grupos)) {
                if ($g.Marca -ne $mx.Marca) { continue }
                if (@($ids | Where-Object { $g.Ids.Contains($_) }).Count -gt 0) {
                    if ($null -eq $alvo) { $alvo = $g; $alvo.Ids.UnionWith($ids) }
                    else { $alvo.Ids.UnionWith($g.Ids); $grupos.Remove($g) }
                }
            }
            if ($null -eq $alvo) { [void]$grupos.Add(@{ Marca = $mx.Marca; Ids = $ids }) }
        }
    }
}
$qtdDisp = @{}
foreach ($g in $grupos) { $qtdDisp[$g.Marca] = 1 + [int]$qtdDisp[$g.Marca] }
foreach ($k in @($qtdDisp.Keys)) { Write-Host "Aparelhos $k identificados nas filas: $($qtdDisp[$k])" -ForegroundColor DarkGray }

$resultado = New-Object System.Collections.Generic.List[object]
$portasAntigas = New-Object System.Collections.Generic.HashSet[string]
$descobertas = $null

if ($Simular) { Write-Host "`n*** MODO SIMULAÇÃO: nenhuma impressora será alterada ***`n" -ForegroundColor Yellow }

foreach ($imp in (Get-Printer -Name $FiltroNome)) {

    $porta = Get-PrinterPort -Name $imp.PortName -ErrorAction SilentlyContinue
    $hostAddr = $porta.PrinterHostAddress
    $ip = $null; $nome = $null; $tipo = $null; $status = $null; $novaPorta = $null; $origemNome = $null
    $ipAtual = $null; $macDev = $null

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
    $m = Get-MarcaImpressora -Imp $imp

    if ($ehFax) { $status = 'Ignorada: impressora de fax' }
    elseif ($tipo -eq 'TCP/IP (já usa nome)') {
        $status = 'Ignorada: já usa nome do nó'
        # só registra no inventário (para reconhecer o aparelho pelo MAC numa futura troca de rede)
        $ipN = Get-IPdoNome -Nome $nome
        if ($ipN -and (Test-Porta -Alvo $ipN -Porta 9100)) {
            $macN2 = Get-MacPorIP -IP $ipN
            Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $m.Marca -Nome $nome -IP $ipN -Mac $macN2
            if ($nomesAtribuidos.ContainsKey($nome.ToUpper())) { $nomesAtribuidos[$nome.ToUpper()].Mac = $macN2 }
        }
    }
    elseif (-not $precisaMigrar) { $status = 'Ignorada: porta local/virtual' }
    else {
        $inv = $inventario[$imp.Name]
        $bloqueioMotivo = $null
        # Caminhos de reserva (4 e 5) só quando há UM aparelho dessa marca; com vários, só identificação forte
        $reservaLiberada = ($m -and [int]$qtdDisp[$m.Marca] -le 1)

        # 1) Mapa manual tem prioridade
        if ($Manual.ContainsKey($imp.Name)) { $nome = $Manual[$imp.Name]; $origemNome = 'manual' }

        # 2) Nome achado no IP antigo: confere se o aparelho é mesmo a impressora
        elseif ($nome -and $ip) {
            $resp  = Test-Porta -Alvo $ip -Porta 9100
            $macIp = if ($resp) { Get-MacPorIP -IP $ip } else { $null }
            if ($inv -and $inv.Mac -and $macIp -and ($inv.Mac -ne $macIp)) {
                Write-Host "  [$($imp.Name)] o IP antigo ($ip) agora é de OUTRO aparelho (MAC diferente do inventário); ignorando." -ForegroundColor DarkYellow
                $nome = $null
            }
            elseif (-not $resp -and -not ($m -and $nome -match "^($($m.Prefixo)|SEC|BRW)")) {
                Write-Host "  [$($imp.Name)] '$nome' no IP antigo ($ip) não parece ser uma impressora; ignorando." -ForegroundColor DarkYellow
                $nome = $null
            }
            else {
                # O nome precisa apontar para o MESMO aparelho que está no IP da fila
                $ipNome = if ($resp) { Get-IPdoNome -Nome $nome } else { $null }
                if ($resp -and $ipNome -and $ipNome -ne $ip) {
                    [void](Test-Porta -Alvo $ipNome -Porta 9100)
                    $macNome = Get-MacPorIP -IP $ipNome
                    if (-not ($macIp -and $macNome -and $macIp -eq $macNome)) {
                        $bloqueioMotivo = "conflito de identidade: o IP da fila ($ip) e o nome '$nome' (IP $ipNome) não são o mesmo aparelho; use `$Manual"
                        $nome = $null
                    }
                }
                if ($nome) { $ipAtual = $ip; $macDev = $macIp }
            }
        }

        # 3) Inventário: acha o aparelho na rede pelo MAC (ou pelo nome do nó) gravado antes da troca
        if (-not $nome -and $BuscarNaRede -and $inv -and ($inv.Mac -or $inv.Nome)) {
            if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
            $d = $null
            if ($inv.Mac) { $d = $descobertas | Where-Object { $_.Mac -eq $inv.Mac } | Select-Object -First 1 }
            if (-not $d -and $inv.Nome) { $d = $descobertas | Where-Object { $_.Nome -and $_.Nome -eq $inv.Nome } | Select-Object -First 1 }
            if ($d) {
                $nome = if ($d.Nome -and $d.Origem -eq 'nome na rede') { $d.Nome } elseif ($inv.Nome) { $inv.Nome } else { $d.Nome }
                $ipAtual = $d.IP; $macDev = $d.Mac
                $origemNome = "inventário: MAC $($d.Mac), IP atual $($d.IP)"
            }
        }

        # 4) Outra fila da mesma marca já usa um nó: se for um só, é o mesmo aparelho
        if (-not $nome -and $reservaLiberada -and -not $bloqueioMotivo -and $conhecidos[$m.Marca] -and @($conhecidos[$m.Marca]).Count -eq 1) {
            $nome = @($conhecidos[$m.Marca])[0]; $origemNome = 'mesmo nó de outra fila'
        }

        # 5) Procura na rede por marca (e modelo, se houver mais de um candidato)
        if (-not $nome -and $BuscarNaRede -and $reservaLiberada -and -not $bloqueioMotivo) {
            if ($null -eq $descobertas) { $descobertas = @(Get-ImpressorasNaRede) }
            $cands = @($descobertas | Where-Object { $_.Marca -eq $m.Marca -and $_.Nome })
            if ($cands.Count -gt 1) {
                $porModelo = @($cands | Where-Object { (Test-ModeloCombina -Driver $imp.DriverName -WebNorm $_.Web) -eq $true })
                if ($porModelo.Count -ge 1) { $cands = $porModelo }
            }
            if ($cands.Count -gt 1 -and $conhecidos[$m.Marca]) {
                $pref = @($cands | Where-Object { $conhecidos[$m.Marca] -contains $_.Nome })
                if ($pref.Count -eq 1) { $cands = $pref }
            }
            if ($cands.Count -eq 1) {
                $nome = $cands[0].Nome; $ipAtual = $cands[0].IP; $macDev = $cands[0].Mac
                $mod = Test-ModeloCombina -Driver $imp.DriverName -WebNorm $cands[0].Web
                $obs = if ($mod -eq $true) { 'modelo confirmado' } elseif ($mod -eq $false) { 'modelo NÃO confirmado' } else { 'modelo não verificável' }
                $origemNome = "achada na rede, $($cands[0].Origem), $obs"
            }
            elseif ($cands.Count -gt 1) {
                $status = "Ignorada: $($cands.Count) impressoras $($m.Marca) na rede ($(($cands.Nome) -join ', ')); use o mapa `$Manual"
            }
            else { $status = "Ignorada: nenhuma impressora $($m.Marca) encontrada na rede" }
        }

        if (-not $nome -and -not $status) {
            if ($bloqueioMotivo) { $status = "Ignorada: $bloqueioMotivo" }
            elseif ($m -and -not $reservaLiberada) {
                $status = "Ignorada: há $([int]$qtdDisp[$m.Marca]) aparelhos $($m.Marca) instalados e não consegui identificar este com segurança (rode antes da troca de rede para gravar o inventário, ou use `$Manual)"
            }
        }

        # Um mesmo nó não pode ir para duas filas de aparelhos diferentes
        $outraFila = $null
        if ($nome -and $origemNome -ne 'manual') { $outraFila = Test-ConflitoNo -Nome $nome -IdIP $ipAtual -IdMac $macDev }

        if (-not $status) {
            if (-not $nome) { $status = 'Ignorada: nome do nó não encontrado' }
            elseif ($outraFila) { $status = "Ignorada: o nó '$nome' já é usado pela fila '$outraFila' com outro IP/aparelho; use `$Manual se forem o mesmo" }
            elseif (-not (Test-NomeResolve $nome) -and -not $Forcar) { $status = "Ignorada: '$nome' não resolve neste PC (use `$Forcar = `$true para ignorar)" }
            else {
                $novaPorta = $nome
                if (-not $ipAtual) { $ipAtual = Get-IPdoNome -Nome $nome }
                if (-not $macDev -and $ipAtual -and (Test-Porta -Alvo $ipAtual -Porta 9100)) { $macDev = Get-MacPorIP -IP $ipAtual }
                $extra = if ($origemNome) { " [$origemNome]" } else { '' }

                if ($Simular) {
                    $status = "SIMULAÇÃO: trocaria '$($imp.PortName)' por '$novaPorta'$extra"
                    Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $m.Marca -Nome $nome -IP $ipAtual -Mac $macDev
                    $nomesAtribuidos[$nome.ToUpper()] = @{ IP = $ipAtual; Mac = $macDev; Fila = $imp.Name }
                }
                else {
                    try {
                        $usaLpr = ($tipo -eq 'TCP/IP (IP estático)' -and $porta.LprQueueName)
                        $tcp = if ($usaLpr) { 515 } elseif ($porta.PortNumber) { [int]$porta.PortNumber } else { 9100 }
                        $criou = $false

                        if (-not (Get-PrinterPort -Name $novaPorta -ErrorAction SilentlyContinue)) {
                            $p = @{ Name = $novaPorta }
                            if ($usaLpr) {
                                $p.LprHostAddress = $nome
                                $p.LprQueueName = $porta.LprQueueName
                                if ($porta.LprByteCounting) { $p.LprByteCounting = $true }
                            } else {
                                $p.PrinterHostAddress = $nome
                                $p.PortNumber = $tcp
                            }
                            if ($porta.SNMPEnabled) {
                                $p.SNMP = $porta.SNMPIndex
                                if ($porta.SNMPCommunity) { $p.SNMPCommunity = $porta.SNMPCommunity }
                            }
                            Add-PrinterPort @p
                            $criou = $true
                        }
                        Set-Printer -Name $imp.Name -PortName $novaPorta

                        # Verifica a nova porta: a impressora precisa responder E ser o mesmo aparelho (MAC); senão volta para a antiga
                        $motivoRev = $null
                        if ($VerificarDepois -and -not $Forcar) {
                            if (-not (Test-ImpressoraResponde -Nome $novaPorta -PortaTcp $tcp)) { $motivoRev = "'$novaPorta' não respondeu na porta $tcp" }
                            elseif ($macDev) {
                                $ipN = Get-IPdoNome -Nome $novaPorta
                                if ($ipN) {
                                    [void](Test-Porta -Alvo $ipN -Porta $tcp)
                                    $macN = Get-MacPorIP -IP $ipN
                                    if ($macN -and $macN -ne $macDev) { $motivoRev = "'$novaPorta' respondeu, mas é OUTRO aparelho (MAC $macN; esperado $macDev)" }
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
                            Update-Inventario -Fila $imp.Name -Driver $imp.DriverName -Marca $m.Marca -Nome $nome -IP $ipAtual -Mac $macDev
                            $nomesAtribuidos[$nome.ToUpper()] = @{ IP = $ipAtual; Mac = $macDev; Fila = $imp.Name }
                        }
                    }
                    catch { $status = "ERRO: $($_.Exception.Message)"; $novaPorta = $null }
                }
            }
        }
    }

    $cor = switch -Wildcard ($status) { 'OK*' {'Green'} 'SIMULA*' {'Cyan'} 'ERRO*' {'Red'} 'REVERT*' {'Yellow'} default {'DarkGray'} }
    Write-Host ("{0,-40} {1}" -f $imp.Name, $status) -ForegroundColor $cor

    $resultado.Add([PSCustomObject]@{
        Impressora = $imp.Name; TipoPorta = $tipo; IPAntigo = $ip; IPAtual = $ipAtual
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
Write-Inventario
