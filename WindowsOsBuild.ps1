<# Inventaire builds Windows -> XLSX + OUPath + Webhook Discord (multi-OU) #>

param(
    [string[]]$SearchBases = @(
        "OU=Domain Controllers,DC=,DC=",
        "OU=,OU=,DC=,DC="   # ajoute d'autres OU ici si besoin
    ),
    [string]$OutFolder   = "C:\Temp",
    [string]$WebhookUrl  = "",
    [int]$PingTimeoutMs  = 1500,
    [int]$CimTimeoutSec  = 5,
    [switch]$SkipPing
)

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel     -ErrorAction Stop

if (-not (Test-Path $OutFolder)) { New-Item -Path $OutFolder -ItemType Directory | Out-Null }

$DateRapport = Get-Date -Format "yyyy-MM-dd_HH-mm"
$TitreRapport = "Inventaire Windows - $DateRapport"
$OutXlsx = Join-Path $OutFolder "Inventaire_Builds_$DateRapport.xlsx"

function Test-HostReachable {
    param([string]$Name,[int]$TimeoutMs=1500)
    try { ($r= (New-Object System.Net.NetworkInformation.Ping).Send($Name,$TimeoutMs)).Status -eq 'Success' } catch { $false }
}

# --- Collecte multi-OU avec compteur et gestion d'erreurs ---
$all = @()
foreach ($sb in $SearchBases) {
    try {
        $items = Get-ADComputer -SearchBase $sb -SearchScope Subtree `
            -LDAPFilter "(objectClass=computer)" `
            -Properties DNSHostName,IPv4Address,OperatingSystem,Enabled,DistinguishedName
        Write-Host ("OU: {0} -> {1} objets" -f $sb, ($items | Measure-Object).Count)
        $all += $items
    } catch {
        Write-Warning "OU inaccessible: $sb -> $($_.Exception.Message)"
    }
}
# Dédup par nom d’objet AD
$computers = $all | Where-Object { $_.Enabled -eq $true } | Sort-Object -Property Name -Unique

$results = foreach ($c in $computers) {
    $compHost = if ($c.DNSHostName) { $c.DNSHostName } else { $c.Name }
    $compIP   = $c.IPv4Address
    $online = if ($SkipPing) { $true } else { Test-HostReachable -Name $compHost -TimeoutMs $PingTimeoutMs }

    if (-not $online) {
        [pscustomobject]@{
            ComputerName   = $c.Name
            DNSHostName    = $compHost
            IPv4           = $compIP
            OSName         = $c.OperatingSystem
            Version        = $null
            Build          = $null
            UBR            = $null
            DisplayVersion = $null
            OUPath         = $c.DistinguishedName
            Status         = "Offline/Timeout"
        }
        continue
    }

    try {
        $sessOpt = New-CimSessionOption -Protocol DCOM
        $sess    = New-CimSession -ComputerName $compHost -SessionOption $sessOpt -OperationTimeoutSec $CimTimeoutSec -ErrorAction Stop
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $sess -ErrorAction Stop

        $HKLM = 0x80000002
        $cvKey = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $ubrRes  = Invoke-CimMethod -CimSession $sess -ClassName StdRegProv -MethodName GetDWORDValue `
                     -Arguments @{ hDefKey = $HKLM; sSubKeyName = $cvKey; sValueName = 'UBR' } -ErrorAction SilentlyContinue
        $dispRes = Invoke-CimMethod -CimSession $sess -ClassName StdRegProv -MethodName GetStringValue `
                     -Arguments @{ hDefKey = $HKLM; sSubKeyName = $cvKey; sValueName = 'DisplayVersion' } -ErrorAction SilentlyContinue

        if (-not $compIP) {
            try {
                $compIP = ([System.Net.Dns]::GetHostAddresses($compHost) |
                           Where-Object {$_.AddressFamily -eq 'InterNetwork'} |
                           Select-Object -First 1).IPAddressToString
            } catch {}
        }

        [pscustomobject]@{
            ComputerName   = $c.Name
            DNSHostName    = $compHost
            IPv4           = $compIP
            OSName         = $os.Caption
            Version        = $os.Version
            Build          = $os.BuildNumber
            UBR            = ($ubrRes.uValue -as [int])
            DisplayVersion = $dispRes.sValue
            OUPath         = $c.DistinguishedName
            Status         = "OK"
        }

        Remove-CimSession -CimSession $sess -ErrorAction SilentlyContinue
    }
    catch {
        [pscustomobject]@{
            ComputerName   = $c.Name
            DNSHostName    = $compHost
            IPv4           = $compIP
            OSName         = $c.OperatingSystem
            Version        = $null
            Build          = $null
            UBR            = $null
            DisplayVersion = $null
            OUPath         = $c.DistinguishedName
            Status         = "Erreur: $($_.Exception.Message)"
        }
    }
}

# --- Excel formaté ---
$sheet = "Inventaire"
$results | Sort-Object DNSHostName | Export-Excel -Path $OutXlsx -WorksheetName $sheet `
    -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -TableName "InventaireWindows" -TableStyle Medium2 `
    -Title $TitreRapport -TitleBold -TitleSize 14
Write-Host "Rapport XLSX généré : $OutXlsx"

# --- Upload Discord (multipart binaire) ---
try {
    $boundary = [System.Guid]::NewGuid().ToString()
    $lf = "`r`n"
    $fileBytes = [System.IO.File]::ReadAllBytes($OutXlsx)

    $ms = New-Object System.IO.MemoryStream
    $sw = New-Object System.IO.StreamWriter($ms, [System.Text.ASCIIEncoding]::new())

    $sw.Write("--$boundary$lf")
    $sw.Write("Content-Disposition: form-data; name=`"file1`"; filename=`"$(Split-Path $OutXlsx -Leaf)`"$lf")
    $sw.Write("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet$lf$lf")
    $sw.Flush()
    $ms.Write($fileBytes, 0, $fileBytes.Length)
    $sw.Write("$lf")

    $sw.Write("--$boundary$lf")
    $sw.Write("Content-Disposition: form-data; name=`"payload_json`"$lf$lf")
    $sw.Write("{""content"":""Rapport inventaire Windows OS Build du $DateRapport""}$lf")
    $sw.Write("--$boundary--$lf")
    $sw.Flush()

    Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType "multipart/form-data; boundary=$boundary" -Body $ms.ToArray()
    Write-Host "Rapport envoyé à Discord."
}
catch {
    Write-Host "Erreur lors de l'envoi du Webhook Discord : $($_.Exception.Message)"
}

# --- Récap ---
Write-Host ("Total machines uniques: {0}" -f (($computers | Measure-Object).Count))

