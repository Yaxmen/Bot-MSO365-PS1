<#

nome:       phoneSystemCat.ps1


uso :       ./phoneSystemCat.ps1 -Acao add|del -Chave user@pertrobras.com.br [-Categoria  DDD|DDI] [-Ramal 2121123456]


08-09-21 adicionado try-catch nas conexões e disconnect-MicrosoftTeams no final
15-09-21 alterado UPN pela chave
14-10-21 alterados nomes dos planos TenantDialPlan 
    '21' = 'Tag:Rio de Janeiro' ---> Tag:DP-BR-21
    '22' = 'Tag:Macae'          ---> Tag:DP-BR-22
    '13' = 'Tag:Santos'         ---> Tag:DP-BR-13
    '27' = 'Tag:Vitoria'        ---> Tag:DP-BR-27
     Categorias de DDD e DDI para Tag:VRP-BR-Nacional e Tag:VRP-BR-Internacional

Get-NetFirewallRule 

#>

Param (
    [Parameter(Mandatory=$true)]
    [String]$Chave,

    [Parameter(Mandatory=$true)]
    [ValidateSet("DDD","DDI")]
    [String]$Categoria
)

Import-Module Microsoft.Graph.Users

#Setando constantes
$ClientId = "b88e924b-530d-497c-8b16-b54456064e5f"
$TenantId = "5b6f6241-9a57-4be4-8e50-1dfa72e79a57"
$key = (1..16)
$Secret = Get-Content "D:\Password\GraphAppPassword.txt" | ConvertTo-SecureString -Key $key -ErrorAction Stop

$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $Secret

#$ClientId = "60c5c5b3-84dd-4519-8ad7-58a34f8d39b4"
#$TenantId = "6af8f826-d4c2-47de-9f6d-c04908aa4e88"
#$Key = Get-Content "D:\Util\Password\KeyFile.key"
#$Secret = Get-Content "D:\Util\Password\Password.txt" | ConvertTo-SecureString -Key $Key -ErrorAction Stop
#$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $Secret

$SKU = "MCOEV"
#$GUID = "f55f89fa-f356-4aa3-a295-180a08fa9957"
$GUID = "0018869d-60f6-4a54-ba71-2813bdcd8d8d"
$SLEEP = 60
$logFile = "D:\Util\AlteraçãoDeCategoria\logs\"+(Get-Date).ToString('yyyy')+"_phoneSystemCat_log.csv"

#$username = "@.com.br"
#$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
#$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
#$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop
#$SLEEP = 300
#$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyyMMdd')+"_phoneSystemCat.csv"

# Definindo credenciais de acesso a tenant
$username = "@.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#$username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
#$PlainPassword="Ror66406"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword

#Função de Log
function Log([string]$message){

    $datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    if ($message -like '*ERRO*') {
        Write-Host "$datetime;$message" -ForegroundColor Red
    }
    else { 
        if ($message -like '*ATENCAO*') {
            Write-Host "$datetime;$message" -ForegroundColor Yellow
        }
        else {
            Write-Host "$datetime;$message" -ForegroundColor Green
        }
    }
    Add-Content -Path $logFile -Value "$datetime;$message"
}

#Conectando no MicrosoftGraph
try {
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential | Out-Null
    Log "SUCESSO:  Microsoft Graph 1/3"
}
catch {
    Log "$UPN;ATENCAO: ao conectar Microsoft Graph 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential | Out-Null
        Log "SUCESSO:  Microsoft Graph 2/3"
    }
    catch {
        Log "$UPN;ATENCAO: ao conectar Microsoft Graph 2/3"
        Start-Sleep $SLEEP
        try { 
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential | Out-Null
            Log "SUCESSO:  Microsoft Graph 3/3"
        }
        catch {
            Log "$UPN;ERRO: ao conectar Microsoft Graph 3/3"
            exit 1
        }
    }
}

#Conectando no MicrosoftTeams
try {

    Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
    Log "SUCESSO: ao conectar MicrosoftTeams 1/3"
}
catch {
    Log "ATENCAO: ao conectar MicrosoftTeams 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
        Log "SUCESSO: ao conectar MicrosoftTeams 2/3"
    }
    catch {
        Log "ATENCAO: ao conectar MicrosoftTeams 2/3"
        Start-Sleep $SLEEP
        try {
            Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
            Log "SUCESSO: ao conectar MicrosoftTeams 2/3"
        }
        catch {    
            Log "*ERRO*: ao conectar MicrosoftTeams 3/3"
            exit 1
        }

    }
}

$cod_hash = @{
    '21' = 'Tag:Rio de Janeiro'
    '22' = 'Tag:Macae'
    '13' = 'Tag:Santos'
    '27' = 'Tag:Vitoria'
}

$cod = ""
$plano = ""

try {
    $ADUser = Get-ADUser $Chave
    $azu = Get-MgUser -UserId $ADUser.UserPrincipalName
    $UPN = $azu.UserPrincipalName
    Log "$UPN;SUCESSO: chave $chave encontrada no Azure AD"
}
catch {
    Log "$chave;ERRO: chave não encontrada no Azure AD"
    exit 1
}

try {
    $csu = Get-CsOnlineUser -Identity $azu.Id

    Log "$UPN;usuário válido - $($csu.DisplayName) - $($csu.Alias)"
    
    if ($Categoria -eq "DDD") {
        $vrota = "Tag:VRP-BR-Nacional"
    }
    elseif ($Categoria -eq "DDI") {
        $vrota = "Tag:VRP-BR-Internacional"
    }
    else {
        $vrota = "Tag:VRP-BR-Interno"
    }
    
    if ($csu.EnterpriseVoiceEnabled -and $csu.LineURI) {
        $configurado = $true

        if ($csu.VoiceRoutingPolicy -eq $vrota) {
            Log "$UPN;ATENCAO: já possui esta categoria"
            exit 0
        }
        else {
            try {
                Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $vrota
                Log "$UPN;configurado: VoiceRoutingPolicy $vrota"
            }
            catch {
                Log "$UPN;ERRO: VoiceRoutingPolicy $vrota"
                exit 1
            }
        }
    }
    else {
        $configurado = $false
        Log "$UPN;ERRO: usuário não configurado com Phone System"
        exit 1
    }
}
catch {
    Log "$UPN;ERRO: ao recuperar usuário"
    exit 1
}

try {
    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
}
catch {
    Start-Sleep 10
    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
}
