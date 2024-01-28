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

    [Parameter( Mandatory=$true)] [String]$Chave, 
    [Parameter( Mandatory=$false)] [ValidateSet("DDD","DDI")][String]$Categoria
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$username = "@.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop
$SLEEP = 300
$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyyMMdd')+"_phoneSystemCat.csv"

function Log([string]$message){

    $datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    if ( $message -like '*ERRO*') {
        Write-Host "$datetime;$message" -ForegroundColor Red
    }
    else { 
        if ( $message -like '*ATENCAO*') {
            Write-Host "$datetime;$message" -ForegroundColor Yellow
        }
        else {
            Write-Host "$datetime;$message" -ForegroundColor Green
        }
    }
    Add-Content -Path $logFile -Value "$datetime;$message"
}


try {
    Connect-MsolService -Credential $credential -InformationAction SilentlyContinue | Out-Null
}
catch {
    Log "$UPN;ATENCAO: ao conectar MSOLService 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-MsolService -Credential $credential -InformationAction SilentlyContinue | Out-Null
    }
    catch {
        Log "$UPN;ATENCAO: ao conectar MSOLService 2/3"
        Start-Sleep $SLEEP
        try { 
            Connect-MsolService -Credential $credential -InformationAction SilentlyContinue | Out-Null
        }
        catch {
            Log "$UPN;ERRO: ao conectar MSOLService 3/3"
            exit 1
        }
    }
}

try {

    Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
}
catch {
    Log "$UPN;ATENCAO: falha ao conectar MicrosoftTeams 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
    }
    catch {
        Log "$UPN;ATENCAO: falha ao conectar MicrosoftTeams 2/3"
        Start-Sleep $SLEEP
        try {
            Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
        }
        catch {    
            Log "$UPN;ERRO: ao conectar MicrosoftTeams 3/3"
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

  $azu = Get-ADUser $Chave
  $UPN = $azu.UserPrincipalName
  Log "$UPN;SUCESSO: chave $chave encontrada no AD"
}
catch {
    Log "$chave;ERRO: chave não encontrada no AD"
    exit 1
}
    
 

try {

    $csu = Get-CsOnlineUser -Identity $UPN

    Log "$UPN;usuário válido - $($csu.DisplayName) - $($csu.Alias)"
    
    
    if ($Categoria -eq "DDD") {
        $vrota = "Tag:VRP-BR-Nacional"
    }
    else {
        if ($Categoria -eq "DDI") {
            $vrota = "Tag:VRP-BR-Internacional"
        }
        else {
            $vrota = "Tag:VRP-BR-Interno"
        }
    }

    
    #if ($($csu.EnterpriseVoiceEnabled) -and $($csu.HostedVoiceMail) -and $($csu.OnPremLineURI))
    if ($($csu.EnterpriseVoiceEnabled) -and $($csu.LineURI))
    {
        $configurado = $true

        if ($($csu.OnlineVoiceRoutingPolicy) -eq $vrota) {
            Log "$UPN;ATENCAO: já possui esta categoria"
            exit 0
        }
        else {
            
            try {
            
                Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $vrota
                Log "$UPN;configurado: OnlineVoiceRoutingPolicy $vrota"
            }
            catch {
                Log "$UPN;ERRO: OnlineVoiceRoutingPolicy $vrota"
                exit 1
            }
        }
    }
    else
    {
        $configurado = $false
        Log "$UPN;ERRO: usuário não configurado com phoneSystem"
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
