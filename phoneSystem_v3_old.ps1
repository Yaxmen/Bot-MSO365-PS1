﻿<#

nome:       phonesystem_v3.ps1


uso :       ./phonesystem_config.ps1 -Acao add|del -Chave user@pertrobras.com.br [-Categoria  DDD|DDI] [-Ramal 2121123456]


08-09-21 adicionado try-catch nas conexões e disconnect-MicrosoftTeams no final
15-09-21 alterado UPN pela chave
14-10-21 alterados nomes dos planos TenantDialPlan 
    '21' = 'Tag:Rio de Janeiro' ---> Tag:DP-BR-21
    '22' = 'Tag:Macae'          ---> Tag:DP-BR-22
    '13' = 'Tag:Santos'         ---> Tag:DP-BR-13
    '27' = 'Tag:Vitoria'        ---> Tag:DP-BR-27
     Categorias de DDD e DDI para Tag:VRP-BR-Nacional e Tag:VRP-BR-Internacional
27-10-21 crida função para verificar se ramal já esta previamente sendo utiliado

Get-NetFirewallRule 

#>

Param (

    [Parameter( Mandatory=$true)] [ValidateSet("add","del")] [String]$Acao, 
    [Parameter( Mandatory=$true)] [String]$Chave, 
    [Parameter( Mandatory=$false)] [ValidateSet("DDD","DDI")][String]$Categoria,
    [Parameter( Mandatory=$false)][String]$Ramal
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$username = "samig01@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop
$SLEEP = 300
$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyy')+"_phoneSystem_log.csv"
$SKU = "petrobrasbr:MCOEV"

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
    Log "SUCESSO: ao conectar MSOLService 1/3"
}
catch {
    Log "ATENCAO: ao conectar MSOLService 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-MsolService -Credential $credential -InformationAction SilentlyContinue | Out-Null
        Log "SUCESSO: ao conectar MSOLService 2/3"
    }
    catch {
        Log "ATENCAO: ao conectar MSOLService 2/3"
        Start-Sleep $SLEEP
        try { 
            Connect-MsolService -Credential $credential -InformationAction SilentlyContinue | Out-Null
            Log "SUCESSO: ao conectar MSOLService 3/3"
        }
        catch {
            Log "ERRO: ao conectar MSOLService 3/3"
            exit 1
        }
    }
}

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
            Log "ERRO: ao conectar MicrosoftTeams 3/3"
            exit 1
        }

    }
}


### consulta, no AD interno qual $UPN da $Chave informada. Caso não encontre encerra o script.

try {

  $azu = Get-ADUser $Chave
  $UPN = $azu.UserPrincipalName
  Log "$UPN;SUCESSO: chave $chave encontrada no AD"
}
catch {
    Log "$chave;ERRO: chave não encontrada no AD"
    exit 1
}
    
### verifica as configurações do usuario $UPN no diretorio do Teams.

try {
    $OnlineUser = Get-CsOnlineUser -Filter "UserPrincipalName -eq '$UPN'"
}
catch {
    Log "$UPN;ERRO: ao executar Get-CsOnlineUser -Filter"
    exit 1
}

### encerra o script, caso dados do usuario $UPN retorne nulo no diretório do Teams

if ($OnlineUser -eq $null) {
    Log "$UPN;ERRO: usuário inválido"
    exit 1
}

### configurando phoneSystem (acao = add)

if ($Acao -eq "add") {

    $cod_hash = @{

    '21' = 'Tag:DP-BR-21'
    '22' = 'Tag:DP-BR-22'
    '13' = 'Tag:DP-BR-13'
    '27' = 'Tag:DP-BR-27'

    }

    $cod = $Ramal -match '\+55(\d\d)\d\d\d\d\d\d\d\d'

    ### encerra processamento se ramal foi enviado no formato errado

    if (-not($cod)) {
        Log "$UPN;ERRO: $Ramal com formato inválido"
        exit 1
    }

    ### encerra processamento se codigo DDD não constar na tabela $cod_hash

    $cod = $matches[1]
    $plano = $cod_hash[$cod]

    if (-not($cod_hash.ContainsKey($cod))) {
        Log "$UPN;ERRO: não encontrado plano para código $cod"
        exit 1
    }

    $linha = "tel:"+"$Ramal"

    ### verifica se usuario já esta configurado com ramal solicitado

    if ($($OnlineUser.LineURI) -eq $linha) {

        Log "$UPN;ATENCAO: já configurado Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"
        exit
    }

    ### verifica se linha solicitada encontra-se disponível

    try {
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq '$linha'"
    }
    catch {
        Log "$UPN;ERRO: ao executar Get-CsOnlineUser -Filter Ramal"
        exit 1
    }

    ### encerra script caso ramal solicitado já esteja em uso

    if ($OnlineRamal -ne $null) {
        
        $status_ramal =  "$($OnlineRamal.Alias)"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"   
        #Log "$UPN;ATENCAO: previamente configurado o numero $Ramal para $($OnlineRamal.UserPrincipalName)"
        Log "$UPN;ATENCAO: previamente configurado o numero $Ramal para chave $status_ramal"  
        exit 1
                
    }

    if ($($OnlineUser.LineURI) -ne $null -and $($OnlineUser.LineURI) -ne "") {
    
        Log "$UPN;ATENCAO: já configurado Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"
        exit 
    }

    ### verifica se usuario possui licença e aplica caso não tenha

    $msu = Get-MsolUser -UserPrincipalName $UPN

    if ($msu.Licenses.AccountSkuId -notcontains $SKU) {

        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "$SKU"
        Log "$UPN;configurado $SKU - wait $SLEEP seg"
        Start-Sleep $SLEEP
        
    }
    else {
        Log "$UPN;ATENCAO: previamente configurado $SKU"
    }
    
    
    try {
        Set-CsUser -Identity $UPN -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI "$linha"
        #Set-CsUser -Identity $UPN -EnterpriseVoiceEnabled $true -LineURI "$linha"
        Log "$UPN;configurado: EnterpriseVoice True - HostedVoiceMail True - LineURI $linha"
    }
    catch {
        Log "$UPN;ERRO: EnterpriseVoice True - HostedVoiceMail True - LineURI $linha"
        exit 1
    }
    
    Start-Sleep 10

    ### define categoria
                        
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
    
    try {
        Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $vrota 2>D:\Util\PhoneSystem\logs\err_CsOnlineVoiceRoutingPolicy.txt
    }
    catch {
        Log "$UPN;ERRO: OnlineVoiceRoutingPolicy $vrota"
        exit 1
    }

    Log "$UPN;configurado: OnlineVoiceRoutingPolicy $vrota"
    Start-Sleep 10 
    
    try {   
        Grant-CsTenantDialPlan -Identity $UPN -PolicyName $plano 2>D:\Util\PhoneSystem\logs\err_CsTenantDialPlan.txt
    }
    catch {
        Log "$UPN;ERRO: TenantDialPlan $plano"
        exit 1
    }
            Log "$UPN;configurado: TenantDialPlan $plano"
            
}
else {

    if ($Acao -eq "del") {

        if ($OnlineUser.LineURI -eq $null)
        {
            Log "$UPN;ATENCAO: usuário não está configurado"
            exit 1
        }

        Log "$UPN;antes de remover - Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"

        try {

            Grant-CsTenantDialPlan -Identity $UPN -PolicyName $null
            Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $null
            Set-CsUser -Identity $UPN -EnterpriseVoiceEnabled $false -HostedVoiceMail $false -OnPremLineURI $null
            #Set-CsUser -Identity $UPN -EnterpriseVoiceEnabled $false -LineURI $null
        }
        catch {

            Log "$UPN;ERRO: uao desconfigurar phoneSystem"
            exit 1

        }
        
        # remover licença
        $msu = Get-MsolUser -UserPrincipalName $UPN

        if ($msu.Licenses.AccountSkuId -contains $SKU) {

            Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "$SKU" 2>&1
            Log("$UPN;removida licença $SKU")
        }
        else {
            Log("$UPN;ATENCAO: previamente removida licença $SKU")
        }

    }
    else {

        Log  "$UPN;ação $Acao inválida"
        exit 4
    }
    

}

try {

    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
}
catch {
    Start-Sleep 10
    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
}
