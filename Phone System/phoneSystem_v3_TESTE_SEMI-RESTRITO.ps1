<#

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
04-03-22 alterados comandos para configuração e remoção do phoneSystem  para Set-CsPhoneNumberAssignment e Remove-CsPhoneNumberAssignment
16-03-22 adicionada verificacao de licencas disponiveis e alterada tempo espera apos aplicar licenca
16-03-23 Inclusão de novos prefixos/dial plans (verificação de dial plan com 6 dígitos = CNXYZW)
05-05-23 Inclusão de categoria SEMI-RESTRITO na oferta do Service Now            


Get-NetFirewallRule 

#>

Param (

    [Parameter( Mandatory=$true)] [ValidateSet("add","del")] [String]$Acao, 
    [Parameter( Mandatory=$true)] [String]$Chave, 
    [Parameter( Mandatory=$false)] [ValidateSet("DDD","DDI","SEMI-RESTRITO")][String]$Categoria,
    [Parameter( Mandatory=$false)][String]$Ramal
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$username = "samig01@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

$SLEEP = 60
$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyy')+"_phoneSystem_log.csv"
$SKU = "petrobrasbr:MCOEV"
$GUID = "f55f89fa-f356-4aa3-a295-180a08fa9957" #GN_PHONE_SYSTEM_SEGUNDA_LINHA
$GUIDLIC = "668cf70e-ff6d-483e-872f-097b954b87b4" #GN_LIC_PHONE_SYSTEM

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


try {

    Connect-AzureAD -Credential $credential -InformationAction SilentlyContinue | Out-Null
    Log "SUCESSO: ao conectar AzureAD 1/3"
}
catch {
    Log "ATENCAO: ao conectar AzureAD 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-AzureAD -Credential $credential -InformationAction SilentlyContinue | Out-Null
        Log "SUCESSO: ao conectar AzureAD 2/3"
    }
    catch {
        Log "ATENCAO: ao conectar AzureAD 2/3"
        Start-Sleep $SLEEP
        try {
            Connect-AzureAD -Credential $credential -InformationAction SilentlyContinue | Out-Null
            Log "SUCESSO: ao conectar AzureAD 2/3"
        }
        catch {    
            Log "ERRO: ao conectar AzureAD 3/3"
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
    #$OnlineUser = Get-CsOnlineUser -Filter "UserPrincipalName -eq '$UPN'"
    $OnlineUser = Get-CsOnlineUser -Identity $UPN
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

    '113523' = 'Tag:DP-11-SAO'
	'113795' = 'Tag:DP-11-MAU'
	'123886' = 'Tag:DP-12-CGA'
	'123928' = 'Tag:DP-12-SJC'
	'133249' = 'Tag:DP-BR-13'
	'133328' = 'Tag:DP-13-RSA'
	'133362' = 'Tag:DP-13-EZR'
	'192116' = 'Tag:DP-19-PLA'
	'212133' = 'Tag:DP-21-IOY'
	'212144' = 'Tag:DP-21-SNDO'
	'212162' = 'Tag:DP-21-IFO'
	'212166' = 'Tag:DP-21-SNDO'
	'212167' = 'Tag:DP-21-CDL'
	'212665' = 'Tag:DP-21-BLS'
	'212677' = 'Tag:DP-21-CES'
	'213224' = 'Tag:DP-21-RJO'
	'213227' = 'Tag:DP-21-TMO'
	'213876' = 'Tag:DP-21-MNA'
	'223377' = 'Tag:DP-BR-22'
	'223379' = 'Tag:DP-BR-22'
	'222797' = 'Tag:DP-22-CBS'
	'273048' = 'Tag:DP-27-UTC'
	'273295' = 'Tag:DP-BR-27'
	'283360' = 'Tag:DP-28-UTA'
	'313472' = 'Tag:DP-31-ITO'
	'313529' = 'Tag:DP-31-BET'
	'323239' = 'Tag:DP-32-TJF'
	'413641' = 'Tag:DP-41-AUC'
	'513415' = 'Tag:DP-51-CAN'
	'673509' = 'Tag:DP-67-TLS'
	'813879' = 'Tag:DP-81-RAL'

    }

    $cod = $Ramal -match '\+55(\d\d\d\d\d\d)\d\d\d\d'

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
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq '$ramal'"
    }
    catch {
        Log "$UPN;ERRO: ao executar Get-CsOnlineUser -Filter Ramal"
        exit 1
    }

    ### encerra script caso ramal solicitado já esteja em uso

    if ($OnlineRamal -ne $null) {
        $chave = Get-ADUser -Filter "UserPrincipalName -eq '$($OnlineRamal.UserPrincipalName)'" | select SamAccountName | %{$_.SamAccountName}
        #$status_ramal =  "$($OnlineRamal.Alias)"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"
        $status_ramal =  "$chave"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"       
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

        # verifica se existem licenças MCOEV disponiveis no ambiente

        $plans = Get-MsolAccountSku

        foreach ($i in $plans) {

            if ($($i.AccountSkuId) -eq "$SKU") {
                $disp_MCOEV = $($i.ActiveUnits) - $($i.ConsumedUnits)
            }
        }

        if ($disp_MCOEV -lt 1) {
##           Log "$UPN;ERRO: $($disp_MCOEV) $SKU - insuficientes para configurar usuario"
            #exit 1
        }

        #Log "$UPN;ATENCAO: $($disp_MCOEV) $SKU - disponiveis" 

        ### aplica licença e aguarda 60 seg
	    Add-AzureADGroupMember -ObjectId $GUIDLIC -RefObjectId $OnlineUser.Identity
        #Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "$SKU"
        #Log "$UPN;ATENCAO: aguardando aplicar licença do phoneSystem aguardar + $SLEEP seg"
        $SLEPPT = $SLEEP
        $NOK = $true

        ### aguarda +20 seg ate que seja verificada a aplicacao da licença no modulo Microsoft Teams
        ### no maximo aguarda ate 600 seg

        while ($NOK) {
            Start-Sleep $SLEEP
            $resp = Get-CsOnlineUser -Identity $UPN | select FeatureTypes
            if ($($resp.FeatureTypes) -ccontains "PhoneSystem") {
                $NOK = $false
                #Log "$UPN;SUCESSO: licenca phoneSystem aplicada em $SLEPPT seg"
            }
            else {
                $SLEEP = 20
                $SLEEPT += $SLEEP
                Log "$UPN;ATENCAO: aguardando aplicar licença do phoneSystem aguardar + $SLEEP seg"
                if ($SLEEPT -gt 600) {
                    Log "$UPN;ERRO: licenca phoneSystem NÃO aplicada após $SLEEPT seg"
                    $NOK = $false
                }
                   
            }
        }    
        
    }
    else {
        Log "$UPN;ATENCAO: previamente configurado $SKU"
    }
    
    Start-Sleep 30

    try {
           Add-AzureADGroupMember -ObjectId $GUID -RefObjectId $OnlineUser.Identity
           Log("$UPN;adicionado ao grupo $GUID")
    }
    catch {
           Log("$UPN;ERRO; ao adicionar do grupo $GUID") 

    }


    
    $cod = Set-CsPhoneNumberAssignment -Identity $UPN -PhoneNumber $Ramal -PhoneNumberType DirectRouting
    if ($cod -ccontains "has already been assigned to another user") {
        Start-Sleep 30
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq '$ramal'"
        $chave = Get-ADUser -Filter "UserPrincipalName -eq '$($OnlineRamal.UserPrincipalName)'" | select SamAccountName | %{$_.SamAccountName}
        #$status_ramal =  "$($OnlineRamal.Alias)"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"
        $status_ramal =  "$chave"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"  
        Log "$UPN;ATENCAO: previamente configurado o numero $Ramal para chave $status_ramal" 
        exit 1
    }

    Set-CsPhoneNumberAssignment -Identity $UPN -EnterpriseVoiceEnabled $true

    Start-Sleep 30

    $csu = Get-CsOnlineUser -Identity $UPN
    if ($($csu.LineUri) -eq $linha -and $($csu.EnterpriseVoiceEnabled)) {
        Log "$UPN;configurado: EnterpriseVoice True - HostedVoiceMail True - LineURI $Ramal"
    }
    else {
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

            #Grant-CsTenantDialPlan -Identity $UPN -PolicyName $null
            #Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $null
            Remove-CsPhoneNumberAssignment -Identity $UPN -RemoveAll
        }
        catch {

            Log "$UPN;ERRO: uao desconfigurar phoneSystem"
            exit 1

        }
        
        # remover licença
        $msu = Get-MsolUser -UserPrincipalName $UPN

        if ($msu.Licenses.AccountSkuId -contains $SKU) {
	        Remove-AzureADGroupMember -ObjectId $GUIDLIC -MemberId $OnlineUser.Identity
            #Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "$SKU" 2>&1
            Log("$UPN;removida licença $SKU")
        }
        else {
            Log("$UPN;ATENCAO: previamente removida licença $SKU")
        }

        try {
           Remove-AzureADGroupMember -ObjectId $GUID -MemberId $OnlineUser.Identity
           Log("$UPN;removido do grupo $GUID")
        }
        catch {
           Log("$UPN;ERRO; ao remover do grupo $GUID") 

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

