

$LISTA = Import-Excel -Path 'D:\Util\PhoneSystem\reports\ramais_teams_em_lote_chave_servico.xlsx'


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

<#$username = "samig01@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop#>

# Definindo credenciais de acesso a tenant
$username = "samsazu@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyy')+"_phoneSystem_lote_log.csv"

#Setando constantes
$ClientId = "b88e924b-530d-497c-8b16-b54456064e5f"
$TenantId = "5b6f6241-9a57-4be4-8e50-1dfa72e79a57"
$key = (1..16)
$Secret = Get-Content "D:\Password\GraphAppPassword.txt" | ConvertTo-SecureString -Key $key -ErrorAction Stop

$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $Secret
$SKU = "MCOEV"

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

#Conectando com o MicrosoftTeams
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

#Conectando ao AzureAD
<#try {

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
}#>

foreach ($user in $LISTA) {

  $Chave = $user.Chave
  $Ramal = $user.Ramal
  $Categoria = $user.Categoria
  $Acao = $user.Acao

  write-host "$Acao $Chave $Ramal $Categoria"
  
  $SLEEP = 60

### consulta, no AD interno qual $UPN da $Chave informada. Caso não encontre encerra o script.



try {

  $azu = Get-ADUser $Chave
  $UPN = $azu.UserPrincipalName
  Log "$UPN;SUCESSO: chave $chave encontrada no AD"
}
catch {
    Log "$chave;ERRO: chave não encontrada no AD"
    # exit 1  ==>> o exit sairia do script, por isso inserido comentário
}
    
### verifica as configurações do usuario $UPN no diretorio do Teams.

try {
    #$OnlineUser = Get-CsOnlineUser -Filter "UserPrincipalName -eq '$UPN'"
    $OnlineUser = Get-CsOnlineUser -Identity $UPN
}
catch {1
    Log "$UPN;ERRO: ao executar Get-CsOnlineUser -Filter"
    
}

### encerra o script, caso dados do usuario $UPN retorne nulo no diretório do Teams

if ($OnlineUser -eq $null) {
    Log "$UPN;ERRO: usuário inválido"
    
}

### configurando phoneSystem (acao = add)

if ($Acao -eq "add") {

    $cod_hash = @{

    '113523' = 'Tag:DP-11-SAO-N'
	'113795' = 'Tag:DP-11-MAU-N'
	'123886' = 'Tag:DP-12-CGA-N'
	'123928' = 'Tag:DP-12-SJC-N'
	'133249' = 'Tag:DP-BR-13-N'
	'133328' = 'Tag:DP-13-RSA-N'
	'133362' = 'Tag:DP-13-EZR-N'
	'192116' = 'Tag:DP-19-PLA-N'
	'212133' = 'Tag:DP-21-IOY-N'
	'212144' = 'Tag:DP-21-SNDO-N'
	'212162' = 'Tag:DP-21-IFO-N'
	'212166' = 'Tag:DP-21-SNDO-N'
	'212167' = 'Tag:DP-21-CDL-N'
	'212665' = 'Tag:DP-21-BLS-N'
	'212677' = 'Tag:DP-21-CES-N'
	'213224' = 'Tag:DP-21-RJO-N'
	'213227' = 'Tag:DP-21-TMO-N'
	'213876' = 'Tag:DP-21-MNA-N'
	'222101' = 'Tag:DP-BR-22-N'
	'223377' = 'Tag:DP-BR-22-N'
	'223379' = 'Tag:DP-BR-22-N'
	'222797' = 'Tag:DP-22-CBS-N'
	'273048' = 'Tag:DP-27-UTC-N'
	'273295' = 'Tag:DP-BR-27-N'
	'283360' = 'Tag:DP-28-UTA-N'
	'313472' = 'Tag:DP-31-ITO-N'
	'313529' = 'Tag:DP-31-BET-N'
	'323239' = 'Tag:DP-32-TJF-N'
	'413641' = 'Tag:DP-41-AUC-N'
	'513415' = 'Tag:DP-51-CAN-N'
	'673509' = 'Tag:DP-67-TLS-N'
	'813879' = 'Tag:DP-81-RAL-N'
	'273771' = 'Tag:DP-BR-27-N'
	'473406' = 'Tag:DP-BR-47-N'
	'713348' = 'Tag:DP-71-SDR-N'
	'713417' = 'Tag:DP-71-SFCO-N'
	'713502' = 'Tag:DP-71-SGO-N'
	'713617' = 'Tag:DP-71-TQE-N'
	'753366' = 'Tag:DP-75-ACK-N'
	'793212' = 'Tag:DP-79-AJU-N'
	'843303' = 'Tag:DP-84-NTL-N'
	'843235' = 'Tag:DP-84-TASSU-N'
	'853411' = 'Tag:DP-85-TCE-N'
	'923627' = 'Tag:DP-92-CPD-N'
	'613429' = 'Tag:DP-61-BSA-N'
    '243361' = 'Tag:DP-BR-24-N'
    }

    $cod = $Ramal -match '\+55(\d\d\d\d\d\d)\d\d\d\d'

    ### encerra processamento se ramal foi enviado no formato errado

    if (-not($cod)) {
        Log "$UPN;ERRO: $Ramal com formato inválido"
        
    }

    ### encerra processamento se codigo DDD não constar na tabela $cod_hash

    $cod = $matches[1]
    $plano = $cod_hash[$cod]

    if (-not($cod_hash.ContainsKey($cod))) {
        Log "$UPN;ERRO: não encontrado plano para código $cod"
        
    }

    $linha = "tel:"+"$Ramal"

    ### verifica se usuario já esta configurado com ramal solicitado

    if ($($OnlineUser.LineURI) -eq $linha) {

        Log "$UPN;ATENCAO: já configurado Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"
        
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
   #     exit 1
                
    }

    if ($($OnlineUser.LineURI) -ne $null -and $($OnlineUser.LineURI) -ne "") {
    
        Log "$UPN;ATENCAO: já configurado Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"
     #   exit 
    }

    ### verifica se usuario possui licença e aplica caso não tenha
    #$msu = Get-MsolUser -UserPrincipalName $UPN

    $user = Get-MgUser -UserId $UPN
    $UserLicenses = Get-MgUserLicenseDetail -UserId $User.Id

    if ($UserLicenses.SkuPartNumber -notcontains $SKU) {

        # verifica se existem licenças MCOEV disponiveis no ambiente

        #$plans = Get-MsolAccountSku
        $plans = Get-MgSubscribedSku
        foreach ($i in $plans) {

            <#if ($($i.AccountSkuId) -eq "$SKU") {
                $disp_MCOEV = $($i.ActiveUnits) - $($i.ConsumedUnits)
            }#>
                if($i.SkuPartNumber -eq $SKU){
                
                $MCOEVSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq "MCOEV" 
                $disp_MCOEV = $i.PrepaidUnits.Enabled
            }

        }

        if ($disp_MCOEV -lt 1) {
##           Log "$UPN;ERRO: $($disp_MCOEV) $SKU - insuficientes para configurar usuario"
            #exit 1
        }

        #Log "$UPN;ATENCAO: $($disp_MCOEV) $SKU - disponiveis" 

        ### aplica licença e aguarda 60 seg

        #Add-AzureADGroupMember -ObjectId $GUIDLIC -RefObjectId $OnlineUser.Identity
		#Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "$SKU"


        Set-MgUserLicense -UserId $user.Id -AddLicense @{SkuId = $MCOEVSku.SkuId} -RemoveLicenses @() -ErrorAction Stop | Out-Null
        New-MgGroupMember -GroupId $GUIDLIC -DirectoryObjectId $user.Id -ErrorAction Stop | Out-Null


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
                if ($SLEEPT -gt 360) {
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
           New-MgGroupMember -GroupId $GUID -DirectoryObjectId $user.Id
           #Add-AzureADGroupMember -ObjectId $GUID -RefObjectId $OnlineUser.Identity
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
     #   exit 1
    }

    Set-CsPhoneNumberAssignment -Identity $UPN -EnterpriseVoiceEnabled $true

    Start-Sleep 30

    $csu = Get-CsOnlineUser -Identity $UPN
    if ($($csu.LineUri) -eq $linha -and $($csu.EnterpriseVoiceEnabled)) {
        Log "$UPN;configurado: EnterpriseVoice True - HostedVoiceMail True - LineURI $Ramal"
    }
    else {
        Log "$UPN;ERRO: EnterpriseVoice True - HostedVoiceMail True - LineURI $linha"
        
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
        
    }

    Log "$UPN;configurado: OnlineVoiceRoutingPolicy $vrota"
    Start-Sleep 10 
    
    try {   
        Grant-CsTenantDialPlan -Identity $UPN -PolicyName $plano 2>D:\Util\PhoneSystem\logs\err_CsTenantDialPlan.txt
    }
    catch {
        Log "$UPN;ERRO: TenantDialPlan $plano"
    #    exit 1        
    }
            Log "$UPN;configurado: TenantDialPlan $plano"
            
}
else {

    if ($Acao -eq "del") {

        if ($OnlineUser.LineURI -eq $null)
        {
            Log "$UPN;ATENCAO: usuário não está configurado"
     #       exit 1            
        }

        Log "$UPN;antes de remover - Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"

        try {

            #Grant-CsTenantDialPlan -Identity $UPN -PolicyName $null
            #Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $null
            Remove-CsPhoneNumberAssignment -Identity $UPN -RemoveAll
        }
        catch {

            Log "$UPN;ERRO: uao desconfigurar phoneSystem"
      #       exit 1           

        }
        
        # remover licença
        #$msu = Get-MsolUser -UserPrincipalName $UPN
        $user = Get-MgUser -UserId $UPN
        $UserLicenses = Get-MgUserLicenseDetail -UserId $User.Id

 if ($UserLicenses.SkuPartNumber -contains $SKU) {
	        #Remove-AzureADGroupMember -ObjectId $GUIDLIC -MemberId $OnlineUser.Identity
            Remove-MgGroupMemberByRef -GroupId $GUIDLIC -DirectoryObjectId $user.Id


            #Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "$SKU" 2>&1
            Log("$UPN;removida licença $SKU")
        }
        else {
            Log("$UPN;ATENCAO: previamente removida licença $SKU")
        }

        try {
           #Remove-AzureADGroupMember -ObjectId $GUID -MemberId $OnlineUser.Identity
           Remove-MgGroupMemberByReF -GroupId $GUID -DirectoryObjectId $user.Id
           Log("$UPN;removido do grupo $GUID")
        }
        catch {
           Log("$UPN;ERRO; ao remover do grupo $GUID") 

        }
          
		  
    }
    else {

        Log  "$UPN;ação $Acao inválida"
   #      exit 4       
    }
    

}

    #Write-Host " Esperando a MS se recuperar em 30 seg ..."
    #Start-Sleep 30


} # end loop lista

try {

    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
}
catch {
    Start-Sleep 10
    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
}

