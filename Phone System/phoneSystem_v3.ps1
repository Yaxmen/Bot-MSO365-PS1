Param (

    [Parameter( Mandatory=$true)] [ValidateSet("add","del")] [String]$Acao, 
    [Parameter( Mandatory=$true)] [String]$Chave, 
    [Parameter( Mandatory=$false)] [String] $Categoria,
    [Parameter( Mandatory=$false)][String]$Ramal
)

#Import-Module Microsoft.Graph.Users
#Import-Module Microsoft.Graph.Groups

#Setando constantes
$ClientId = "b88e924b-530d-497c-8b16-b54456064e5f"
$TenantId = "5b6f6241-9a57-4be4-8e50-1dfa72e79a57"
$key = (1..16)
$Secret = Get-Content "D:\Password\GraphAppPassword.txt" | ConvertTo-SecureString -Key $key -ErrorAction Stop

$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $Secret
$SKU = "MCOEV"


$GUID = "f55f89fa-f356-4aa3-a295-180a08fa9957" #GN_PHONE_SYSTEM_SEGUNDA_LINHA
$GUIDLIC = "668cf70e-ff6d-483e-872f-097b954b87b4" #GN_LIC_PHONE_SYSTEM

$SLEEP = 60
$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyyMMdd')+"_phoneSystem_log.csv"


# Definindo credenciais de acesso a tenant
$username = "samsazu@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop


#$username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
#$PlainPassword="Ror66406"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword

#Função de Log
function Log([string]$message){

    $datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    if ( $message -like '*ERRO*') {
        Write-Host "$datetime;$message" -ForegroundColor Red

    }

    elseif($message -like '*ATENCAO*') { 
        
        Write-Host "$datetime;$message" -ForegroundColor Yellow

    } else {
        
        Write-Host "$datetime;$message" -ForegroundColor Green
        
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

# Conectando com o MicrosoftTeams
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




### Inicio das Funções ###

<# Função responsável por definir e validar o ramal, conceder Licenças,
configurar linha de ramal e adicionar o usuário ao grupo no AzureAD #>

function AddRamal($Ramal, $Categoria, $User, $TeamsUser){

    
    #Criando Hash com códigos válidos
    $CodHash = @{
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
	'853411' = 'Tag:DP-85-FLA-N'
	'923627' = 'Tag:DP-92-CPD-N'
	'613429' = 'Tag:DP-61-BSA-N'
	'243361' = 'Tag:DP-BR-24-N'
    }

    #Validando o Ramal

    $cod = $Ramal -match '\+55(\d\d\d\d\d\d)\d\d\d\d'

    #Encerra processamento se ramal foi enviado no formato errado

    if (-not($cod)) {
        Log "$($User.UserPrincipalName);*ERRO*: $Ramal com formato inválido"
        exit 1
    } 

    #Validando se o Código recebido é válido, caso contrário encerra o Script

    $cod = $matches[1]
    $plano = $CodHash[$cod]

    if (-not($CodHash.ContainsKey($cod))) {
        Log "$($User.UserPrincipalName);*ERRO*: não encontrado plano para código $cod"
        exit 1
    }

    $Linha = "tel:$Ramal"

    #Verifica se o usuário já esta configurado com o Ramal solicitado
    if($TeamsUser.LineURI -eq $Linha){

        Log "$UPN;ATENCAO: já configurado Linha $($TeamsUser.LineURI) - $($TeamsUser.OnlineVoiceRoutingPolicy) - $($TeamsUser.TenantDialPlan)"
        exit 1
    }


    #Verifica se o ramal solicitado está disponível
    try{
    
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq '$Linha'" -ErrorAction SilentlyContinue

    } catch{
        
        Log "$($User.UserPrincipalName); *ERRO*: Erro ao executar Get-CsOnlineUser"
        exit 1

    }

    if($null -ne $OnlineRamal){
        
        $chave = Get-ADUser -Filter "UserPrincipalName -eq '$($OnlineRamal.UserPrincipalName)'" | select SamAccountName | %{$_.SamAccountName}
        #$status_ramal =  "$($OnlineRamal.Alias)"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"
        $status_ramal =  "$chave"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"       
        Log "$UPN;ATENCAO: previamente configurado o numero $Ramal para chave $status_ramal"

        exit 1
    }

    #Verifica se o usuário já possui Ramal

    if($($TeamsUser.LineURI) -ne $null -and $($TeamsUser.LineURI) -ne ""){
    
        Log "$($User.UserPrincipalName);*ATENÇÃO*: Usuário já possui linha configurada"
        #exit 1
    }


    #Resgata as licenças do usuário
    $UserLicenses = Get-MgUserLicenseDetail -UserId $User.Id

    #Verifica se o usuário tem licença, caso não tenha atribui uma a ele
    if($UserLicenses.SkuPartNumber -notcontains $SKU){

        $TenantLicenses = Get-MgSubscribedSku

        #Resgatando a quantidade de Licenças disponíveis
        Foreach($License in $TenantLicenses){
        
            if($License.SkuPartNumber -eq $SKU){
                
                $MCOEVSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq "MCOEV" 
                $DispMCOEV = $License.PrepaidUnits.Enabled
            }
        }

        #Verifica se existe licença disponível
        if($DispMCOEV -lt 1){
            
            Log "$($User.UserPrincipalName);*ERRO*: Não há Licenças $SKU disponíveis para atribuição"
            exit 1
        }
        $RefObjectId = $User.Id
        try{

        #Atribui a licença para o usuário
        Set-MgUserLicense -UserId $User.Id -AddLicense @{SkuId = $MCOEVSku.SkuId} -RemoveLicenses @() -ErrorAction Stop | Out-Null

        } catch {
            
            Log "*ERRO*: ao atribuir a licença ao usuário"
            exit 1        
        }

        # Resgatando os Membros do grupo de Licenças
        try{
            $GroupLIC = Get-MgGroupMember -GroupId $GUIDLIC -All
        } catch {
            Log "*ERRO*: Ao resgatar membros do grupo"
        }

        # Verificando se o usuário já está no grupo

        if($GroupLIC.Id -notcontains $RefObjectId){
            
            try {

                New-MgGroupMember -GroupId $GUIDLIC -DirectoryObjectId $RefObjectId -ErrorAction Stop | Out-Null

            } catch {
                Log "$($User.UserPrincipalName); *ERRO*: Errro ao atribuir a licença"
                exit 1
            }

        } else {
            Log "*ATENCAO*: Usuário já esta no grupo de licenças"
        }

        

        #Log "$UPN;ATENCAO: aguardando aplicar licença do phoneSystem aguardar + $SLEEP seg"
        $SLEPPT = 60
        $NOK = $true

        ### aguarda +20 seg ate que seja verificada a aplicacao da licença no modulo Microsoft Teams
        ### no maximo aguarda ate 300 seg

        while ($NOK) {
            Start-Sleep $SLEEP
            $resp = Get-CsOnlineUser -Identity $User.UserPrincipalName
            if ($resp.FeatureTypes.Contains("PhoneSystem")) {
                $NOK = $false
                Log "$UPN;SUCESSO: licenca phoneSystem aplicada em $SLEEPT seg"
            }
            else {
                $SLEEP = 20
                $SLEEPT += $SLEEP
                Log "$($User.UserPrincipalName);*ATENCAO*: aguardando aplicar licença do phoneSystem aguardar + $SLEEP seg"
                if ($SLEEPT -gt 600) {
                    Log "$($User.UserPrincipalName);*ERRO*: licenca phoneSystem NÃO aplicada após $SLEEPT seg"
                    exit 1
                }
                   
            }
        }    
    } else {
        Log "$($User.UserPrincipalName);*ATENÇÃO*: Licença já configurada previamente"
        ##Usuário já possuia Licença
    }




    # Capturando membros do grupo
    $Group = Get-MgGroupMember -GroupId $GUID -All

    # Verificando se o usuário já está no grupo
    if ($Group.id -notcontains $User.id) {

        #Adicionando o usuário ao grupo
        try{

            New-MgGroupMember -GroupId $GUID -DirectoryObjectId $User.Id -ErrorAction Stop | Out-Null
    
        } catch {
        
            Log "$($User.UserPrincipalName);*ERRO*: erro ao adicionar ao grupo: $GUID"
            exit 1
        }

    } else {
        
        Log "*ATENCAO*: Usuário já incluso no grupo"
    }



   
    #Atribuindo o Ramal ao usuário
    try {
        $cod = Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumber $Ramal -PhoneNumberType DirectRouting | Out-Null
    } catch {
        Log "$($User.UserPrincipalName); *ERRO*: Erro ao atribuir o ramal"
    }
    if ($cod -contains "has already been assigned to another user") {
        Start-Sleep 30
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq '$Ramal'"
        $status_ramal =  "$($OnlineRamal.Alias)"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"
        Log "$($User.UserPrincipalName);*ATENCAO*: previamente configurado o numero $Ramal para chave $status_ramal" 
        exit 1
    }


    try {

        Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -EnterpriseVoiceEnabled $true | Out-Null

    } catch {
        
        Log "$($User.UserPrincipalName); *ERRO*: Erro ao ativar EnterpriseVoice"
    }

    #Start-Sleep 30

    #$csu = Get-CsOnlineUser -Identity $User.UserPrincipalName
    #if ($($csu.LineUri) -eq $linha -and $($csu.EnterpriseVoiceEnabled)) {
     #   Log "$($User.UserPrincipalName);configurado: EnterpriseVoice True - HostedVoiceMail True - LineURI $Ramal"
    #}
    #else {
     #   Log "$($User.UserPrincipalName);ERRO: EnterpriseVoice True - HostedVoiceMail True - LineURI $linha"
      #  exit 1
    #}
    
    
    Start-Sleep 10





    ### define categoria
                        
    if ($Categoria -eq "DDD") {
        $vrota = "Tag:VRP-BR-Nacional"

    } elseif($Categoria -eq "DDI") {
        $vrota = "Tag:VRP-BR-Internacional"

    } else {
        $vrota = "Tag:VRP-BR-Interno"  
    }


    
    try {
        Grant-CsOnlineVoiceRoutingPolicy -Identity $User.UserPrincipalName -PolicyName $vrota 2>D:\Util\PhoneSystem\logs\err_CsOnlineVoiceRoutingPolicy.txt | Out-Null
    }
    catch {
        Log "$($User.UserPrincipalName);*ERRO*: OnlineVoiceRoutingPolicy $vrota"
        exit 1
    }

    Log "$($User.UserPrincipalName);configurado: OnlineVoiceRoutingPolicy $vrota"
    Start-Sleep 10 
    
    try {   
        Grant-CsTenantDialPlan -Identity $User.UserPrincipalName -PolicyName $plano 2>D:\Util\PhoneSystem\logs\err_CsTenantDialPlan.txt | Out-Null
    }
    catch {
        Log "$($User.UserPrincipalName);*ERRO*: TenantDialPlan $plano"
        exit 1
    }
    
    Log "$($User.UserPrincipalName);configurado: TenantDialPlan $plano"

    #Desconectando do MgGraph

    try{
        Disconnect-MgGraph | Out-Null
    } catch {
        Start-Sleep 10
        Disconnect-MgGraph | Out-Null
    }

    #Desconectando do Microsoft Teams
    try {

    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null

    } catch {
        Start-Sleep 10
        Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
    }
          

}



<# Função responsável por remover licenças do usuário, desconfigurar 
a linha de ramal e remover o usuário do grupo no AzureAD #>

function RemoveRamal($User, $TeamsUser){
    
    #Verificando se o usuário tem linha configurada

    if($null -eq $TeamsUser.LineURI){
        
        Log "$UPN;ATENCAO: usuário não está configurado"
        exit 1

    } 

    #Removendo o número de telefone
    try{

        Grant-CsTenantDialPlan -Identity $User.UserPrincipalName -PolicyName $null
        Grant-CsOnlineVoiceRoutingPolicy -Identity $User.UserPrincipalName -PolicyName $null
        Remove-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -RemoveAll -ErrorAction Stop | Out-Null
    } catch {
        
        Log "$($User.UserPrincipalName);*ERRO*: Erro ao tentar desconfigurar Linha"
        exit 1
    }
    
    <#
    #Capturando as Licenças do usuário
    $UserLicenses = Get-MgUserLicenseDetail -UserId $User.Id

    #Verifica se o usuário possui a licença
    if($UserLicenses.SkuPartNumber.Contains($SKU)){
        
        #Resgata o Id da Licença
        foreach($License in $UserLicenses){
            if($License.SkuPartNumber -eq $SKU){
                $MCOEVSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'MCOEV'
            }
        }
       
        #Removendo a licença do usuário
        try{
            Set-MgUserLicense -UserId $User.Id -RemoveLicense @($MCOEVSku.SkuId) -AddLicenses @{} -ErrorAction Stop | Out-Null
        } catch {
            Log "$($User.UserPrincipalName);*ERRO*: Erro ao remover a licença do usuário"
            exit 1
        }

    } else {
    
        Log "$($User.UserPrincipalName);*ATENÇÃO*: Não foi possível remover a licença $SKU. Não localizada"
    }

    #>
    

    #Remover usuário do Grupo de Politica
    try{
        Remove-MgGroupMemberByRef -GroupId $GUID -DirectoryObjectId $User.Id -ErrorAction Stop | Out-Null
    } catch {
        Log "$($User.UserPrincipalName);*ERRO*: Erro ao remover usuário do grupo $GUID"
        exit 1
    }

    #Remover usuário do Grupo de Licença
    try{
        Remove-MgGroupMemberByRef -GroupId $GUIDLIC -DirectoryObjectId $User.Id -ErrorAction Stop | Out-Null
    } catch {
        Log "$($User.UserPrincipalName);*ERRO*: Erro ao remover usuário do grupo $GUIDLIC"
        exit 1
    }

    Write-Output "Sucesso ao remover o Ramal"

    #Desconectando do MgGraph

    try{
        Disconnect-MgGraph | Out-Null
    } catch {
        Start-Sleep 10
        Disconnect-MgGraph | Out-Null
    }

    #Desconectando do Microsoft Teams
    try {

    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null

    } catch {
        Start-Sleep 10
        Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
    }

}

### Fim das Funções ###






### Inicio do Processamento ####

#Capturando o usuário no AD
try{
    $ADUser = Get-ADUser $Chave
    #$User = Get-MgUser -Filter "UserPrincipalName eq $($ADUser.UserPrincipalName)"
    $user = Get-MgUser -UserId $ADUser.UserPrincipalName
} catch {

    Log "$($Chave);*ERRO*: Chave $Chave não localizada"
    exit 1 
}

#Capturando o Usuario no MsTeams

try{
    
    $TeamsUser = Get-CsOnlineUser -Identity $User.Id

} catch {
    
    Log "$($User.UserPrincipalName);*ERRO*: Usuário nao localizado no Teams"
    exit 1
}



#Definindo qual a função que o script vai realizar
if($Acao -eq "add") {

    AddRamal -Ramal $Ramal -Categoria $Categoria -User $User -TeamsUser $TeamsUser

} elseif ($Acao -eq "del"){

    RemoveRamal -User $User -TeamsUser $TeamsUser

}

### Fim do Processamento ###

