Param (

    [Parameter( Mandatory=$true)] [ValidateSet("add","del")] [String]$Acao, 
    [Parameter( Mandatory=$true)] [String]$Chave, 
    [Parameter( Mandatory=$true)] [ValidateSet("DDD","DDI")][String]$Categoria,
    [Parameter( Mandatory=$true)][String]$Ramal
)

Import-Module Microsoft.Graph.Users

#Setando constantes
$ClientId = "60c5c5b3-84dd-4519-8ad7-58a34f8d39b4"
$TenantId = "6af8f826-d4c2-47de-9f6d-c04908aa4e88"
$Certificate = "E02B30DF5D23D49516FE28029C577D5A13862F8B"
$SKU = "MCOEV"
#$GUID = "f55f89fa-f356-4aa3-a295-180a08fa9957"
$GUID = "0018869d-60f6-4a54-ba71-2813bdcd8d8d"
$SLEEP = 60
$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyy')+"_phoneSystem_log.csv"


# Definindo credenciais de acesso a tenant
#$username = "@.com.br"
#$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
#$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
#$Credentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop


#Conectando no MicrosoftGraph
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Certificate

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


# Conectando com o MicrosoftTeams
try {

    Connect-MicrosoftTeams #-Credential $credential -InformationAction SilentlyContinue | Out-Null
    Log "SUCESSO: ao conectar MicrosoftTeams 1/3"
}
catch {
    Log "ATENCAO: ao conectar MicrosoftTeams 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-MicrosoftTeams #-Credential $credential -InformationAction SilentlyContinue | Out-Null
        Log "SUCESSO: ao conectar MicrosoftTeams 2/3"
    }
    catch {
        Log "ATENCAO: ao conectar MicrosoftTeams 2/3"
        Start-Sleep $SLEEP
        try {
            Connect-MicrosoftTeams #-Credential $credential -InformationAction SilentlyContinue | Out-Null
            Log "SUCESSO: ao conectar MicrosoftTeams 2/3"
        }
        catch {    
            Log "*ERRO*: ao conectar MicrosoftTeams 3/3"
            exit 1
        }

    }
}




### Inicio das Funções ###
function AddRamal($Ramal, $Categoria, $User, $TeamsUser){

    
    #Criando Hash com códigos válidos
    $CodHash = @{
        '21' = 'Tag:DP-BR-21'
        '22' = 'Tag:DP-BR-22'
        '13' = 'Tag:DP-BR-13'
        '27' = 'Tag:DP-BR-27'
    }

    #Validando o Ramal

    $cod = $Ramal -match '\+55(\d\d)\d\d\d\d\d\d\d\d'

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

        Log "$($User.UserPrincipalName);*ERRO*: Linha $Linha já configurada para o usuário"
        exit 1
    }


    #Verifica se o ramal solicitado está disponível
    try{
    
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq $Linha" -ErrorAction SilentlyContinue

    } catch{
        
        Log "$($User.UserPrincipalName); *ERRO*: Erro ao executar Get-CsOnlineUser"
        exit 1

    }

    if($null -ne $OnlineRamal){

        Log "$($User.UserPrincipalName);*ERRO*: Ramal solicitado já está configurado para a Chave $($OnlineRamal.Alias)"
        exit 1
    }

    #Verifica se o usuário já possui Ramal

    if($($TeamsUser.LineURI) -ne $null -and $($TeamsUser.LineURI) -ne ""){
    
        Log "$($User.UserPrincipalName);*ERRO*: Usuário já possui linha configurada"
        exit 1
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

        try{
        #Atribui a licença para o usuário
            Set-MgUserLicense -UserId $User.Id -AddLicense @{SkuId = $MCOEVSku.SkuId} -RemoveLicenses @() -ErrorAction Stop
        } catch {
            Log "$($User.UserPrincipalName); *ERRO*: Errro ao atribuir a licença"
            exit 1
        }

        #Log "$UPN;ATENCAO: aguardando aplicar licença do phoneSystem aguardar + $SLEEP seg"
        $SLEPPT = $SLEEP
        $NOK = $true

        ### aguarda +20 seg ate que seja verificada a aplicacao da licença no modulo Microsoft Teams
        ### no maximo aguarda ate 600 seg

        while ($NOK) {
            Start-Sleep $SLEEP
            $resp = Get-CsOnlineUser -Identity $User.UserPrincipalName
            if ($resp.FeatureTypes.Contains("PhoneSystem")) {
                $NOK = $false
                #Log "$UPN;SUCESSO: licenca phoneSystem aplicada em $SLEPPT seg"
            }
            else {
                $SLEEP = 20
                $SLEEPT += $SLEEP
                Log "$($User.UserPrincipalName);*ATENCAO*: aguardando aplicar licença do phoneSystem aguardar + $SLEEP seg"
                if ($SLEEPT -gt 600) {
                    Log "$($User.UserPrincipalName);*ERRO*: licenca phoneSystem NÃO aplicada após $SLEEPT seg"
                    $NOK = $false
                }
                   
            }
        }    
    } else {
        Log "$($User.UserPrincipalName);*ATENÇÃO*: Licença já configurada previamente"
        ##Usuário já possuia Licença
    }






    #Adicionando o usuário ao grupo
    try{

        New-MgGroupMember -GroupId $GUID -DirectoryObjectId $User.Id
    
    } catch {
        
        Log "$($User.UserPrincipalName);*ERRO*: erro ao adicionar ao grupo: $GUID"
        exit 1
    }






    #Atribuindo o Ramal ao usuário
    $cod = Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumber $Ramal -PhoneNumberType DirectRouting
    if ($cod -ccontains "has already been assigned to another user") {
        Start-Sleep 30
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq '$Ramal'"
        $status_ramal =  "$($OnlineRamal.Alias)"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"
        Log "$($User.UserPrincipalName);*ATENCAO*: previamente configurado o numero $Ramal para chave $status_ramal" 
        exit 1
    }

    Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -EnterpriseVoiceEnabled $true

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
        Grant-CsOnlineVoiceRoutingPolicy -Identity $User.UserPrincipalName -PolicyName $vrota 2>D:\Util\PhoneSystem\logs\err_CsOnlineVoiceRoutingPolicy.txt
    }
    catch {
        Log "$($User.UserPrincipalName);*ERRO*: OnlineVoiceRoutingPolicy $vrota"
        exit 1
    }

    Log "$($User.UserPrincipalName);configurado: OnlineVoiceRoutingPolicy $vrota"
    Start-Sleep 10 
    
    try {   
        Grant-CsTenantDialPlan -Identity $User.UserPrincipalName -PolicyName $plano 2>D:\Util\PhoneSystem\logs\err_CsTenantDialPlan.txt
    }
    catch {
        Log "$($User.UserPrincipalName);*ERRO*: TenantDialPlan $plano"
        exit 1
    }
            Log "$($User.UserPrincipalName);configurado: TenantDialPlan $plano"
          

}


function RemoveRamal($User, $TeamsUser){
    
    #Verificando se o usuário tem linha configurada

    if($null -eq $TeamsUser.LineURI){
        
        Log "$($User.UserPrincipalName);*ERRO*: Usuário não tem linha configurada"
        exit 1

    } 

    #Removendo o número de telefone
    try{
        
        Remove-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -RemoveAll -ErrorAction Stop
    } catch {
        
        Log "$($User.UserPrincipalName);*ERRO*: Erro ao tentar desconfigurar Linha"
        exit 1
    }

    #Capturando as Licenças do usuário
    $UserLicenses = Get-MgUserLicenseDetail -UserId $User.Id

    #Verifica se o usuário possui a licença
    if($UserLicenses.SkuPartNumber.Contains($SKU)){
        
        #Resgata o Id da Licença
        foreach($License in $UserLicenses){
            if($License.SkuPartNumber -eq $SKU){
                $MCOEVSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq "MCOEV"
            }
        }
       
       #try{
            #Removendo a licença do usuário
            Set-MgUserLicense -UserId $User.Id -AddLicense @() -RemoveLicenses @{SkuId = $MCOEVSku.SkuId} -ErrorAction Stop
       #} catch {
           # Log "$($User.UserPrincipalName);*ERRO*: Erro ao remover a licença do usuário"
            #exit 1
       #}

    } else {
    
        Log "$($User.UserPrincipalName);*ATENÇÃO*: Não foi possível remover a licença $SKU. Não localizada"
    }


    #Remover usuário do Grupo
    try{
        Remove-MgGroupMemberByRef -GroupId $GUID -DirectoryObjectId $User.Id #-ErrorAction Stop
    } catch {
        Log "$($User.UserPrincipalName);*ERRO*: Erro ao remover usuário do grupo $GUID"
        exit 1
    }

}
### Fim das Funções ###


### Inicio do Processamento ####


#Capturando o usuário no AD
try{
    $User = Get-MgUser -Filter "displayName eq '$Chave'"
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

