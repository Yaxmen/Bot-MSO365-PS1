Param (

    [Parameter( Mandatory=$true)] [String]$Action,
    [Parameter( Mandatory=$true)] [String]$Office,
    [Parameter( Mandatory=$true)] [String]$Subsidiaria,
    [Parameter( Mandatory=$true)] [String]$Building,
    [Parameter( Mandatory=$true)] [String]$Floor, 
    [Parameter( Mandatory=$true)] [String]$IdRoom, 
    [Parameter( Mandatory=$true)] [String]$Capacity,
    [Parameter( Mandatory=$true)] [String]$Address,
    [Parameter( Mandatory=$true)] [String]$RoomType,
    [Parameter( Mandatory=$true)] [String]$Chaves, 
    [Parameter( Mandatory=$true)] [String]$RoomName,
    [Parameter( Mandatory=$true)] [String]$RoomAlias,
    [Parameter( Mandatory=$true)] [String]$ChangeOption,
    [Parameter( Mandatory=$true)] [String]$NewDisplayName,
    [Parameter( Mandatory=$true)] [String]$HiddenOption
   

)


Import-Module Az.Resources

# Constantes e conexão com o Graph (PROD)
$ClientId = "b88e924b-530d-497c-8b16-b54456064e5f"
$TenantId = "5b6f6241-9a57-4be4-8e50-1dfa72e79a57"
$key = (1..16)
$Secret = Get-Content "D:\Password\GraphAppPassword.txt" | ConvertTo-SecureString -Key $key -ErrorAction Stop
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $Secret
 
# Constantes de conexão com o Graph (DEV)
#$ClientId = "60c5c5b3-84dd-4519-8ad7-58a34f8d39b4"
#$TenantId = "6af8f826-d4c2-47de-9f6d-c04908aa4e88"
#$Key = Get-Content "D:\Util\KeyFile\KeyFile.key"
#$Secret = Get-Content "D:\Util\KeyFile\Password.txt" | ConvertTo-SecureString -Key $Key
#$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $Secret

$Sleep = 40

# Constantes
$Domain_UPN = "petrobrasbr.onmicrosoft.com" #"petrobrasbrteste.onmicrosoft.com"
$Domain_EMAIL = "petrobras.com.br" #"petrobrasteste.petrobras.com.br"
$CompanyName = "PETROBRAS"
$Department = "TIC/OI/SSH/SC"
$UsageLocation = "BR"
$PreferredLanguage = "pt-BR"

$Group = "GN_DU_SALAS-REUNIAO" #"Grupo Teste - Room"
$GUID =  "01d04540-8f23-4843-baae-b5fba0ea7888" #"d2c99326-f0e3-4452-a5a6-33c8363c0fd6"


$logFile = "d:\Util\rooms\"+(Get-Date).ToString('yyyy')+"_xRoomlog.csv"

 


# Definindo credenciais de acesso a tenant
$username = "samsazu@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

#$Username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
#$PlainPass = "Ror66406"
#$SecurePass = $PlainPass | ConvertTo-SecureString -AsPlainText -Force
#$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePass



#Função de Log
function Log([string]$message){


    $datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    if ( $message.Contains('ERRO')) {
        Write-Host "$datetime;$message" -ForegroundColor Red
    }
    else { 
        if ( $message.Contains('ATENCAO')) {
            Write-Host "$datetime;$message" -ForegroundColor Yellow
        }
        else {
            Write-Host "$datetime;$message" -ForegroundColor Green
        }
    }

    Add-Content -Path $logFile -Value "$datetime;$message"
}


#Conectando no MicrosoftGraph
try{
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential | Out-Null
    Log "SUCESSO: ao Conectar com o Graph 1/3"
} catch {
    Log "ATENCAO: ao conectar com o Graph 1/3"
    Start-Sleep $SLEEP
    try{
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential | Out-Null
        Log "SUCESSO: ao Conectar com o Graph 2/3"
    }catch{
        Log "ATENCAO: ao conectar com o Graph 2/3"
        Start-Sleep $SLEEP
        try {
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential | Out-Null
            Log "SUCESSO: ao Conectar com o Graph 3/3"
        } catch{
            Log "ERRO: ao conectar com o Graph 3/3"
            exit 1
        }
    }
}


#Conectando com o ExchangeOnline
try{
    Connect-ExchangeOnline -Credential $Credential -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
    Log "SUCESSO: ao Conectar com o Exchange 1/3"
} catch {
    Log "ATENCAO: ao conectar com o Exchange 1/3"
    Start-Sleep $SLEEP
    try{
        Connect-ExchangeOnline -Credential $Credential -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
        Log "SUCESSO: ao Conectar com o Exchange 2/3"
    }catch{
        Log "ATENCAO: ao conectar com o Exchange 2/3"
        Start-Sleep $SLEEP
        try {
            Connect-ExchangeOnline -Credential $Credential -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
            Log "SUCESSO: ao Conectar com o Exchange 3/3"
        } catch{
            Log "ERRO: ao conectar com o Exchange 3/3"
            exit 1
        }
    }
}

 

####### Funções Auxiliares #########

function GetAliasFree(){

    # function para retornar primeiro alias sem uso

    for ($x=1; $x -le 9999; $x++){

        #formartar $x com 4 digitos
        $y = "{0:D4}" -f $x;$s = "SL"+"$y"
        if ($hash_alias[$s]) {
            continue
        }
        else {

            $ADUser = Get-User -Identity $s 2>$null

            if ($s -eq $($ADUser.Identity)) {
                #write-host "descartando $s"
                continue
            }
            else {   
                return $s
            }
        }
    }
}

function IncludeApprover($RoomMailBox, $Chaves){
    
    # Capturando o Calendário
    $Calendar = Get-CalendarProcessing -Identity $RoomMailBox.Alias -ErrorAction Stop

    # Capturando o Resource Delegate
    $Resource = $Calendar.ResourceDelegates

    # Capturando o Book in Policy
    $BookInPolicy = $Calendar.BookInPolicy

    # Criando um array com as chaves
    $ArrayChaves = $Chaves.Split(",")
    
    # Criando Arrays de Sucesso e Falha

    $SuccessList = [System.Collections.ArrayList]::new()

    $FailedList = [System.Collections.ArrayList]::new()

    # Percorrendo todas as chaves para fazer a inclusão
    foreach($Chave in $ArrayChaves){

        # Captura a mailbox da chave
        try{ 
            $MailUser = Get-MailBox -Identity $Chave -ErrorAction Stop
        } catch {
            Log "*** ATENCAO: Chave $Chave não localizada ***"
            Continue
        }

        # Capturando o LegacyExchange e o Alias do usuario

        $LegacyExchange = $MailUser.LegacyExchangeDN

        $AliasUser = $MailUser.Alias

        # Verifica se a Chave ja esta cadastrada como Aprovador
        if(($Resource.Contains($MailUser.DisplayName)) -or ($BookInPolicy.Contains($LegacyExchange))) {
            
            # Adicionando Chave a lista de Falhas
            $FailedList.Add($AliasUser) | Out-Null

        } else {

            # Incluindo como aprovador da Sala pelo Alias
            $Resource.Add($AliasUser) | Out-Null

            # Incluindo como Aprovador da Sala pelo LegacyExchange
            $BookInPolicy.Add($LegacyExchange) | Out-Null

            # Definindo a Folder do calendário 
            $Folder = $RoomMailBox.DisplayName + ":\Calendar"

            # Setando cada usuário como aprovador
            Add-MailboxFolderPermission -Identity $Folder -User $AliasUser -AccessRights Editor -SharingPermissionFlags Delegate | Out-Null

            # Adicionando chave a lista de sucesso
            $SuccessList.Add($AliasUser) | Out-Null
        
        }
    }

    if($SuccessList.Count -lt 1){

        Log "*ATENCAO*: Todas as chaves informadas já estão cadastradas como aprovador ou nenhuma chave foi localizada no Outlook"
    
    }


    try {    
        
        # Definindo lista de aprovadores
        Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource -ErrorAction Stop | Out-Null

        # Defininfo Lista de pré-Aprovadores
        Set-CalendarProcessing -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy -ErrorAction Stop | Out-Null

        if($FailedList.Count -gt 0){

            foreach($User in $FailedList) {
                Log "*ATENCAO*: Chave já cadastrada como aprovador da Sala - $User"
            }
        }
        
        Log "*** SUCESSO: Lista de aprovadores atualizada"
    }
    catch {

        Log "*** ERRO: Erro ao redefinir a lista de aprovadores"
        Exit 1
    }
}

function HandleCapacity($Capacity){

    if($Capacity.Length -eq 6){
        
        $NewCapacity = $Capacity[5]
        return $NewCapacity

    } elseif($Capacity.Length -eq 7){

        $NewCapacity = $Capacity[5] + $Capacity[6]
        return $NewCapacity
    }

}

function HandleInput([String]$Input){
    
    $WithoutCL = $Input.Replace("CL_","")

    $WithoutUnderLine = $WithoutCL.Replace("_"," ")

    Write-Host $WithoutUnderLine

    return $WithoutUnderLine
}


function HandleFloor($Floor){

    # Verifica se o Floor é de um digito
    if($Floor.Length -eq 4){
        
        $FinalFloor = $Floor[3]
        return $FinalFloor
    }

    # Verifica se o Floor é de dois Digitos
    if($Floor.Length -eq 7){
        
        $FinalFloor = $Floor[5]+$Floor[6]
        return $FinalFloor
    }

    # Verifica se o Floor é EMBAS
    if($Floor -match "EMBAS"){
        
        $FinalFloor = $Floor.Replace("CL_","")

        $FinalFloor = $FinalFloor.Replace("_", " ")

        return $FinalFloor
    }

    # Verifica se o Floor é P
    if($Floor -match "P"){
        
        $FinalFloor = $Floor.Replace("CL_","")

        $FinalFloor = $FinalFloor.Replace("_","")

        return $FinalFloor
    }

    # Verifica se o Floor é SS ou Térreo
    if($Floor -match "SS" -or $Floor -match "TERREO"){

        $FinalFloor = $Floor.Replace("CL_0_","")

        return $FinalFloor
    
    }

    # Verifica se o Floor é "Não Encontrado"
    if($Floor -match "NAO_ENCONTRADO"){
        
        $FinalFloor = $Floor.Replace("CL_Z_", "")

        $FinalFloor = $FinalFloor.Replace("_", " ")

        return $FinalFloor
    
    }

}

function GetDisplayName(){

    $NewBuilding = HandleInput -Input $Building
    $NewFloor = HandleFloor -Floor $Floor

    $FirstPart = $Office

    if($Building -eq "Não Aplicavel"){
        
        $SecondPart = ""
    } else {

        $SecondPart = $NewBuilding
    }

    $ThirdPart = $NewFloor

    if($NewFloor -eq "Terreo"){
        
        $FourthPart = ""
    } elseif ($NewFloor -eq "SS"){
        
        $FourthPart = ""
    } else {
        
        $FourthPart = "Andar "
    }

    $FifthPart = $IdRoom

    if($RoomType -eq "Sala Privativa"){
        $SixthPart = "(Privativa)"
    } else {
        $SixthPart = ""
    }

    $DisplayName = $FirstPart + " " + $SecondPart + " " + $ThirdPart + " " + $FourthPart + $FifthPart + " " + $SixthPart

    Return $DisplayName

}

####### Fim das Auxiliares ########



function CreateRoom(){

 
    #### CRIAR HASH COM SALAS E OFFICES ######
    $rooms = Get-Mailbox -Filter "(RecipientTypeDetails -eq 'RoomMailBox' -and Alias -like 'SL*')" -ResultSize unlimited


    $hash_alias = @{}
    $hash_office = @{}

    foreach ($r in $rooms) {

        $alias = $r.Alias
        $hash_alias.Add($alias,$alias)
        $build = $r.Office
        if ( -not $hash_office[$build]) {
            $hash_office.Add($build,$alias)
        }
    }

    #Resgatando o Alias a ser utilizado e definindo o UPN e Email
    try{
        $Alias = GetAliasFree
        Write-Output $Alias
    } catch {

        Log "ERRO: Não foi possível capturar um Alias"
        exit 1
    }

    $UPN = $alias+"@"+$Domain_UPN
    $EMAIL = $alias+"@"+$Domain_EMAIL

    # Gerando Password
    $Password = [System.Web.Security.Membership]::GeneratePassword(10, 2)
    $PasswordProfile = New-Object -TypeName Microsoft.Azure.PowerShell.Cmdlets.Resources.MSGraph.Models.ApiV10.MicrosoftGraphPasswordProfile 
    $PasswordProfile.Password = $Password

    # Capturando as Licenças da Tenant
    $TenantPlans = Get-MgSubscribedSku

    # Capturando a quantidade de Licenças Disponíveis
    foreach($Plan in $TenantPlans){

        if($Plan.SkuPartNumber -eq "STANDARDPACK"){
            $E1Sku = Get-MgSubscribedSku -All | Where SkupartNumber -eq "STANDARDPACK"
            $AvailableE1 = $Plan.PrepaidUnits.Enabled

        } elseif ($Plan.SkuPartNumber -eq "ENTERPRISEPACK") {
            $E3Sku = Get-MgSubscribedSku -All | Where SkupartNumber -eq "ENTERPRISEPACK"
            $AvailableE3 = $Plan.PrepaidUnits.Enabled

        } elseif ($Plan.SkupartNumber -eq "MEETING_ROOM"){
            $MEETINGSku = Get-MgSubscribedSku -All | Where SkupartNumber -eq "MEETING_ROOM"
            $AvailableMTR = $Plan.PrepaidUnits.Enabled
        }
    }

    # Checando se existem licenças suficientes
    if($AvailableE1 -lt 1 -and $AvailableE3 -lt 1){

        Log "*ERRO*: Quantidade de licenças disponíveis não são o suficiente para criar conta Azure"
        exit 1

    }

    # Validando qual a licença que vai ser utilizada
    if($AvailableE3 -gt $AvailableE1){

        $SKUTemplate = $E3Sku.SkuId
    } else {
        $SKUTemplate = $E1Sku.SkuId
    }

    #Gerando o DisplayName
    $DisplayName = GetDisplayName

    # Criando conta no Azure AD
    try{
        New-MgUser -UserPrincipalName $UPN -AccountEnabled -CompanyName $CompanyName -Department $Department -DisplayName $DisplayName -GivenName $Office -MailNickName $alias -PreferredLanguage $PreferredLanguage -Surname $Office -UsageLocation $UsageLocation -PasswordProfile $PasswordProfile -StreetAddress $Address -ErrorAction Stop | Out-Null
    } catch{
        Log "$Alias; *ERRO*: ao Criar conta no AzureAD: $DisplayName"
        exit 1
    }

    Start-Sleep $Sleep

    # Recuperando o Id da conta recém criada
    try{
        $RoomUser = Get-MgUser -UserId $UPN
        $UserObjectId = $RoomUser.Id
    } catch {
        Log "*ERRO*: ao recuperar o ID da conta criada"
        exit 1
    }

    # Adicionando o Usuário ao grupo
    try {
        New-MgGroupMember -GroupId $GUID -DirectoryObjectId $UserObjectId -ErrorAction Stop | Out-Null
    } catch {
        Log "*ERRO*: Ao adicionar o usuário ao grupo"
        exit 1
    }

    # Adicionando a Licença ao usuário
    try {
        Set-MgUserLicense -UserId $UserObjectId -AddLicense @{SkuId = $SKUTemplate} -RemoveLicenses @() -ErrorAction Stop | Out-Null
        Log "SUCESSO: Licença aplicada. Aguardando conversão para sala de reunião"
    } catch {
        Log "*ERRO*: ao Atribuir licença para a conta Azure"
        exit 1
    }

    



    ### Aguardando a criação de uma Mailbox no exchange após a atribuição da Licença ###


    $NOK = $true
    while($NOK){
        start-sleep $Sleep
        $UserMailbox = Get-Mailbox -Identity $alias -ErrorAction SilentlyContinue

        if($Null -eq $UserMailbox) {
        
            $Sleept = $Sleept + $Sleep
            
            if($Sleept -gt 119) {
                Log "$Alias; *ERRO*: Não foi possível localizar a caixa de correio no exchange"
                exit 1
            }
        } else {
            $NOK = $false
        }
    }

    #####################################################################################



    # Converter a sala pra Room
    try {
        Set-Mailbox -Identity $UserMailbox.alias -Type Room -ErrorAction Stop | Out-Null #-EmailAddresses smtp:$Email
    } catch {
        Log "*ERRO: Ao modificar o tipo da Mailbox para ROOM*"
        exit 1
    }


    Start-Sleep 60
    

    # Alterando configurações regionais
    try {
        Set-MailboxRegionalConfiguration -Identity $UserMailbox.alias -Language "pt-br" -TimeZone "E. South America Standard Time" -ErrorAction Stop | Out-Null
    } catch {
        Log "*ERRO*: ao configurar as alterações regionais"
        exit 1
    }
   
    # Alterando configurações de Calendário

    if($RoomType -eq "CL_SALA_PUBLICA"){
        
        try {
            Set-CalendarProcessing -Identity $UserMailbox.alias -AutomateProcessing "None" -AllBookInPolicy $true -AllRequestInPolicy $false -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -AddAdditionalResponse $false -ProcessExternalMeetingMessages $true -ErrorAction Stop
        } catch{
            Log "*ATENCAO*: Ao alterar as configurações de calendário"
        }

    } elseif ($RoomType -eq "CL_SALA_PRIVATIVA"){                try {
            Set-CalendarProcessing -Identity $UserMailbox.alias -AutomateProcessing "AutoAccept" -AllBookInPolicy $false -AllRequestInPolicy $true -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -AddAdditionalResponse $false -ProcessExternalMeetingMessages $true -ErrorAction Stop

            Start-Sleep 20

            IncludeApprover -Room $UserMailBox -Chaves $Chaves 

            # Modificando o nome da sala para privativa
            $NewRoomDisplayName = $UserMailBox.DisplayName + " (Privativa)"

            Set-Mailbox -Identity $UserMailbox.Alias -DisplayName $NewRoomDisplayName

        } catch{
            Log "*ATENCAO*: Ao alterar as configurações de calendário"
        }    }
    
    # Capturando o Office na Hash de Offices para copiar as propriedades
    if($hash_office[$office]){
        $Place = Get-Place -Identity $($hash_office[$office])
        try{
            Set-User -Identity $alias -StateOrProvince $Place.StateOrProvince -PostalCode $Place.PostalCode -City $Place.City -StreetAddress $Place.StreetAddress -Confirm:$false -ErrorAction Stop | Out-Null
        } catch {
            Log "*ERRO*: ao configurar informações de Local."
            exit 1
       }
    } else {
       Log "*ATENCAO*: ao configurar informações de Local. $Office não encontrado"
        
    }

        #Set-User -Identity $UserMailBox.Alias -StateOrProvince "São Paulo" -PostalCode "01234-123" -City "São Paulo" -StreetAddress "StreetAddress" -Confirm:$false -ErrorAction Stop | Out-Null

    # Realizando configurações de Set-Place

    # Convertendo o Andar para Int
    $Floor = HandleFloor -Floor $Floor

    # Tratando a entrada de Subsidiaria
    $NewSubsidiaria = $Subsidiaria.Replace("CL_", "")

    if(($Floor -match "EMBAS") -or ($Floor -match "P") -or ($Flor -match "SS") -or ($Floor -match "TERREO") -or ($Floor -match "NAO_ENCONTRADO")){

        try{

            Set-Place -Identity $UserMailbox.Alias -Building $NewSubsidiaria -Capacity $IntCapacity -DisplayDeviceName "TV" -ErrorAction Stop | Out-Null
        } catch {
            Log "*ERRO*: ao configurar o DisplayDeviceName"
            exit 1
        }

    } else {
        
        $IntFloor = [int]::Parse($Floor)

        try{

            Set-Place -Identity $UserMailbox.Alias -Building $NewSubsidiaria -Capacity $IntCapacity -DisplayDeviceName "TV" -Floor $IntFloor -ErrorAction Stop | Out-Null
        } catch {
            Log "*ERRO*: ao configurar o DisplayDeviceName"
            exit 1
        }
    }
    

    
       
    

    

    # Removendo a licença E3/E1 da sala
    try{
        Set-MgUserLicense -UserId $UserObjectId -RemoveLicenses @($SKUTemplate) -AddLicense @{} -ErrorAction Stop | Out-Null
    } catch {
        Log "*ERRO*: ao remover a licença da sala de reunião"
        exit 1
    }

   
}

function HandleTypeRoom($Room, $Option){

    # Resgatando o Calendar da sala
    try{
        $RoomCalendar = Get-CalendarProcessing -Identity $Room.Alias -ErrorAction Stop
    } catch{
        Log "*Erro*: Ao capturar o Calendário"
    }

    # Verificando o tipo de ação a ser tomado pela Função

    # Ação de tornar a Sala Privativa
    if($Option -eq "CL_SALA_PRIVATIVA"){

        # Verifica se a sala já é privativa
        if($RoomCalendar.AutomateProcessing -ne "AutoAccept"){

            #Setando a sala para privativa
            try{
                Set-CalendarProcessing -Identity $Room.Alias -AutomateProcessing "AutoAccept" -AllBookInPolicy $false -AllRequestInPolicy $true -ErrorAction Stop
            }catch{
                Log "*ERRO: Ao alterar o AutomateProcessing para AutoAccept"
            }

            #Adicionando aprovadores
            try{
                IncludeApprover -Room $Room -Chaves $Chaves
            }catch{
                Log "*ERRO*: Ao incluir aprovadores na Sala"
            }

        } else {
            Log "*ERRO*: Sala $($Room.Displayname) já esta classificada como privativa"
            exit 1
        }

        # Capturando o novo nome da Sala
        $NewRoomName = $Room.DisplayName + " (Privativa)"

        # Redefinindo o nome da Sala
        try{
            Set-Mailbox -Identity $Room.Alias -DisplayName $NewRoomName
        } catch {
            Log "*ERRO*: Ao redefinir o nome da sala"
            exit 1
        }

        Log "SUCESSO: Agora a sala $($Room.DisplayName) é uma Sala Privativa"
    }

    # AÃ§Ã£o para tornar a sala Pública
    if($Option -eq "Sala Publica"){

        # Verifica se a sala jÃ¡ Ã© PÃºblica
        if($RoomCalendar.AutomateProcessing -eq "AutoAccept"){

            try {
                Set-CalendarProcessing -Identity $Room.Alias -AutomateProcessing "None" -AllBookInPolicy $true -AllRequestInPolicy $false
            } catch {
                Log "*ERRO*: Ao redefinir o calendário da sala"
            }

        } else {
            Log "*ERRO*: Sala $($Room.DisplayName) já está configurada como Pública"
            exit 1
        }

        # Capturando o novo Nome da Sala
        $NewRoomName = $Room.DisplayName.replace(" (Privativa)", "")

        # Redefinindo o nome da Sala
        try{
            Set-Mailbox -Identity $Room.Alias -DisplayName $NewRoomName
            Log "Sucesso, a Sala $($Room.DisplayName) agora é uma sala pública"
        } catch{
            Log "*ERRO*: Ao redefinir o nome da sala"
            exit 1
        }
    }
}

function HandleCapacityRoom($Room, $Capacity){

    # Modificando a capacidade da Sala
    try{    
        Set-Place -Identity $Room.Alias -Capacity $IntCapacity
        Log "*SUCESSO*: Sala $($Room.Alias) teve a capacidade modificada para $Capacity"
    } catch {
        Log "*ERRO*: Ao redefinir a capacidade da Sala $($Room.Alias)"
    }
}

function HandleDisplayNameRoom($Room, $NewRoomName){

    # Alterando o Nome da Sala de Reunião
    try{
        Set-Mailbox -Identity $Room.Alias -DisplayName $NewRoomName
        Log "*SUCESSO*: Sala $($Room.DisplayName) foi renomeada para $NewRoomName"
    } catch {
        Log "*ERRO*: Ao renomear o Número/Identificador da sala $($Room.DisplayName)"
    }
}

function HandleHiddenRoom($Room, $Option){

    # Verificando qual Ã© o tipo de ação a ser tomada pela função

    # Ação de ocultar sala
    if($Option -eq "Ocultar"){

        # Verificando se a sala está ví­sivel para ser ocultado
        if(-not $($Room.HiddenFromAddressListsEnabled)){

            #Setando a caixa de correio como oculta
            try {
                Set-Mailbox -Identity $Room.Alias -HiddenFromAddressListsEnabled $true
                Log "*Sucesso*: Sala $($Room.DisplayName) Ocultada da lista de salas" 
            }
            catch {
                Log "*Erro*: ao Ocultar a sala no Outlook"
                exit 1
            }
        # Caso a sala já esteja oculta, lança erro
        } else {
            Log "*Erro*: ao Ocultar sala. Sala já esta oculta"
            exit 1
        }

    # Ação de desocultar a sala
    } elseif ($Option -eq "Desocultar"){

        # Verificando se a sala está oculta pra ser desocultada
        if($Room.HiddenFromAddressListsEnabled){

            # Setando a caixa de correio como Ví­sivel
            try {
                Set-Mailbox -Identity $Room.Alias -HiddenFromAddressListsEnabled $false
                Log "*SUCESSO*: Sala $($Room.DisplayName) Desocultada da lista de salas"
            }
            catch {
                Log "*Erro*: Ao desocultar sala no outlook"
                exit 1
            }
        } else {
            Log "*Erro*: Ao desocultar sala. Sala já está vísivel"
        }

    }
}

function DeleteRoom($Room){
   try{
        Remove-MgUser -UserId $Room.UserPrincipalName
        Log "$($Room.alias); SUCESSO: A Sala de Reunião $($Room.DisplayName) foi excluída."
   ; } catch {
        Log "$($Room.alias); *ERRO*: ao Deletar a sala do Exchange"
    }
}




###### Inicio do Processamento ######

# Convertendo Capacidades para Int
if($Capacity -ne "NA"){
    
    $NewCapacity = HandleCapacity -Capacity $Capacity
   
    $IntCapacity = [int]::Parse($NewCapacity)
}



# Verificando qual o tipo de ação que o script vai tomar
if ($action -eq "CL_Alterar" -or $action -eq "CL_Excluir"){

    # Capturando a Sala que vai ser Deletada ou modificada
    $Room = Get-Mailbox -Identity $RoomAlias -ErrorAction SilentlyContinue

    # Verificando se a Room existe
    if($Room){

        if($action -eq "CL_Alterar"){

            switch($ChangeOption){
                "CL_TIPO_SALA"{
                    HandleTypeRoom -Room $Room -Option $RoomType
                }
                "CL_CAPACIDADE"{
                    HandleCapacityRoom -Room $Room -Capacity $Capacity
                }
                "CL_NOME_SALA"{
                    HandleDisplayNameRoom -Room $Room -NewRoomName $NewDisplayName
                }
                "CL_OCULTA_DESOCULTA"{
                    HandleHiddenRoom -Room $Room -Option $HiddenOption
                }
            }


        } elseif ($action -eq "CL_Excluir"){
            DeleteRoom -Room $Room
        }

    } else {
        Log "$RoomAlias; *ERRO*: Sala $RoomAlias não foi localizada no exchange"
    }
    


} elseif ($action -eq "CL_Criar"){
    CreateRoom

    $DisplayName = GetDisplayName

    Write-Output "Sucesso na criação da Sala $DisplayName"
}



