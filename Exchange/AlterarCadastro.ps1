Param (

    [Parameter( Mandatory=$true)] [String]$Action, 
    [Parameter( Mandatory=$true)] [String]$Capacity,
    [Parameter( Mandatory=$true)] [String]$RoomType,
    [Parameter( Mandatory=$true)] [String]$Chaves, 
    [Parameter( Mandatory=$true)] [String]$RoomAlias,
    [Parameter( Mandatory=$true)] [String]$NewDisplayName,
    [Parameter( Mandatory=$true)] [String]$HiddenOption
   

)

 
# Constantes de conexão com o Graph ambiente DEV
<#$ClientId = "60c5c5b3-84dd-4519-8ad7-58a34f8d39b4"
$TenantId = "6af8f826-d4c2-47de-9f6d-c04908aa4e88"
$Key = Get-Content "D:\Util\KeyFile\KeyFile.key"
$Secret = Get-Content "D:\Util\KeyFile\Password.txt" | ConvertTo-SecureString -Key $Key
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $Secret #>

# Consantes de conexão com o Graph ambiente PROD
$ClientId = "b88e924b-530d-497c-8b16-b54456064e5f"
$TenantId = "5b6f6241-9a57-4be4-8e50-1dfa72e79a57"
$key = (1..16)
$Secret = Get-Content "D:\Password\GraphAppPassword.txt" | ConvertTo-SecureString -Key $key -ErrorAction Stop
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $Secret

$Sleep = 40


$logFile = "d:\Util\rooms\"+(Get-Date).ToString('yyyy')+"_xRoomlog.csv"

# Definindo credenciais de acesso a tenant

#Ambiente PROD
$username = "e-mail"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

#Ambiente DEV
#$Username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
#$PlainPass = "Ror66406"
#$SecurePass = $PlainPass | ConvertTo-SecureString -AsPlainText -Force
#$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePass



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
    Connect-ExchangeOnline -Credential $Credentials -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
    Log "SUCESSO: ao Conectar com o Exchange 1/3"
} catch {
    Log "ATENCAO: ao conectar com o Exchange 1/3"
    Start-Sleep $SLEEP
    try{
        Connect-ExchangeOnline -Credential $Credentials -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
        Log "SUCESSO: ao Conectar com o Exchange 2/3"
    }catch{
        Log "ATENCAO: ao conectar com o Exchange 2/3"
        Start-Sleep $SLEEP
        try {
            Connect-ExchangeOnline -Credential $Credentials -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
            Log "SUCESSO: ao Conectar com o Exchange 3/3"
        } catch{
            Log "ERRO: ao conectar com o Exchange 3/3"
            exit 1
        }
    }
}

 

# --------------------- Funções Auxiliares ------------------------ #


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
            Set-MailboxFolderPermission -Identity $Folder -User $AliasUser -AccessRights Editor -SharingPermissionFlags Delegate | Out-Null

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

# --------------------- Fim das Auxiliares ------------------------ #




# ------------------- Funções de Processamento -------------------- #

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
    if($Option -eq "CL_SALA_PUBLICA"){

        # Verifica se a sala já é Pública
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
        Set-Place -Identity $Room.Alias -Capacity $Capacity
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
    if($Option -eq "CL_OCULTAR_SALA"){

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
    } elseif ($Option -eq "CL_DESOCULTAR_SALA"){

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

# ------------------- Funções de Processamento -------------------- #





# ------------------- Inicio do Processamento --------------------- #

    # Convertendo Capacidades para Int
    if($Capacity -ne "NA"){
    
        $NewCapacity = HandleCapacity -Capacity $Capacity
   
        $IntCapacity = [int]::Parse($NewCapacity)
    }


    # Capturando a Sala que vai ser Deletada ou modificada
    $Room = Get-Mailbox -Identity $RoomAlias -ErrorAction SilentlyContinue

    # Verificando se a Room existe
    if($Room){

        switch($Action){
            "CL_TIPO_SALA"{
               HandleTypeRoom -Room $Room -Option $RoomType
            }
            "CL_CAPACIDADE"{
               HandleCapacityRoom -Room $Room -Capacity $IntCapacity
            }
            "CL_NOME_SALA"{
               HandleDisplayNameRoom -Room $Room -NewRoomName $NewDisplayName
            }
            "CL_OCULTA_DESOCULTA"{
               HandleHiddenRoom -Room $Room -Option $HiddenOption
            }
        }
    } else {
        Log "$RoomAlias; *ERRO*: Sala $RoomAlias não foi localizada no exchange"
		exit 1
    }
    
# ------------------- Final do Processamento ----------------------- #

