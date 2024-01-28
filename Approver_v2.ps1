lParam(
    [Parameter(Mandatory=$true)] [string] $Acao,
    [Parameter(Mandatory=$true)] [string] $SalaDeReuniao,
    [Parameter(Mandatory=$true)] [string] $Chaves
)

# Definindo credenciais de acesso a tenant
$username = "email"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop


#$username = "email"
#$PlainPassword="senha"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#s$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
 



$SLEEP = 60
$logFile = "d:\Util\Approver\logs\"+(Get-Date).ToString('yyyyMMdd')+"_approver_log.csv"

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
    Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false -InformationAction SilentlyContinue | Out-Null 
    Log "SUCESSO: ao conectar ExchangeOnline 1/3"
}
catch {
    
    Start-Sleep $SLEEP
    try {
        Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false -InformationAction SilentlyContinue | Out-Null
        Log "SUCESSO: ao conectar ExchangeOnline 2/3"
       
    }
    catch {
        
        Start-Sleep $SLEEP
        try {
            Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false -InformationAction SilentlyContinue | Out-Null
            Log "SUCESSO: ao conectar ExchangeOnline 3/3"
            
        }
        catch {    
            Log "ERRO: ao conectar Exchange"
            exit 1
        }

    }
}


function IncludeApprover($RoomMailBox, $Calendar, $Chaves){

    # Capturando o Resource Delegate
    $Resource = $Calendar.ResourceDelegates

    # Capturando o Book in Policy
    $BookInPolicy = $Calendar.BookInPolicy

    # Verificando se a sala aceita aprovadores
    if (($Calendar.AutomateProcessing -eq "AutoAccept") -and (-not($Calendar.AllBookInPolicy)) -and ($Calendar.AllRequestInPolicy)){

        Log "Sucesso: a sala aceita aprovadores"

    } else {

        Log "*ATENCAO*: Sala não aceita aprovadores - Setando para aceitar"

        try {
            # Setando a sala para aceitar apovadores

            Set-CalendarProcessing -Identity $RoomMailBox.Alias -AutomateProcessing AutoAccept -AllBookInPolicy $false -AllRequestInPolicy $true -Confirm:$false -ErrorAction Stop

            Log "Sucesso: Sala Setada para aceitar aprovadores"    
                 
        } catch {
            
            Log "Erro: Não foi possível configurar a sala para aceitar aprovadores"
        }
    }

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
            Log "*** Chave $Chave não localizada ***"
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
            $Folder = $RoomMailBox.alias + "@petrobrasbrteste.onmicrosoft.com:\Calendar"

            # Setando cada usuário como aprovador
            #Add-MailboxFolderPermission -Identity $Folder -User $AliasUser -AccessRights Editor -SharingPermissionFlags Delegate | Out-Null

            # Adicionando chave a lista de sucesso
            $SuccessList.Add($AliasUser) | Out-Null
        
        }
    }

    if($SuccessList.Count -lt 1){

        Log "*ERRO*: Todas as chaves informadas já estão cadastradas como aprovador ou nenhuma chave foi localizada no Outlook"
        exit 1
    
    }


    try {    
        
        # Definindo lista de aprovadores
        Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource -ErrorAction Stop -Confirm:$false | Out-Null

        # Defininfo Lista de pré-Aprovadores
        Set-CalendarProcessing -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy -Confirm:$false -ErrorAction Stop | Out-Null

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


function ExcludeApprover($RoomMailBox, $Calendar, $Chaves){

    # Capturando o Resource Delegate
    $Resource = $Calendar.ResourceDelegates

    # Capturando o Book in Policy
    $BookInPolicy = $Calendar.BookInPolicy

    # Criando um Array com as Chaves
    $ArrayChaves = $Chaves.Split(",")
    
    # Criando a folder
    $folder = $RoomMailBox.DisplayName + ":\Calendar"

    #Criando os arrays de sucesso e falha
    $SuccessList = [System.Collections.ArrayList]::new()
    $FailedList = [System.Collections.ArrayList]::new()

    # Percorrendo o array de chaves para remover o acesso de aprovador
    foreach($Chave in $ArrayChaves){

        # Capturando a mailbox da chave
        try{
            $MailUser = Get-Mailbox -Identity $Chave -ErrorAction Stop
        } catch{
            Log "*ERRO*: Chave $chave não localizada"
        }

        # Verificando se o Usuário está cadastrado como aprovador
        if(-not ($Resource.Contains($MailUser.Name)) -and -not ($BookInPolicy.Contains($MailUser.LegacyExchangeDN))){
            
            # Adiciona o usuário a lista de falhas
            $FailedList.Add($MailUser.Alias) | Out-Null
            

        } else {

            # Remove a chave da lista de aprovadores
            $Resource.Remove($MailUser.Name)

            # Remove a chave da lista do LegacyExchange
            $BookInPolicy.Remove($MailUser.LegacyExchangeDN)
            
            try {
            
                # Removendo as permissões de calendário
                Remove-MailboxFolderpermission -Identity $folder -User $MailUser.Alias -Confirm:$false -ErrorAction Stop | Out-Null
            
            } catch {
                
                Log "Erro ao Remover Usuário do Calendário"
                exit 1
            }

            #Adicionando Chave a lista de sucesso
            $SuccessList.Add($MailUser.Alias) | Out-Null
            

        }
    }

    # Verificando se houve algum sucesso
    if ($SuccessList.Count -lt 1){
        
        Log "*ERRO*: ao remover aprovadores. Nenhuma das chaves informadas foram encontradas como aprovador ou nenhuma chave foi localizada no outlook"
        exit 1

    }

    try{
        
        # Redefinindo a lista de aprovadores da sala de reunião
        Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource -ErrorAction Stop -Confirm:$false | Out-Null

        # Redefinindo lista de pré-aprovadores da sala
        Set-CalendarProcessing -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy -ErrorAction Stop -Confirm:$false | Out-Null

        if($FailedList.Count -gt 0){
            
            foreach($User in $FailedList){
                Log "*ATENCAO*: Chave não localizada como aprovador da sala - $User"
            }
        }

        Log "*** SUCESSO: Chaves removidas com sucesso ***"
    
    } catch {
    
        Log "*** ERRO: Falha na remoção das Chaves ***"
    }

    

    # Redefinindo as permissões de calendário
    try{
 
       Set-MailboxFolderPermission -Identity $folder -User Default -AccessRights AvailabilityOnly -ErrorAction Stop -Confirm:$false | Out-Null

       Log "*** SUCESSO: Calendário atualizado ***"

    } catch {
       Log "*** ERRO: erro ao atualizar o calendário ***"
    }
}


# Capturando a Mailbox da Sala de Reunião
$RoomMailBox = Get-Mailbox -Identity $SalaDeReuniao -ErrorAction SilentlyContinue

# Conferindo se a MailBox existe
if($null -eq $RoomMailBox){

    Log "*** ERRO $SalaDeReunião não localizada"
    exit 1

} else {

    # Conferindo se a Sala de Reunião é privativa
    if($RoomMailBox.DisplayName -match "Privativa"){

        try {

            # Resgatando o calendário
            $Calendar = Get-CalendarProcessing -Identity $RoomMailBox.Alias -ErrorAction Stop

        }
        catch {
            Log "*** ERRO ao capturar as informações do Calendário ***"
            exit 1
        }

        # Se a ação for Incluir
        if($Acao -eq "CA_INCLUIR_3055"){

            IncludeApprover -RoomMailBox $RoomMailBox -Calendar $Calendar -Chaves $Chaves

        }

        # Se a ação for Excluir
        if($Acao -eq "CA_EXCLUIR_3055") {
            
            ExcludeApprover -RoomMailBox $RoomMailBox -Calendar $Calendar -Chaves $Chaves

        }

    } else{
        Log "*** ERRO $($RoomMailBox.DisplayName) não classificada como privativa***"
        exit 1
    }
}
