Param(
    [Parameter(Mandatory=$true)] [string] $Acao,
    [Parameter(Mandatory=$true)] [string] $SalaDeReuniao,
    [Parameter(Mandatory=$true)] [string] $Chaves
)

# Definindo credenciais de acesso a tenant
$username = "samsazu@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

$SLEEP = 60
$logFile = "d:\Util\Approver\logs\"+(Get-Date).ToString('yyyyMMdd')+"_approver_log.csv"

function Log([string]$message){
    $datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    Add-Content -Path $logFile -Value "$datetime;$message"
}

try {
    Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -InformationAction SilentlyContinue
    Log "SUCESSO: ao conectar ExchangeOnline 1/3"
} catch {
    Start-Sleep $SLEEP
    try {
        Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -InformationAction SilentlyContinue
        Log "SUCESSO: ao conectar ExchangeOnline 2/3"
    } catch {
        Start-Sleep $SLEEP
        try {
            Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -InformationAction SilentlyContinue
            Log "SUCESSO: ao conectar ExchangeOnline 3/3"
        } catch {    
            Log "ERRO: ao conectar Exchange"
            exit 1
        }
    }
}

function IncludeApprover($RoomMailBox, $Calendar, $Chaves){
    $Resource = $Calendar.ResourceDelegates
    $BookInPolicy = $Calendar.BookInPolicy

    if (($Calendar.AutomateProcessing -eq "AutoAccept") -and (-not($Calendar.AllBookInPolicy)) -and ($Calendar.AllRequestInPolicy)){
        Log "Sucesso: a sala aceita aprovadores"
    } else {
        Log "*ATENCAO*: Sala não aceita aprovadores - Setando para aceitar"
        try {
            Set-CalendarProcessing -Identity $RoomMailBox.Alias -AutomateProcessing AutoAccept -AllBookInPolicy $false -AllRequestInPolicy $true -Confirm:$false -ErrorAction Stop
            Log "Sucesso: Sala Setada para aceitar aprovadores"
        } catch {
            Log "Erro: Não foi possível configurar a sala para aceitar aprovadores"
        }
    }

    $ArrayChaves = $Chaves.Split(",")
    $SuccessList = [System.Collections.ArrayList]::new()
    $FailedList = [System.Collections.ArrayList]::new()

    foreach($Chave in $ArrayChaves){
        try{
            $MailUser = Get-MailBox -Identity $Chave -ErrorAction Stop
        } catch {
            Log "*** Chave $Chave não localizada ***"
            Continue
        }

        $LegacyExchange = $MailUser.LegacyExchangeDN
        $AliasUser = $MailUser.Alias

        if(($Resource.Contains($MailUser.DisplayName)) -or ($BookInPolicy.Contains($LegacyExchange))) {
            $FailedList.Add($AliasUser)
        } else {
            $Resource.Add($AliasUser)
            $BookInPolicy.Add($LegacyExchange)
            $Folder = $RoomMailBox.alias + "@petrobrasbrteste.onmicrosoft.com:\Calendar"
            $SuccessList.Add($AliasUser)
        }
    }

    if($SuccessList.Count -lt 1){
        Log "*ERRO*: Todas as chaves informadas já estão cadastradas como aprovador ou nenhuma chave foi localizada no Outlook"
        exit 1
    }

    try {    
        Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource -ErrorAction Stop -Confirm:$false
        Set-CalendarProcessing -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy -Confirm:$false -ErrorAction Stop
        if($FailedList.Count -gt 0){
            foreach($User in $FailedList) {
                Log "*ATENCAO*: Chave já cadastrada como aprovador da Sala - $User"
            }
        }
        Log "*** SUCESSO: Lista de aprovadores atualizada"
    } catch {
        Log "*** ERRO: Erro ao redefinir a lista de aprovadores"
        Exit 1
    }
}

function ExcludeApprover($RoomMailBox, $Calendar, $Chaves){
    $Resource = $Calendar.ResourceDelegates
    $BookInPolicy = $Calendar.BookInPolicy
    $ArrayChaves = $Chaves.Split(",")
    $folder = $RoomMailBox.DisplayName + ":\Calendar"
    $SuccessList = [System.Collections.ArrayList]::new()
    $FailedList = [System.Collections.ArrayList]::new()

    foreach($Chave in $ArrayChaves){
        try{
            $MailUser = Get-Mailbox -Identity $Chave -ErrorAction Stop
        } catch{
            Log "*ERRO*: Chave $chave não localizada"
        }

        if(-not ($Resource.Contains($MailUser.Name)) -and -not ($BookInPolicy.Contains($MailUser.LegacyExchangeDN))){
            $FailedList.Add($MailUser.Alias)
        } else {
            $Resource.Remove($MailUser.Name)
            $BookInPolicy.Remove($MailUser.LegacyExchangeDN)
            try {
                Remove-MailboxFolderpermission -Identity $folder -User $MailUser.Alias -Confirm:$false -ErrorAction Stop
            } catch {
                Log "Erro ao Remover Usuário do Calendário"
                exit 1
            }
            $SuccessList.Add($MailUser.Alias)
        }
    }

    if ($SuccessList.Count -lt 1){
        Log "*ERRO*: ao remover aprovadores. Nenhuma das chaves informadas foram encontradas como aprovador ou nenhuma chave foi localizada no outlook"
        exit 1
    }

    try{
        Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource -ErrorAction Stop -Confirm:$false
        Set-CalendarProcessing -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy -ErrorAction Stop -Confirm:$false
        if($FailedList.Count -gt 0){
            foreach($User in $FailedList){
                Log "*ATENCAO*: Chave não localizada como aprovador da sala - $User"
            }
        }
        Log "*** SUCESSO: Chaves removidas com sucesso ***"
    } catch {
        Log "*** ERRO: Falha na remoção das Chaves ***"
    }

    try{
       Set-MailboxFolderPermission -Identity $folder -User Default -AccessRights AvailabilityOnly -ErrorAction Stop -Confirm:$false
       Log "*** SUCESSO: Calendário atualizado ***"
    } catch {
       Log "*** ERRO: erro ao atualizar o calendário ***"
    }
}

$RoomMailBox = Get-Mailbox -Identity $SalaDeReuniao -ErrorAction SilentlyContinue

if($null -eq $RoomMailBox){
    Log "*** ERRO $SalaDeReunião não localizada"
} else {
    if($RoomMailBox.DisplayName -match "Privativa"){
        try {
            $Calendar = Get-CalendarProcessing -Identity $RoomMailBox.Alias -ErrorAction Stop
        } catch {
            Log "*** ERRO ao capturar as informações do Calendário ***"
        }

        if($Acao -eq "CA_INCLUIR_3055"){
            IncludeApprover -RoomMailBox $RoomMailBox -Calendar $Calendar -Chaves $Chaves
        }

        if($Acao -eq "CA_EXCLUIR_3055") {
            ExcludeApprover -RoomMailBox $RoomMailBox -Calendar $Calendar -Chaves $Chaves
        }
    } else{
        Log "*** ERRO $($RoomMailBox.DisplayName) não classificada como privativa***"
    }
}
