Param(
��� [Parameter(Mandatory=$true)] [string] $Acao,
��� [Parameter(Mandatory=$true)] [string] $SalaDeReuniao,
��� [Parameter(Mandatory=$true)] [string] $Chaves
)

# Definindo credenciais de acesso a tenant
#$username = "samsazu@petrobras.com.br"
#$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
#$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
#$Credentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop


$username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
$PlainPassword="Ror66406"
$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
�


$SLEEP = 60

try {
    Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -InformationAction SilentlyContinue | Out-Null 
}
catch {
    
    Start-Sleep $SLEEP
    try {
        Connect-ExchangeOnline -Credential $credential -InformationAction SilentlyContinue | Out-Null
       
    }
    catch {
        
        Start-Sleep $SLEEP
        try {
            Connect-ExchangeOnline -Credential $credential -InformationAction SilentlyContinue | Out-Null
            
        }
        catch {    
            Log "ERRO: ao conectar Exchange"
            exit 1
        }

    }
}


function IncludeApprover($RoomMailBox, $Calendar, $Chaves){

��� # Capturando o Resource Delegate
��� $Resource = $Calendar.ResourceDelegates

��� # Capturando o Book in Policy
��� $BookInPolicy = $Calendar.BookInPolicy

��� # Verificando se a sala aceita aprovadores
��� if (($Calendar.AutomateProcessing -eq "AutoAccept") -and (-not($Calendar.AllBookInPolicy)) -and ($Calendar.AllRequestInPolicy)){
��� } else {

������� try {
����������� # Setando a sala para aceitar apovadores
����������� Set-CalendarProcessing -Identity $RoomMailBox.Alias -AutomateProcessing AutoAccept -AllBookInPolicy $false -AllRequestInPolicy $true -ErrorAction Stop��������
������� } catch {
����������� Write-Output "*** ERRO: N�o foi poss�vel configurar a sala para aceitar aprovadores ***"
������� }
��� }

��� # Criando um array com as chaves
��� $ArrayChaves = $Chaves.Split(",")

��� # Percorrendo todas as chaves para fazer a inclus�o
��� foreach($Chave in $ArrayChaves){

��� # Captura a mailbox da chave
��� try{�
������� $MailUser = Get-MailBox -Identity $Chave -ErrorAction Stop
��� } catch {
������� Write-Host "*** Chave $Chave n�o localizada ***"
������� Continue
��� }

��� # Capturando o LegacyExchange e o Alias do usuario
��� $LegacyExchange = $MailUser.LegacyExchangeDN
��� $AliasUser = $MailUser.Alias

��� # Verifica se a Chave ja esta cadastrada como Aprovador
��� if(($Resource.Contains($MailUser.DisplayName)) -or ($BookInPolicy.Contains($LegacyExchange))) {
������� Write-Output "Chave j� esta cadastrada como aprovador: $AliasUser"
��� } else {

������� # Incluindo como aprovador da Sala pelo Alias
������� $Resource.Add($AliasUser) | Out-Null

������� # Incluindo como Aprovador da Sala pelo LegacyExchange
������� $BookInPolicy.Add($LegacyExchange) | Out-Null

������� # Definindo a Folder do calend�rio�
������� $Folder = $RoomMailBox.DisplayName + ":\Calendar"

������� # Setando cada usu�rio como aprovador
        
        Add-MailboxFolderPermission -Identity $Folder -User $AliasUser -AccessRights Editor -SharingPermissionFlags Delegate | Out-Null
        
        }
��� }

��� try {����
�����   
������� # Definindo lista de aprovadores
������� Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource -ErrorAction Stop | Out-Null

������� # Defininfo Lista de pr�-Aprovadores
������� Set-CalendarProcessing -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy -ErrorAction Stop | Out-Null
        
        Write-host "*** SUCESSO: Lista de aprovadores atualizada"
��� }
��� catch {

������� Write-Output "*** ERRO: Erro ao redefinir a lista de aprovadores"
��� }
}


function ExcludeApprover($RoomMailBox, $Calendar, $Chaves){

��� # Capturando o Resource Delegate
��� $Resource = $Calendar.ResourceDelegates

��� # Capturando o Book in Policy
��� $BookInPolicy = $Calendar.BookInPolicy

��� # Criando um Array com as Chaves
��� $ArrayChaves = $Chaves.Split(",")

��� # Percorrendo o array de chaves para remover o acesso de aprovador
��� foreach($Chave in $ArrayChaves){

������� # Capturando a mailbox da chave
������� try{
����������� $MailUser = Get-Mailbox -Identity $Chave -ErrorAction Stop
������� } catch{
����������� Write-Output "*** Chave $chave n�o localizada ***"
������� }

������� # Verificando se o Usu�rio est� cadastrado como aprovador
������� if(-not ($Resource.Contains($MailUser.Name)) -and -not ($BookInPolicy.Contains($MailUser.LegacyExchangeDN))){

����������� Write-Host "ATEN��O: Chave n�o localizada como aprovador da sala - $Chave"

������� } else {

����������� # Remove a chave da lista de aprovadores
����������� $Resource.Remove($MailUser.Name)

����������� # Remove a chave da lista do LegacyExchange
����������� $BookInPolicy.Remove($MailUser.LegacyExchangeDN)

������� }
��� }

    try{
���     
        
���     $folder = $RoomMailBox.DisplayName + ":\Calendar"
        
        # Redefinindo a lista de aprovadores da sala de reuni�o
���     Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource -ErrorAction Stop | Out-Null

���     # Redefinindo lista de pr�-aprovadores da sala
���     Set-CalendarProcessing -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy -ErrorAction Stop | Out-Null

        # Removendo as permiss�es de calend�rio
        Remove-MailboxFolderpermission -Identity $folder -User $MailUser.Alias -Force

        Write-Host "*** SUCESSO: Chave removida com sucesso ***"
    
    } catch {
    
        Write-host "*** ERRO: Falha na remo��o das Chaves ***"
    }

��� 

��� # Redefinindo as permiss�es de calend�rio
��� try{
�
������ Set-MailboxFolderPermission -Identity $folder -User Default -AccessRights AvailabilityOnly -ErrorAction Stop | Out-Null

       Write-Host "*** SUCESSO: Calend�rio atualizado ***"

��� } catch {
����   Write-Host "*** ERRO: erro ao atualizar o calend�rio ***"
��� }
}


# Capturando a Mailbox da Sala de Reuni�o
$RoomMailBox = Get-Mailbox -Identity $SalaDeReuniao -ErrorAction SilentlyContinue

# Conferindo se a MailBox existe
if($null -eq $RoomMailBox){

��� Write-Output "*** ERRO $SalaDeReuni�o n�o localizada"
} else {

��� # Conferindo se a Sala de Reuni�o � privativa
��� if($RoomMailBox.DisplayName -match "Privativa"){

������� try {

����������� # Resgatando o calend�rio
����������� $Calendar = Get-CalendarProcessing -Identity $RoomMailBox.Alias -ErrorAction Stop

������� }
������� catch {
����������� Write-Output "*** ERRO ao capturar as informa��es do Calend�rio ***"
������� }

������� # Se a a��o for Incluir
������� if($Acao -eq "Incluir"){

����������� IncludeApprover -RoomMailBox $RoomMailBox -Calendar $Calendar -Chaves $Chaves

������� }

������� # Se a a��o for Excluir
������� if($Acao -eq "Excluir") {
            
            ExcludeApprover -RoomMailBox $RoomMailBox -Calendar $Calendar -Chaves $Chaves

������� }

��� } else{
������� Write-Output "*** ERRO $SalaDeReuni�o n�o classificada como privativa***"
��� }
}