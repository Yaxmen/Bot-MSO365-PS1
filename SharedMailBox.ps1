Param(
    [Parameter(Mandatory=$true)] [string] $AffectedUser,
    [Parameter(Mandatory=$true)] [string] $acao,
    [Parameter(Mandatory=$true)] [string] $Name,
    [Parameter(Mandatory=$true)] [string] $NewName,
    [Parameter(Mandatory=$true)] [string] $Mail,
    [Parameter(Mandatory=$true)] [string] $NewMail
)

 

#------------------------------------------------------------------------------------------#
# Este script tem a finalidade de atender as solicitações de "Alteração de Nome e Email de #
# caixa de correio compartilhada"                                                          #
# Ele recebe por parametro o Nome e Email atual da caixa e o novo nome e novo email da     #
# caixa e executa o seguinte:                                                              #
#   1) Verifica se o Usuário que abriu a demanda existe no exchange                        #
#   2) Verifica se a caixa de correio existe                                               #
#   3) Verifica se o usuário que abriu o chamado é membro da caixa de correio              #
#   4) Verifica qual é a demanda                                                           #
#   5) Caso necessário verifica se um novo email esta disponível                           #
#   6) Executa as alterações identificadas                                                 #
#------------------------------------------------------------------------------------------#


$username = "samsazu@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop



#$username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
#$PlainPassword="Ror66406"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword

$logFile = "d:\Util\Outlook\logs\"+(Get-Date).ToString('yyyyMMdd')+"_sharedMailBox_alterar_log.csv"

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
 

$SLEEP = 60

try {
    Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -InformationAction SilentlyContinue | Out-Null 
}
catch {
    
    Start-Sleep $SLEEP
    try {
        Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -InformationAction SilentlyContinue | Out-Null
       
    }
    catch {
        
        Start-Sleep $SLEEP
        try {
            Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -InformationAction SilentlyContinue | Out-Null
            
        }
        catch {    
            Log "ERRO: ao conectar Exchange"
            exit 1
        }

    }
}


##################### Funções de Processamento ########################


function RenameMailBox($Mailbox, $Name){

    try {   
        
        # Altera o Nome
        Set-MailBox -Identity $Mailbox.Alias -DisplayName $Name -ErrorAction Stop
                
        # Mensagem de Sucesso
        Log "Caixa de Correio $($Mailbox.DisplayName) teve seu Nome alterado para $Name com Sucesso"

    } catch {
        Log "*** ERRO ao alterar o nome e email da caixa de correio ***"
        exit 1
    }
}

function ChangeMailFromMailBox($Mailbox, $Mail){

    # Captura uma caixa de correio com o novo email
    $MailNameAlreadyExists = Get-Mailbox -Identity $Mail -ErrorAction SilentlyContinue

    # Verifica se o novo Email já está em uso
    if($MailNameAlreadyExists){

        # Saída de erro: Email já em uso
        Log "*** ERRO o Email $Mail já esta em uso ***"
        exit 1

    } else {
        
        try {
        
            # Altera o E-mail
            Set-MailBox -Identity $Mailbox.Alias -WindowsEmailAddress $Mail -ErrorAction Stop
            Log "SUCESSO: Email alterado para $Mail"

        } catch {
            
            Log "*** ERRO ao alterar o nome e email da caixa de correio ***"
            exit 1
        }
    }
}

 




############################ Processamento ############################


# Capturando a Mailbox do Usuário que abriu o chamado
$User = Get-Mailbox -Identity $AffectedUser -ErrorAction SilentlyContinue

# Validando se o usuário foi localizado
if($User) {

    # Capturando a caixa de correio que vai sofrer as alterações
    $SharedMailBox = Get-Mailbox -Identity $Mail -ErrorAction SilentlyContinue

    # Validando se a caixa de correio foi localizada
    if($SharedMailBox){

        # Resgatando os membros da caixa de correio compartilhada
        $SharedMailBoxMembers = Get-MailBoxPermission -Identity $SharedMailbox.Alias

        # Verificando se o Usuário que abriu o chamado é membro da Caixa de correio a ser alterada
        if($SharedMailBoxMembers.User.Contains($User.UserPrincipalName)){

            

            switch($acao){
                
                "CL_NOME" {

                    RenameMailbox -Mailbox $SharedMailBox -Name $NewName
                
                }

                "CL_EMAIL"{

                    ChangeMailFromMailBox -Mailbox $SharedMailBox -Mail $NewMail
                
                }

                "CL_NOME_EMAIL"{
                    
                    RenameMailbox -Mailbox $SharedMailBox -Name $NewName
                    ChangeMailFromMailBox -Mailbox $SharedMailBox -Mail $NewMail
                
                }
            
            }


        } else {

            # Erro caso o usuário afetado não seja membro da caixa de correio compartilhada a ser alterada
            
            Log "ERRO: O usuário afetado $($User.DisplayName) não é membro da Caixa de correio compartilhada $($SharedMailBox.DisplayName) `nEssa ação pode ser realizada pelos seguintes usuários: "

            # Removendo 'NT AUTHORITY\SELF'da lista de usuários
            $NewSharedMailboxMembersArray = [System.Collections.ArrayList]::new($SharedMailBoxMembers.User)

            $NewSharedMailboxMembersArray.Remove("NT AUTHORITY\SELF")
            $NewSharedMailboxMembersArray.Remove("NT AUTHORITY\SELF")

            foreach($Name in $NewSharedMailboxMembersArray){

                Write-Host "- $Name"
            
            }
        }

    } else {
        
        # Erro caso a caixa de correio não seja encontrada no Exchange 

        Log "*** ERRO: a Caixa de correio $Mail não foi localizada no Exchange ***"
        exit 1
    }

} else {

    # Erro caso o usuário afetado não seja encontrado no Exchange

    Log "*** Erro o Usuário Afetado: $User não foi localizado no Exchange ***"
    exit 1

}
