# Incluir/excluir membros de uma caixa de email compartilhada (shared mail box)

# sharedmailbox_gerenciar_membros.ps1 -acao [incluir|excluir] -chave_solicitante chave_solicitante -email_caixa email@caixa -lista_membros email@membro 


Param(
��� [Parameter(Mandatory=$true)] [string] $acao,
��� [Parameter(Mandatory=$true)] [string] $chave_solicitante,
��� [Parameter(Mandatory=$true)] [string] $email_caixa,
��� [Parameter(Mandatory=$true)] [array]  $lista_membros,
    [Parameter(Mandatory=$true)] [string] $chamado

��� 
)

# TENANT PRODU��O

$username = "SAMSAZU@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

# TENANT TESTE

#$username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
#$PlainPassword="Ror66406"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
�
$SLEEP = 60
$logFile = "d:\Util\Outlook\logs\"+(Get-Date).ToString('yyyyMMdd')+"_sharedmail_gerenciar_membros.csv"


function Log([string]$message){

    $datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    if ( $message -like '*ERRO*') {
        Write-Host "$message" -ForegroundColor Red
    }
    else { 
        if ( $message -like '*ATENCAO*') {
            Write-Host "$message" -ForegroundColor Yellow
        }
        else {
            Write-Host "$message" -ForegroundColor Green
        }
    }
    
     Add-Content -Path $logFile -Value "$datetime;$chamado;$message" -ErrorAction Ignore
         
}


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


# RECUPERA EMAIL DO SOLICITANTE A APARTIR DA CHAVE

try {
  $userAD = Get-ADUser -Identity $chave_solicitante
  $email_solicitante = $userAD.UserPrincipalName
  
}
catch {
    Log "ERRO: chave do solicitante $chave_solicitante n�o encontrada no AD interno"
    exit 1
    Disconnect-ExchangeOnline -Confirm:$false
}

$shared = Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $email_caixa -ErrorAction SilentlyContinue
if (!$shared) {
    Write-Host "ERRO: Caixa compartilhada $email_caixa inexistente."
    exit 1
    Disconnect-ExchangeOnline -Confirm:$false
}
else {
  $solicitanteEhMembro = $shared | Get-MailboxPermission -User $email_solicitante -ErrorAction SilentlyContinue
  if (!$solicitanteEhMembro) {
    log "ERRO: O usuario solicitante $chave_solicitante n�o � membro da caixa compartilhada $email_caixa."
    log "Segue a lista de membros que podem abrir esta solicita��o no ServiceNow:"
    $lista = Get-MailboxPermission -Identity $email_caixa | where {$_.User -like "*petrobras*"} | select User
    $lista = $lista.User -join "`n"
    log "$lista"
    exit 1
    Disconnect-ExchangeOnline -Confirm:$false
    
  }
  else {
    $erro_parcial = $false
    foreach ($email_membro in $lista_membros) {
      $membroTemCorreio = Get-Mailbox -Identity $email_membro -ErrorAction SilentlyContinue
      if (!$membroTemCorreio) {
        log "ERRO: Usu�rio $email_membro inexistente ou n�o tem caixa de correio no Outlook."
        $erro_parcial = $true
        
      }
      else {
        if ( $acao -eq "incluir") {
          Add-MailboxPermission -Identity $email_caixa -User $email_membro -AccessRights FullAccess -InheritanceType All | Out-Null
          Add-RecipientPermission -Identity $email_caixa -Trustee $email_membro -AccessRights SendAs -Confirm:$false | Out-Null
          log "SUCESSO: Usu�rio $email_membro inclu�do como membro da caixa compartilhada $email_caixa."
        }
        else {
          if ( $acao -eq "excluir" ) {
            $membroTemAcesso = $shared | Get-MailboxPermission -User $email_membro -ErrorAction SilentlyContinue
            if (!$membroTemAcesso) {
              Write-Host "Usu�rio $email_membro n�o � membro da caixa compartilhada $email_caixa."
            }
            else {
              Remove-MailboxPermission -Identity $email_caixa -User $email_membro -AccessRights FullAccess -InheritanceType All -Confirm:$false | Out-Null
              Remove-RecipientPermission -Identity $email_caixa -Trustee $email_membro -AccessRights SendAs -Confirm:$false | Out-Null
              log "SUCESSO: Usu�rio $email_membro removido dos membros da caixa compartilhada $email_caixa."
            }
          }
          else {
            Write-Host "ERRO: A��o inv�lida"
            exit 1
            Disconnect-ExchangeOnline -Confirm:$false
           
          }
        }
      }
    } 
  }
}

if ($erro_parcial) {
  exit 1
  Disconnect-ExchangeOnline -Confirm:$false
}
else {
  Disconnect-ExchangeOnline -Confirm:$false
}