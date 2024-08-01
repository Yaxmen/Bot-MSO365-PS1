
# Lista membros de uma caixa de email compartilhada (shared mail box)

# sharedmail_consultar_membros.ps1 -email_caixa cc-atividade@petrobras.com.br


Param (

    [Parameter( Mandatory=$false)][String]$email_caixa
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$username = "SAMSAZU@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

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


$shared = Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $email_caixa -ErrorAction SilentlyContinue
if (!$shared)
{
    Write-Host "ERRO: Caixa compartilhada $email_caixa inexistente."
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}
else {

  $membros = Get-MailboxPermission -Identity $email_caixa -ErrorAction SilentlyContinue
  Write-Host "Membros da caixa compartilhada $email_caixa :"

  foreach ($email in $membros) {

    if ($email.User -match '@') {
      Write-Host $email.User
    }

  }
}
Disconnect-ExchangeOnline -Confirm:$false
