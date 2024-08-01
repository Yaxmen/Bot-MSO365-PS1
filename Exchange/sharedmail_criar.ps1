#.\sharedmail_criar.ps1 -nome 'nome da caixa' -email cc-nomedacaixa@petrobras.com.br -chave_solicitante chave_solicitante@petrobras.com.br -chamado chamado00001

Param(
    [Parameter(Mandatory=$true)] [string] $nome_caixa,
    [Parameter(Mandatory=$true)] [string] $email_caixa,
    [Parameter(Mandatory=$true)] [string] $chave_solicitante,
    [Parameter(Mandatory=$false)] [string] $chamado
   
)



# TENANT PRODUÇÃO

$username = "SAMSAZU@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

# TENANT TESTE

#$username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
#$PlainPassword="Ror66406"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
 
$SLEEP = 60
$logFile = "d:\Util\Outlook\logs\"+(Get-Date).ToString('yyyyMMdd')+"_sharedmail_criar.csv"


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

# RECUPERA EMAIL DO SOLICITANTE A APARTIR DA CHAVE

try {
  $userAD = Get-ADUser -Identity $chave_solicitante
  $email_solicitante = $userAD.UserPrincipalName
  
}
catch {
    Log "ERRO: chave do solicitante $chave_solicitante não encontrada no AD interno"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# VERIFICA SE chave_solicitante POSSUI CAIXA

$s1 = Get-Recipient -Identity $email_solicitante -ErrorAction SilentlyContinue

if (!$s1)
{
  Log "ERRO: solicitante $email_solicitante não encontrado no exchange"
  Disconnect-ExchangeOnline -Confirm:$false
  exit 1
}



# VERIFICA SE OPCOES ESCOLHIDAS JÁ EXISTEM

$n1 = Get-Recipient -Identity $nome_caixa -ErrorAction SilentlyContinue

if ($n1)
{
  Log "ERRO: O nome solicitado, ""$nome_caixa"", não está disponível."
  Disconnect-ExchangeOnline -Confirm:$false
  exit 1
}

$e1 = Get-Recipient -Identity $email_caixa -ErrorAction SilentlyContinue

if ($e1)
{
  Log "ERRO: O email solicitado, ""$email_caixa"", não está disponível."
  Disconnect-ExchangeOnline -Confirm:$false
  exit 1
}


$localpart = $email_caixa.split("@")
$alias = $localpart[0]

try {
  New-Mailbox -Shared -Name $nome_caixa -Alias $alias -PrimarySmtpAddress $email_caixa -ErrorAction Stop | Out-Null
}
catch {
  Log "ERRO: na criação da caixa compartilhada"
  Disconnect-ExchangeOnline -Confirm:$false
  exit 1
}

try {
  Set-Mailbox $email_caixa -MessageCopyForSentAsEnabled $True -MessageCopyForSendOnBehalfEnabled $True -ErrorAction Stop | Out-Null
}
catch {
  Log "ERRO: em Set-Mailbox "
  Disconnect-ExchangeOnline -Confirm:$false
  exit 1
}

try {
  Set-User $email_caixa -Notes $email_solicitante -Confirm:$false -ErrorAction Stop -WarningAction Stop | Out-Null
}
catch {
  Log "ERRO: em Set-User"
  Disconnect-ExchangeOnline -Confirm:$false
  exit 1
}

try {
  Add-MailboxPermission -Identity $email_caixa -User $email_solicitante -AccessRights FullAccess -InheritanceType All -ErrorAction Stop -WarningAction SilentlyContinue| Out-Null
}
catch {
  Log "ERRO: em Add-MailboxPermission"
  Disconnect-ExchangeOnline -Confirm:$false
  exit 1
}

try {
  Add-RecipientPermission -Identity $email_caixa -Trustee $email_solicitante -AccessRights SendAs -Confirm:$false -ErrorAction Stop| Out-Null
}
catch {
  Log "ERRO: em Add-RecipientPermission"
  Disconnect-ExchangeOnline -Confirm:$false
  exit 1
}

Log "Caixa compartilhada criada. Em até 6 horas, estará disponível no Outlook."
Log ""
Log "Nome: $nome_caixa"
Log "E-mail: $email_caixa"
Log "Usuário $email_solicitante incluído como membro."
Log ""
Log "Para inclusão de novos membros, pesquise por Outlook em https://petrobras.service-now.com/cs"

Disconnect-ExchangeOnline -Confirm:$false

