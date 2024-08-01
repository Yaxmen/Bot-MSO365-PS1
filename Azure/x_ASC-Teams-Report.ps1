<#

 Relatorio com todos usuarios nos Grupos GU_ASCRecording_* e respectivos ramais Teams

Get-NetFirewallRule 

#>

<#


Param (

    [Parameter( Mandatory=$true)] [String]$send_to
    
)
#>


$send_to = 
@(
"maicol.hasegawa@petrobras.com.br"
)


$PREFIX = "GU_ASCRecording"
$excecao = "4b135cd3-e816-4637-9142-bc03a569e5de" # "GU_ASCRecording_admins"
$smtp = "smtp.gcorp.petrobras.com.br"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$username = "samig01@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

$SLEEP = 60
$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyy')+"_ASC-phoneSystem_log.csv"
$SKU = "petrobrasbr:MCOEV"

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

    Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
    Log "SUCESSO: ao conectar MicrosoftTeams 1/3"
}
catch {
    Log "ATENCAO: ao conectar MicrosoftTeams 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
        Log "SUCESSO: ao conectar MicrosoftTeams 2/3"
    }
    catch {
        Log "ATENCAO: ao conectar MicrosoftTeams 2/3"
        Start-Sleep $SLEEP
        try {
            Connect-MicrosoftTeams -Credential $credential -InformationAction SilentlyContinue | Out-Null
            Log "SUCESSO: ao conectar MicrosoftTeams 2/3"
        }
        catch {    
            Log "ERRO: ao conectar MicrosoftTeams 3/3"
            exit 1
        }

    }
}


try {

    Connect-AzureAD -Credential $credential -InformationAction SilentlyContinue 
    Log "SUCESSO: ao conectar no AzureAD"
}
catch {

    Start-Sleep 5

    try {

        Connect-AzureAD -Credential $credencial -InformationAction SilentlyContinue 
        Log "SUCESSO: ao conectar no AzureAD"
       
    }
    catch {
        Log "ERRO: ao tentar conectar no AzureAD"
        exit 1
    }
}


$GRUPOS = Get-AzureADGroup -SearchString $PREFIX
  

$OUT = @()

foreach ($grupo in $GRUPOS) {

    $id = $grupo.ObjectId

    if ($id -ne $excecao) {

    $MEMBROS = Get-AzureADGroupMember -ObjectId $id -All $true

    foreach ($membro in $MEMBROS) {

        $upn = $membro.UserPrincipalName
        $csu = Get-CsOnlineUser -Identity $upn

         
        $OUT += [pscustomobject]@{
    
        Usuário = $($csu.DisplayName)
        EMail = $($csu.UserPrincipalName)
        Lotação = $($csu.Department)
        TeamsPhone = $($csu.LineUri) -replace ("tel:","") 
        Grupo = $($grupo.DisplayName)

        }
        
    }
    }
}

$datetime = (Get-Date).ToString('yyyy-MM-dd')
$mes = (Get-Date).ToString('MM')
$ano = (Get-Date).ToString('yyyy')

try {

    $OUT | Export-Excel -Path D:\Util\PhoneSystem\logs\ASC-Teams_$datetime.xlsx -AutoSize -AutoFilter -NoNumberConversion *
    Log "SUCESSO: ao criar arquivo - D:\Util\PhoneSystem\logs\ASC-Teams_$datetime.xlsx"
}
catch {

    Log "ERRO: ao tentar criar arquivo"

}

try {

  
    Send-MailMessage -Attachments D:\Util\PhoneSystem\logs\ASC-Teams_$datetime.xlsx -Body "Arquivo anexo." -From $username -Subject "Relatório de Usuários em Gravação no Sistema Compliance do Teams $mes/$ano." -To $send_to -SmtpServer $smtp -Encoding utf8
    Log "SUCESSO: ao enviar email"

}
catch {

    Log "$Alias;ERRO: ao tentar enviar email"
    
}