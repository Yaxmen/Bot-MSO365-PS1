

$LISTA = Import-Excel -Path 'D:\Util\PhoneSystem\reports\ramal_teams_sem_DDR.xlsx'


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$username = "samig01@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

$SLEEP = 60
$logFile = "d:\Util\PhoneSystem\logs\"+(Get-Date).ToString('yyyy')+"_phoneSystem_lote_log.csv"
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
    Connect-MsolService -Credential $credential -InformationAction SilentlyContinue | Out-Null
    Log "SUCESSO: ao conectar MSOLService 1/3"
}
catch {
    Log "ATENCAO: ao conectar MSOLService 1/3"
    Start-Sleep $SLEEP
    try {
        Connect-MsolService -Credential $credential -InformationAction SilentlyContinue | Out-Null
        Log "SUCESSO: ao conectar MSOLService 2/3"
    }
    catch {
        Log "ATENCAO: ao conectar MSOLService 2/3"
        Start-Sleep $SLEEP
        try { 
            Connect-MsolService -Credential $credential -InformationAction SilentlyContinue | Out-Null
            Log "SUCESSO: ao conectar MSOLService 3/3"
        }
        catch {
            Log "ERRO: ao conectar MSOLService 3/3"
            exit 1
        }
    }
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


foreach ($user in $LISTA) {

  $Chave = $user.Chave_ou_Email
  $Ramal = $user.Ramal
  $Categoria = $user.Categoria
  $Acao = $user.Acao

  write-host "$Acao $Chave $Ramal $Categoria"
  pause
  
  

### consulta, no AD interno qual $UPN da $Chave informada. Caso não encontre encerra o script.



if ($Chave -like "*@petrobras.com.br") {

 $UPN = $Chave

}
else { 

    try {

        $azu = Get-ADUser $Chave
        $UPN = $azu.UserPrincipalName
        Log "$UPN;SUCESSO: chave $chave encontrada no AD"
    }
    catch {
        Log "$chave;ERRO: chave não encontrada no AD"
    
    
    }

}
    
### verifica as configurações do usuario $UPN no diretorio do Teams.

try {
    #$OnlineUser = Get-CsOnlineUser -Filter "UserPrincipalName -eq '$UPN'"
    $OnlineUser = Get-CsOnlineUser -Identity $UPN
}
catch {
    Log "$UPN;ERRO: ao executar Get-CsOnlineUser -Filter"
    
}

### encerra o script, caso dados do usuario $UPN retorne nulo no diretório do Teams

if ($OnlineUser -eq $null) {
    Log "$UPN;ERRO: usuário inválido"
    
}

### configurando phoneSystem (acao = add)

if ($Acao -eq "add") {

    $cod_hash = @{

    '21' = 'Tag:DP-BR-21'
    '22' = 'Tag:DP-BR-22'
    '13' = 'Tag:DP-BR-13'
    '27' = 'Tag:DP-BR-27'

    }

    $cod = $Ramal -match '\+55(\d\d)\d\d\d\d\d\d\d\d'

    ### encerra processamento se ramal foi enviado no formato errado

    if (-not($cod)) {
        Log "$UPN;ERRO: $Ramal com formato inválido"
        
    }

    ### encerra processamento se codigo DDD não constar na tabela $cod_hash

    $cod = $matches[1]
    $plano = $cod_hash[$cod]

    if (-not($cod_hash.ContainsKey($cod))) {
        Log "$UPN;ERRO: não encontrado plano para código $cod"
        
    }

    $linha = "tel:"+"$Ramal"

    ### verifica se usuario já esta configurado com ramal solicitado

    if ($($OnlineUser.LineURI) -eq $linha) {

        Log "$UPN;ATENCAO: já configurado Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"
        
    }

    ### verifica se linha solicitada encontra-se disponível

    try {
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq '$ramal'"
    }
    catch {
        Log "$UPN;ERRO: ao executar Get-CsOnlineUser -Filter Ramal"
        exit 1
    }

    ### encerra script caso ramal solicitado já esteja em uso

    if ($OnlineRamal -ne $null) {

        $status_ramal =  "$($OnlineRamal.Alias)"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"        
        Log "$UPN;ATENCAO: previamente configurado o numero $Ramal para chave $status_ramal" 
         
    }

    if ($($OnlineUser.LineURI) -ne $null -and $($OnlineUser.LineURI) -ne "") {
    
        Log "$UPN;ATENCAO: já configurado Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"
         
    }

    ### verifica se usuario possui licença e aplica caso não tenha

    $msu = Get-MsolUser -UserPrincipalName $UPN

    if ($msu.Licenses.AccountSkuId -notcontains $SKU) {

        # verifica se existem licenças MCOEV disponiveis no ambiente

        $plans = Get-MsolAccountSku

        foreach ($i in $plans) {

            if ($($i.AccountSkuId) -eq "$SKU") {
                $disp_MCOEV = $($i.ActiveUnits) - $($i.ConsumedUnits)
            }
        }

        if ($disp_MCOEV -lt 1) {
            Log "$UPN;ERRO: $($disp_MCOEV) $SKU - insuficientes para configurar usuario"
            #exit 1
        }

        #Log "$UPN;ATENCAO: $($disp_MCOEV) $SKU - disponiveis" 

        ### aplica licença e aguarda 60 seg

        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "$SKU"
        #Log "$UPN;ATENCAO: aguardando aplicar licença do phoneSystem aguardar + $SLEEP seg"
        $SLEPPT = $SLEEP
        $NOK = $true

        ### aguarda +20 seg ate que seja verificada a aplicacao da licença no modulo Microsoft Teams
        ### no maximo aguarda ate 600 seg

        while ($NOK) {
            Start-Sleep $SLEEP
            $resp = Get-CsOnlineUser -Identity $UPN | select FeatureTypes
            if ($($resp.FeatureTypes) -ccontains "PhoneSystem") {
                $NOK = $false
                #Log "$UPN;SUCESSO: licenca phoneSystem aplicada em $SLEPPT seg"
            }
            else {
                $SLEEP = 20
                $SLEEPT += $SLEEP
                Log "$UPN;ATENCAO: aguardando aplicar licença do phoneSystem aguardar + $SLEEP seg"
                if ($SLEEPT -gt 360) {
                    Log "$UPN;ERRO: licenca phoneSystem NÃO aplicada após $SLEEPT seg"
                    $NOK = $false
                }
                   
            }
        }    
        
    }
    else {
        Log "$UPN;ATENCAO: previamente configurado $SKU"
    }
    
    Start-Sleep 30
    
    $cod = Set-CsPhoneNumberAssignment -Identity $UPN -PhoneNumber $Ramal -PhoneNumberType DirectRouting
    if ($cod -ccontains "has already been assigned to another user") {
        Start-Sleep 30
        $OnlineRamal = Get-CsOnlineUser -Filter "LineURI -eq '$ramal'"
        $status_ramal =  "$($OnlineRamal.Alias)"+" categoria "+"$($OnlineRamal.OnlineVoiceRoutingPolicy)"
        Log "$UPN;ATENCAO: previamente configurado o numero $Ramal para chave $status_ramal" 
        
    }

    Set-CsPhoneNumberAssignment -Identity $UPN -EnterpriseVoiceEnabled $true

    Start-Sleep 30

    $csu = Get-CsOnlineUser -Identity $UPN
    if ($($csu.LineUri) -eq $linha -and $($csu.EnterpriseVoiceEnabled)) {
        Log "$UPN;configurado: EnterpriseVoice True - HostedVoiceMail True - LineURI $Ramal"
    }
    else {
        Log "$UPN;ERRO: EnterpriseVoice True - HostedVoiceMail True - LineURI $linha"
        
    }
    
    
    Start-Sleep 10

    ### define categoria
                        
    if ($Categoria -eq "DDD") {
        $vrota = "Tag:VRP-BR-Nacional"
    }
    else {
        if ($Categoria -eq "DDI") {
            $vrota = "Tag:VRP-BR-Internacional"
        }
        else {
            $vrota = "Tag:VRP-BR-Interno"
        }
    }
    
    try {
        Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $vrota 
    }
    catch {
        Log "$UPN;ERRO: OnlineVoiceRoutingPolicy $vrota"
        
    }

    Log "$UPN;configurado: OnlineVoiceRoutingPolicy $vrota"
    Start-Sleep 10 
    
    try {   
        Grant-CsTenantDialPlan -Identity $UPN -PolicyName $plano 
    }
    catch {
        Log "$UPN;ERRO: TenantDialPlan $plano"
        
    }
            Log "$UPN;configurado: TenantDialPlan $plano"
            
}
else {

    if ($Acao -eq "del") {

        if ($OnlineUser.LineURI -eq $null)
        {
            Log "$UPN;ATENCAO: usuário não está configurado"
            
        }

        Log "$UPN;antes de remover - Linha $($OnlineUser.LineURI) - $($OnlineUser.OnlineVoiceRoutingPolicy) - $($OnlineUser.TenantDialPlan)"

        try {

            #Grant-CsTenantDialPlan -Identity $UPN -PolicyName $null
            #Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $null
            Remove-CsPhoneNumberAssignment -Identity $UPN -RemoveAll
        }
        catch {

            Log "$UPN;ERRO: uao desconfigurar phoneSystem"
            

        }
        
        # remover licença
        $msu = Get-MsolUser -UserPrincipalName $UPN

        if ($msu.Licenses.AccountSkuId -contains $SKU) {

            Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "$SKU" 2>&1
            Log("$UPN;removida licença $SKU")
        }
        else {
            Log("$UPN;ATENCAO: previamente removida licença $SKU")
        }

    }
    else {

        Log  "$UPN;ação $Acao inválida"
        
    }
    

}

    #Write-Host " Esperando a MS se recuperar em 30 seg ..."
    #Start-Sleep 30


} # end loop lista

try {

    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
}
catch {
    Start-Sleep 10
    Disconnect-MicrosoftTeams -InformationAction SilentlyContinue | Out-Null
}

