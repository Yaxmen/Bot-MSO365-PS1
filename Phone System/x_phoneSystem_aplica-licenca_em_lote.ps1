

$LISTA = Import-Excel -Path 'D:\Util\PhoneSystem\reports\Lista de usuários para criação de ramal Teams - Gravando Sem DDR Teams.xlsx'


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$username = "samig01@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

$SLEEP = 10
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




foreach ($user in $LISTA) {

  $Chave = $user.Chave
  $Ramal = $user.Ramal
  $Categoria = $user.Categoria
  $Acao = $user.Acao

  write-host "$Acao $Chave $Ramal $Categoria"


    ### consulta, no AD interno qual $UPN da $Chave informada. Caso não encontre encerra o script.

  try {
  
    $azu = Get-ADUser $Chave
    $UPN = $azu.UserPrincipalName
    Log "$UPN;SUCESSO: chave $chave encontrada no AD"
  }
  catch {
    Log "$chave;ERRO: chave não encontrada no AD"
  }
  
  $msu = Get-MsolUser -UserPrincipalName $UPN

  if ($msu.Licenses.AccountSkuId -notcontains $SKU) {
      
    Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "$SKU"
    Log "$UPN;SUCESSO: aplicada licença $SKU"
    
  }
  else {
    Log "$UPN;ATENCAO: previamente configurado $SKU"
  }
  
  #Write-Host "aguarda $SLEEP seg ..."
  #Start-Sleep $SLEEP

} # end loop lista


