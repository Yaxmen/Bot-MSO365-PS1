 param (
    [string]$chave = "XXXX"
 )


. "$PSScriptRoot/autenticacao.ps1"


Try {
    $usuario = Get-ADUser -Identity $chave -Credential $adCredential -ErrorAction Stop
} Catch {
    Echo 'Erro: CHAVE NAO ENCONTRADA'
    Break
}

$userSpn = $usuario.UserPrincipalName

<#Try {#>
    Get-AzWvdUserSession -ResourceGroupName $resourceGroup -HostPoolName $vdiPool -Filter "(UserPrincipalName eq '$userSpn') and (SessionState eq 'Active')" | %{
        $hn = (($_.Name -split '/')[1]).Replace('.petrobras.biz','')
        $sid = [int] ($_.Id -split '/')[-1]
        $ok = Disconnect-AzWvdUserSession -ResourceGroupName $resourceGroup -HostPoolName $vdiPool -SessionHostName "$hn.petrobras.biz" -Id $sid
        if ($ok) {
            Write-Host "Logoff de sess&atilde;o em $hn foi realizado!"
        } else {
            Write-Host "Problemas para realizar logoff de sess&atilde;o em $hn!"
        }

<#        $session = ((quser /server:$hn | ? { $_ -match $chave }) -split ' +')[2]
        logoff $session /server:$hn
        Write-Host "Logoff de sess&atilde;o em $hn foi realizado!"#>
    }
<#} Catch {
    Write-Host "ERRO: Problemas para realizar logoff de sess&atilde;o em $hn!"
    Break
}#>