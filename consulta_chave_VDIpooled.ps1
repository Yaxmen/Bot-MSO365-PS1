 param (
    [string]$chave = "yt17"
 )

. "$PSScriptRoot/autenticacao.ps1"

Try {
    $usuario = Get-ADUser -Identity $chave -Credential $adCredential -ErrorAction Stop
} Catch {
    Echo 'Erro: CHAVE NAO ENCONTRADA'
    Break
}

$userSpn = $usuario.UserPrincipalName
$sessions = @{}

Try {
    Get-AzWvdUserSession -ResourceGroupName $resourceGroup -HostPoolName $vdiPool -Filter "(UserPrincipalName eq '$userSpn')" | %{
        $nome = (($_.Name -split '/')[1]).Replace('.petrobras.biz','')
        Write-Host "Host: $nome"
		Write-Host "Session Id: $(($_.Id -split '/')[-1]) &rarr; $($_.SessionState)"
		Write-Host
        $ip = ([System.Net.Dns]::GetHostAddresses($nome)).IPAddressToString
        $sessions[$ip] = $nome
    }
    Try {    
        Invoke-Command -ComputerName $fileServer -ScriptBlock {Try {Get-SmbSession |Where-Object ClientUserName -eq "PETROBRAS\$Using:chave" } Catch {}} -Credential $adCredential |%{
            $ip = $_.ClientComputerName
            if ($sessions.ContainsKey($ip)) {
                Write-Host "Host onde perfil est&aacute; montado: $($sessions[$ip]) (ok)"
            } else {
                Write-Host "Host onde perfil estava montado equivocadamente: "
                Write-Host "   - $ip"
                Write-Host "   - n&atilde;o h&aacute; sess&atilde;o VDI neste host"
                Write-Host "   - <b>perfil liberado deste host</b>"
                $session = ((quser /server:$ip | ? { $_ -match $chave }) -split ' +')[2]
                logoff $session /server:$ip
                $sid = $_.SessionId
                Invoke-Command -ComputerName $fileServer -ScriptBlock {Close-SmbSession -SessionId $Using:sid -Confirm:$false} -Credential $adCredential
            }
		    Write-Host
        }
    } Catch {}
} Catch {
    Echo 'Erro: NAO FOI POSSIVEL CONTACTAR A NUVEM'
    Break
}