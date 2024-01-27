 param (
    [string]$chave = "xxxx"
 )


#. "$PSScriptRoot/autenticacao.ps1"
$fileServer = "fsvdi-win-vm"
$adUser = 'sasmav'
$adPasswordCifrado = 'Y29tQENlc3Mw'
$adPassword = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($adPasswordCifrado))
$adSecurePassword = ConvertTo-SecureString "$adPassword" -AsPlainText -Force
$adCredential = New-Object -TypeName PSCredential -ArgumentList $adUser, $adSecurePassword

#Try {
#    $usuario = Get-ADUser -Identity $chave -Credential $adCredential -ErrorAction Stop
#} Catch {
#    Echo 'Erro: CHAVE NAO ENCONTRADA'
#    Break
#}

#$userSpn = $usuario.UserPrincipalName

# Derruba sessão na Azure que possa haver para essa chave (pode estar disconnected)
$vms = Invoke-Command -ComputerName $fileServer -ScriptBlock {Get-SmbOpenFile |Where-Object {$_.ClientUserName -eq "PETROBRAS\$Using:chave" } | Select-Object -ExpandProperty ClientComputerName |Sort-Object |Get-Unique} -Credential $adCredential
$vms |% {
    $vm = $_
    $session = ((quser /server:$vm | ? { $_ -match $chave }) -split ' +')[2]
    logoff $session /server:$vm
}

#Get-AzWvdUserSession -ResourceGroupName $resourceGroup -HostPoolName $vdiPool -Filter "(UserPrincipalName eq '$userSpn')" | %{
#	$vm = (($_.Name -split '/')[1]).Replace('.petrobras.biz','')
#	$sessionid = $(($_.Id -split '/')[-1])
#	logoff $sessionid /server:$vm
#}

# Tenta remover disco de perfil
Try {    
    Invoke-Command -ComputerName $fileServer -ScriptBlock {Try { If (Get-ChildItem -Path s:\*$Using:chave.vhd -Recurse -ErrorAction Ignore) { Get-ChildItem -Path s:\*$Using:chave.vhd -Recurse | Move-Item -Destination "T:\" -Force -Confirm:$false } } Catch {}} -Credential $adCredential
} Catch {}

# Verifica se disco foi removido
Try {    
    Invoke-Command -ComputerName $fileServer -ScriptBlock { If (Get-ChildItem -Path s:\*$Using:chave.vhd -Recurse -ErrorAction Ignore) { Write-Host "Problema para remover perfil $Using:chave" } else { Write-Host "Perfil $Using:chave removido com sucesso!" } } -Credential $adCredential
} Catch {
    Write-Host "Problema para remover perfil $chave"
} 
