$tenantname = "-vdi"

$spnApplicationId = '89d51029-4ceb-4aaf-90fd-815f463784ff' # Cloud Automation
$spnPasswd = '=ot2.qtZ=56z3NJ=N2uvB8.C8Fn96MjU'

$aadTenantId = '5b6f6241-9a57-4be4-8e50-1dfa72e79a57'
$spnCredential = New-Object System.Management.Automation.PSCredential($spnApplicationId, (ConvertTo-SecureString $spnPasswd -AsPlainText -Force))

$vdiPool = "-POOL-001"
$rdpPool = "-POOL-001"
$resourceGroup = "proj-00016-wvd-rg"

$adUser = ''
$adPasswordCifrado = ''
$adPassword = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($adPasswordCifrado))
$adSecurePassword = ConvertTo-SecureString "$adPassword" -AsPlainText -Force
$adCredential = New-Object -TypeName PSCredential -ArgumentList $adUser, $adSecurePassword

$admUser = ''
$admPasswordCifrado = '='
$admPassword = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($admPasswordCifrado))
$admSecurePassword = ConvertTo-SecureString "$admPassword" -AsPlainText -Force
$admCredential = New-Object -TypeName PSCredential -ArgumentList $admUser, $admSecurePassword

$fileServer = "fsvdi-win-vm"

Try {
    Connect-AzAccount -Credential $spnCredential -ServicePrincipal -Tenant $aadTenantId -ErrorAction Stop -WarningAction silentlyContinue 2>&1 | Out-Null
} Catch {
   Echo 'Erro: NAO FOI POSSIVEL SE LOGAR NA NUVEM'
    Break
}


function setaTag {
    param ($vm, $tags)

    $vm_resource = Get-AzResource -ResourceGroupName $resourceGroup -Name $vm
    Set-AzResource -ResourceId $vm_resource.Id -Tag $tags -Force
}

function removeTag {
    param ($vm, $tagname)

    $vm_resource = Get-AzResource -ResourceGroupName $resourceGroup -Name $vm
    $vm_resource.Tags.Remove($tagname)
	$vm_resource | Set-AzResource -Force
}

function setDrainMode {
    param ($vm, $hostpool, $drainmode)

    $allownewsession = !$drainmode
    $name = $vm.Split(".")[0]
    Try {
        Update-AzWvdSessionHost -ResourceGroupName $resourceGroup -HostPoolName $hostpool -Name "$name" -AllowNewSession:$allownewsession |Out-Null
    } Catch {}
    Try {
        Update-AzWvdSessionHost -ResourceGroupName $resourceGroup -HostPoolName $hostpool -Name "$name.petrobras.biz" -AllowNewSession:$allownewsession |Out-Null
    } Catch {}
}
