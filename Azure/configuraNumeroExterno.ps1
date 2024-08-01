$lote = Import-Csv -Path D:\Util\PhoneSystem\configuraNumeroExterno.csv -ErrorAction Stop

$username = "@.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

Connect-MicrosoftTeams -Credential $credential | Out-Null

$i = 0
foreach ($u in $lote)
{
    $i++
    $upn = $u.UserName
    $tel = $u.PhoneNumber

    Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $tel -PhoneNumberType DirectRouting
    Write-Host "$i $upn $tel"
}
