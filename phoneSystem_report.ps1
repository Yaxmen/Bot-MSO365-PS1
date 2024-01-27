$username = "samig01@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Util\PhoneSystem\_pass.sec" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

Connect-MicrosoftTeams -Credential $credential | Out-Null

$users = Get-CsOnlineUser -Filter {LineURI -ne $null}

$OUT = @()

foreach ($u in $users) {

    
    
    $OUT += [pscustomobject]@{
    
        UPN = $($u.UserPrincipalName)
        Alias = $($u.Alias)
        DisplayName = $u.DisplayName
        SipAddress = $u.SipAddress
        Department = $u.Department
        Office = $($u.Office)
        PreferredLanguage = $($u.PreferredLanguage)
        TenantDialPlan = $($u.TenantDialPlan)
        OnlineVoiceRoutingPolicy = $($u.OnlineVoiceRoutingPolicy)
        DialPlan = $($u.DialPlan)
        LineURI = $($u.LineURI)
        
        
    }
 }


 $data = date -Format "yyyy-MM-dd"
 $OUT | Export-Csv -Path D:\Util\PhoneSystem\reports\phoneSystem_report_$data.csv -NoTypeInformation -Delimiter ';'