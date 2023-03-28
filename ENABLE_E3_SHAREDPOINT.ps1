
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

Import-Module MSOnline

$licenseOption = "OneDrive"
$option = "SHAREPOINTENTERPRISE"

$SkuID = "globalvale:SPE_E3"            

$notfound = @()
#$SPOUsers = @()

$data = GetTickets('ADD_E3')


foreach($ticket in $data)
{    
    $error.Clear()    
    $sposite = $null
    $site = $null
    $Status = $ticket.status
    $errorMessage = $ticket.errorMessage
    $closeNote = $ticket.closeNote
    $WO = $ticket.WO
    if($Status -ne "step6")
    {
        continue
    }

    $upnInfo = Get-Mailbox -Identity $ticket.UPN_Add -ResultSize 5

    $UPN = $upnInfo.userprincipalname
    
    #$UserOnline = Get-MsolUser -UserPrincipalName ($UPN).trim() -ErrorAction SilentlyContinue
    
    
    #$UserOnline = Get-MsolUser -UserPrincipalName ($UPN).trim() | Select UserPrincipalName,DisplayName,BlockCredential,Licenses
    $UserOnline = Get-MsolUser -UserPrincipalName $UPN | Select UserPrincipalName,DisplayName,BlockCredential,Licenses
    if($UserOnline.BlockCredential -eq "True")
    {
        continue
    }

    
    if ($UPN -ne $null)
    {
        for ($lic = 0; $lic -lt $UserOnline.Licenses.Count; $lic++)
        {
            if ($UserOnline.Licenses[$lic].AccountSkuId -eq $SkuID)
            {
                $CorrectLicense = $UserOnline.Licenses[$lic]
                break
            }
        }
        if ($CorrectLicense -eq $null)
        {
            Write-Output ($UserOnline.UserPrincipalName + ": SPO site not requested. No SPE_E3 subscription assigned")
            $Status = "ERROR"
            $errorMessage = "SPO site not requested. No SPE_E3 subscription has been found"
        }
             
        else 
        {
            $site = "https://globalvale-my.sharepoint.com/personal/" + $UserOnline.UserPrincipalName.replace('.', '_').replace('@', '_').toLower()
            try
            {
                $sposite = Get-SpoSite -Identity $site
            }
            catch {}
            if ($sposite -eq $null)
            {
                Request-SPOPersonalSite -UserEmails @( $UPN)

                if($Error[0].FullyQualifiedErrorId.Contains("OneDrive for Business site collection"))
                {
                    $Status = "ERROR"
                    $errorMessage = "[" + $sposite + "] " + "is a OneDrive for Business site collection"
                }
        
                Write-Output ($UserOnline.UserPrincipalName + ": requested personal site creation")
            }
            else 
            {
                Write-Output ($UserOnline.UserPrincipalName + ": site already created")
                if ($sposite.LockState -eq "NoAccess")
                {
                    $sposite | Set-SpoSite -LockState Unlock
                    Write-Output ($UserOnline.UserPrincipalName + ": site already unlocked, no change")
                }
                else
                {
                    Write-Output ($UserOnline.UserPrincipalName + ": site already unlocked, no change")
                }
                
            }
            $Status="Success"
                    
            $CurrentDate = Get-Date -Format G
            InsertProvisioningCheck $WO  $UPN $site $CurrentDate
        }
    }
    else
    {
        Write-Output ($UPN+": user not found")
        $notfound += $UPN
        $Status = "ERROR"
    }
        UpdateTicket $WO $Status $errorMessage $closeNote
    }
 $notfound | Out-Default