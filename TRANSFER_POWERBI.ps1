#importa funcoes de banco de dados
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

#Baixa informações do banco de dados
$data = GetTickets('PowerBI_TRANSFER')

foreach($user in $data)
{
    $WO = $user.WO
    $UPN_Remove = $user.UPN_Remove
    $UPN_Add = $user.UPN_Add   
    [datetime] $sbDate = $user.SubmitDate
    [datetime] $ticketDate = $user.BotLastUpdate
    [datetime] $currentDate = Get-Date
    $state = $user.state
    $status = $user.Status

    $elapsedTime =  ($currentDate - $ticketDate).TotalMinutes
   
    ## tratamento de tickets com mais de 20 horas ################################
    # if($elapsedTime -gt  1200) # 1200 = 20*60
    #    {
    #        $status = "ERROR"
    #        $errorMessage = "SLA breach risk. Ticket awaiting more than 20 hours."
    #        UpdateTicket $WO $Status $errorMessage  $closeNote
    #        continue
    #    }
    ###############################################################################


    
    $upnInfo = ""    
    $errorMessage = ""
    $closeNote = ""
    $SPOSite = "" 

    if($WO -eq "")
    {
        $WO = "ERROR"
        $UPN = "ERROR"
        $status = "ERROR"
        $upnInfo = "ERROR"
        $errorMessage = "ERROR"
        $closeNote = "ERROR"
    }

    if (($status -eq "step 0") -or ($status -eq "step0"))
    {
        #Check if have remove_user have POWER_BI_PRO
        try
        {
            $upnInfo_Remove = Get-MsolUser -UserPrincipalName $UPN_Remove

            #Check if user has the current license
            $haveLicense_Remove = $false
        
            $userLicense = (Get-MsolUser -UserPrincipalName $upnInfo_Remove.UserPrincipalName).licenses.accountskuid #Get all user licenses
            $numLicense = $userLicense.Count

            #Check if user already have E3
            if ($numLicense -eq 1)
            {
                if ($userLicense -eq "globalvale:POWER_BI_PRO") #check if the only licenses is the E3
                {
                    $haveLicense_Remove = $true
                }
            }
            else
            {
                For ($i=0; $i -ne $numLicense; $i++) #check in all licenses if have the E3
                {
                    if ($userLicense[$i] -eq "globalvale:POWER_BI_PRO")
                    {
                        $haveLicense_Remove = $true
                    }
                }
            }            
        }
        catch
        {
            $status = "ERROR"
            if($errorMessage -eq "")
            {
                $errorMessage = "Cannot found user and licenses of $upn in tenant."
            }
        }

        #Check if have Add_user have POWER_BI_PRO
        try
        {
            $upnInfo_Add = Get-MsolUser -UserPrincipalName $UPN_Add

            #Check if user has the current license
            $haveLicense_Add = $false
        
            $userLicense = (Get-MsolUser -UserPrincipalName $upnInfo_Add.UserPrincipalName).licenses.accountskuid #Get all user licenses
            $numLicense = $userLicense.Count

            #Check if user already have E3
            if ($numLicense -eq 1)
            {
                if ($userLicense -eq "globalvale:POWER_BI_PRO") #check if the only licenses is the E3
                {
                    $haveLicense_Add = $true
                }
            }
            else
            {
                For ($i=0; $i -ne $numLicense; $i++) #check in all licenses if have the E3
                {
                    if ($userLicense[$i] -eq "globalvale:POWER_BI_PRO")
                    {
                        $haveLicense_Add = $true
                    }
                }
            }
        }
        catch
        {
            $status = "ERROR"
            if($errorMessage -eq "")
            {
                $errorMessage = "Cannot found user and licenses of $upn in tenant."
            }
        }

        if($haveLicense_Remove -eq $true)
        {
            Set-MsolUserLicense -UserPrincipalName $upnInfo_Remove.UserPrincipalName -RemoveLicenses "globalvale:POWER_BI_PRO" -LicenseOptions $disabledPlans
            $haveLicense_Remove = $false
        }

        if($haveLicense_Add -eq $false)
        {
            Set-MsolUserLicense -UserPrincipalName $upnInfo_Add.UserPrincipalName -AddLicenses "globalvale:POWER_BI_PRO" -LicenseOptions $disabledPlans
            $haveLicense_Add = $true
        }

        if(($haveLicense_Remove -eq $false) -and ($haveLicense_Add -eq $true))
        {
            $status = "Success"
            $closeNote = "Prezado usuário,|jump||jump|Aplicamos a licença de PowerBI Pro para o usuário "+ $upnInfo_Add.UserPrincipalName + " e removemos a licença do usuário " + $upnInfo_Remove.UserPrincipalName + " conforme solicitado.|jump||jump||jump|Dear User,|jump||jump|We assigned the license PowerBI Pro to the user "+ $upnInfo_Add.UserPrincipalName + " and we removed the license from the user " + $upnInfo_Remove.UserPrincipalName + " as requested."
        }
        else
        {
            $status = "ERROR"
            if($errorMessage -eq "")
            {
                $errorMessage = "Failed while transfering licenses."
            }
        }
    }
    else
    {
        $status = "ERROR"
        if($errorMessage -eq "")
        {
            $errorMessage = "PowerQuery errors"
        }
    }

    UpdateTicket $WO $Status $errorMessage  $closeNote
}