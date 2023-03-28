#importa funcoes de banco de dados
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

#Baixa informações do banco de dados
$data = GetTickets('PowerBI_ADD')


#Check licenses availability
$Licenses = Get-MsolAccountSku
$Global:totalLicenses = 0

foreach($lic in $Licenses)
{
    if($lic.AccountSkuId -eq 'globalvale:POWER_BI_PRO')
    {
        echo("Available licenses: " + ($lic.ActiveUnits - $lic.ConsumedUnits).ToString())
        $totalLicenses = ($lic.ActiveUnits - $lic.ConsumedUnits)
        $totalLicenses = $totalLicenses - 30
    }
}

foreach($user in $data)
{
    $WO = $user.WO
    $UPN = $user.UPN_Add    
    
    [datetime] $sbDate = $user.SubmitDate
    [datetime] $ticketDate = $user.BotLastUpdate
    [datetime] $currentDate = Get-Date
    $status = $user.Status

    $elapsedTime =  ($currentDate - $ticketDate).TotalMinutes
   
    ## tratamento de tickets com mais de 20 horas ##########################
    # if($elapsedTime -gt  1200) # 1200 = 20*60
    #     {
    #         $status = "ERROR"
    #         $errorMessage = "SLA breach risk. Ticket awaiting more than 20 hours."
    #     }
    ########################################################################

    
    
    $upnInfo = ""    
    $errorMessage = ""
    $closeNote = ""
    $SPOSite = "" 

    if($totalLicenses -ge 1)
    {
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
            #Check if have POWER_BI_PRO
            try
            {
                $upnInfo = Get-MsolUser -UserPrincipalName $UPN

                #Check if user has the current license
                $haveLicense = $false
        
                $userLicense = (Get-MsolUser -UserPrincipalName $upnInfo.UserPrincipalName).licenses.accountskuid #Get all user licenses
                $numLicense = $userLicense.Count
        
                #Check if user already have E3
                if ($numLicense -eq 1)
                {
                    if ($userLicense -eq "globalvale:POWER_BI_PRO") #check if the only licenses is the E3
                    {
                        $haveLicense = $true
                        $closeNote = "Prezado usuário,|jump||jump|Aplicamos a licença de PowerBI Pro para o usuário $UPN conforme solicitado.|jump||jump||jump|Dear User,|jump||jump|We assigned the license PowerBI Pro to the user $UPN as requested."
                        $status = "Success"
                    }
                }
                else
                {
                    For ($i=0; $i -ne $numLicense; $i++) #check in all licenses if have the E3
                    {
                        if ($userLicense[$i] -eq "globalvale:POWER_BI_PRO")
                        {
                            $haveLicense = $true
                            $closeNote = "Prezado usuário,|jump||jump|Aplicamos a licença de PowerBI Pro para o usuário $UPN conforme solicitado.|jump||jump||jump|Dear User,|jump||jump|We assigned the license PowerBI Pro to the user $UPN as requested."
                            $status = "Success"
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

            if($haveLicense -ne $true)
            {
                try
                {
                    Set-MsolUserLicense -UserPrincipalName $upnInfo.UserPrincipalName -AddLicenses "globalvale:POWER_BI_PRO"
                    $closeNote = "Prezado usuário,|jump||jump|Aplicamos a licença de PowerBI Pro para o usuário $UPN conforme solicitado.|jump||jump||jump|Dear User,|jump||jump|We assigned the license PowerBI Pro to the user $UPN as requested."
                    $status = "Success"

                    $totalLicenses = $totalLicenses - 1
                }
                catch
                {
                    $status = "ERROR"
                    if($errorMessage -eq "")
                    {
                        $errorMessage = "Cannot apply BI Pro license."
                    }
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
    else
    {
        #Filling in the variables with report data
        $ID = $row.WO
        $Status = "ERROR"
        $errorMessage = "Out of licenses"
        $ticketType = $row.ticketType
        $CloseNote = ""

        UpdateTicket $ID $Status $errorMessage  $closeNote
    }
}

