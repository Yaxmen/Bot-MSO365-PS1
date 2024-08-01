#importa funcoes de banco de dados
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

#Baixa informações do banco de dados
$data = GetTickets 'REMOVE_E3'

foreach ($user in $data)
{
    $WO = $user.WO
    $UPN = $user.UPN_Remove    
    [datetime]$sbDate = $user.SubmitDate

    [datetime] $ticketDate = $user.BotLastUpdate
    [datetime] $currentDate = Get-Date
    $status = $user.Status
    $state = $user.state

    $elapsedTime =  ($currentDate - $ticketDate).TotalMinutes
   
    ## tratamento de tickets com mais de 20 horas ################################
    if($elapsedTime -gt  1200) # 1200 = 20*60
       {
           $status = "ERROR"
           $errorMessage = "SLA breach risk. Ticket awaiting more than 20 hours."
           UpdateTicket $WO $Status $errorMessage  $closeNote
           continue
       }
    ###############################################################################


    $upnInfo = ""    
    $errorMessage = ""
    $closeNote = ""

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
        try
	    {
            $upnInfo = Get-mailbox -identity $UPN
		    $userPrincipalName = Get-MsolUser -UserPrincipalName $upnInfo.UserPrincipalName
		    $status = "Step1"
	    }
	    catch
	    {
		    $status = "ERROR"
            if($errorMessage -eq "")
            {
                $errorMessage = "Error while getting UPN from tenant"
            }
	    }

	    if($status -eq "Step1")
	    {
		    try
		    {
			    #Set-msoluserlicense -UserPrincipalName $userPrincipalName.UserPrincipalName -removelicense "globalvale:ENTERPRISEPACK"
                Set-msoluserlicense -UserPrincipalName $userPrincipalName.UserPrincipalName -removelicense "globalvale:SPE_E3"
			    Set-msoluserlicense -UserPrincipalName $userPrincipalName.UserPrincipalName -removelicense "globalvale:ATP_ENTERPRISE"
			    Set-msoluserlicense -UserPrincipalName $userPrincipalName.UserPrincipalName -removelicense "globalvale:RIGHTSMANAGEMENT"

			    $status = "Success"

                $closeNote = "Usuário Removido do Ambiente O365|jump||jump|User Terminated from O365 Environment"
		    }
		    catch
		    {
			    $status = "ERROR"
                if($errorMessage -eq "")
                {
                    $errorMessage = "Failed to remove licenses"
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