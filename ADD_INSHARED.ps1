#importa funcoes de banco de dados
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

#Baixa informações do banco de dados
$data = GetTickets('ADD_InShared')

foreach ($row in $data) 
{
    $WO = $row.WO
    $UPN = $row.UPN_Add  
    $SharedMailbox_Address = $row.targetAddress
    [datetime] $sbDate = $row.SubmitDate
    [datetime] $ticketDate = $row.BotLastUpdate
    [datetime] $currentDate = Get-Date
    $status = $row.Status
    $state = $row.state

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

    if(($WO -eq "") -or (($SharedMailbox_Address -eq '') -or ($SharedMailbox_Address -eq $null)))
    {
        $status = "ERROR"
        $errorMessage = "Shared mailbox not provided"
    }

	if (($status -eq "step 0") -or ($status -eq "step0"))
    {
        $upnInfo = Get-Mailbox -Identity $UPN

        #check if user exist or has E3
        if(($upnInfo.UserPrincipalName -ne "") -or ($upnInfo.UserPrincipalName -ne $null))
        {
            $msolUser = Get-MsolUser -UserPrincipalName $upnInfo.userprincipalname

            #Check if user has the current license
            $haveLicense = $false
        
            $userLicense = (Get-MsolUser -UserPrincipalName $upnInfo.UserPrincipalName).licenses.accountskuid #Get all user licenses
            $numLicense = $userLicense.Count
        
            #Check if user already have E3
            if ($numLicense -eq 1)
            {
                #if ($userLicense -eq "globalvale:ENTERPRISEPACK") #check if the only licenses is the E3
                if ($userLicense -eq "globalvale:SPE_E3") #check if the only licenses is the E3
                {
                    $haveLicense = $true
                }
            }
            else
            {
                For ($i=0; $i -ne $numLicense; $i++) #check in all licenses if have the E3
                {
                    if ($userLicense -eq "globalvale:SPE_E3") #check if the only licenses is the E3
                    {
                        $haveLicense = $true
                    }
                }
            }

            if($haveLicense -eq $true)
            {
                $status = "step2"
            }
            else
            {
                $status = "ERROR"
                if($errorMessage -eq "")
                {
                    $errorMessage = "User does not have SPE_E3"
                }
            }
        }

        #check if sharedmailbox exist
        if($status -eq "step2")
        {
            $sharedInfo = $null
            $sharedInfo = Get-Mailbox -Identity "$SharedMailbox_Address"
            if (($sharedInfo -eq "" -or $sharedInfo -eq $null))
            {
             $status = "ERROR"
             $errorMessage = "Shared mailbox $SharedMailbox_Address not found"
            }

            #check if user exist or has E3
            elseif(($sharedInfo.UserPrincipalName -ne "") -or ($sharedInfo.UserPrincipalName -ne $null))
            {
                $status = "step3"
            }
            else
            {
                $status = "ERROR"
                if($errorMessage -eq "")
                {
                    $errorMessage = "User does not have SPE_E3"
                }
            }
        }

        #if user is ok and shared is ok, apply license
        if($status -eq "step3")
        {
            $Error.Clear()
            Add-MailboxPermission $sharedInfo.UserPrincipalName –User $upnInfo.UserPrincipalName –AccessRights FullAccess –InheritanceType all
            Add-RecipientPermission -Identity $sharedInfo.UserPrincipalName -Trustee $upnInfo.UserPrincipalName -AccessRights SendAs -confirm:$false

            if ($Error[0] -eq $null)
            {
                $status = "Success"
                $closeNote = "Olá|jump||jump|Usuário adicionado à caixa compartilhada conforme solicitação.|jump||jump|Att.|jump||jump|O365 Team|jump||jump||jump|RESOLUTION NOTE:|jump||jump||jump|Root cause: N/A|jump||jump|Immediate action: N/A|jump||jump|Corrective action: Usuário adicionado à caixa compartilhada conforme solicitação.|jump||jump|Tests: N/A|jump||jump|Additional comments: N/A|jump||jump||jump|****************************************************************************************************************************** |jump||jump||jump|Hola|jump||jump|Según lo solicitado hemos agregado el usuario al buzón compartido.|jump||jump|Saludos|jump||jump|O365 Team |jump||jump||jump|RESOLUTION NOTE:|jump||jump||jump|Root cause: N/A|jump||jump|Immediate action: N/A|jump||jump|Corrective action: Según lo solicitado hemos agreagado el usuario al buzón compartido.|jump||jump|Tests: N/A|jump||jump|Additional comments: N/A |jump||jump||jump|******************************************************************************************************************************|jump||jump||jump|Hello|jump||jump|User added on shared mailbox as requested.|jump||jump|Regards|jump||jump|O365 Team|jump||jump||jump|RESOLUTION NOTE:|jump||jump||jump|Root cause: N/A|jump||jump|Immediate action: N/A|jump||jump|Corrective action: User added on shared mailbox as requested.|jump||jump|Tests: N/A|jump||jump|Additional comments: N/A"
            }
            else {
                $Status = "ERROR"
                $errorMessage = "Error when adding user in Shared Mailbox"
            }
        }
        else
        {
            $Status = "ERROR"
            if($errorMessage -eq "")
            {
                $errorMessage = "User or shared not found."
            }
        }
    }
    else
    {
        $Status = "ERROR"
        if($errorMessage -eq "")
        {
            $errorMessage = "PowerQuery problems."
        }
    }

    <#$status = "ERROR"
    $errorMessage = "Template not fully added yet"#>
       
    UpdateTicket $WO $Status $errorMessage  $closeNote 
}