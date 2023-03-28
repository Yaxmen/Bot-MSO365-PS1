#importa funcoes de banco de dados
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

#Baixa informações do banco de dados
$data = GetTickets('MODIFY_E3')

foreach ($row in $data) 
{
    $WO = $row.WO
    $UPN = $row.UPN_Add  
    [datetime] $sbDate = $ticket.SubmitDate
    [datetime] $ticketDate = $ticket.BotLastUpdate
    [datetime] $currentDate = Get-Date
    $status = $ticket.Status
    $state = $ticket.state

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
        if(Get-Mailbox $UPN)
        {
            $Status = "Success"
            $closeNote = "ID tranferido para ID " + $UPN + " conforme solicitação.|jump|Att,|jump||jump|O365 Team|jump||jump|********************************************* |jump||jump||jump|Transfered ID to " + $UPN + " completed as requested.|jump|Regards,|jump||jump|O365 Team"

        }
        else
        {
            $Status = "ERROR"
            if($errorMessage = "")
            {
                $errorMessage = "Mailbox " + $UPN + " unupdated"
            }            
        }
    }
    else
    {
        $Status = "ERROR"
        if($errorMessage = "")
        {
            $errorMessage = "PowerQuery problems."
        }
    }
       
    UpdateTicket $WO $Status $errorMessage  $closeNote
}
