#importa funcoes de banco de dados
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

#Baixa informações do banco de dados
$data = GetTickets 'GrantOneDrive'

foreach($user in $data)
{
    $WO = $user.WO
    $UPN = $user.UPN_Add    
    
    [datetime] $sbDate = $ticket.SubmitDate
    [datetime] $ticketDate = $ticket.BotLastUpdate
    [datetime] $currentDate = Get-Date

    $elapsedTime =  ($currentDate - $ticketDate).TotalMinutes
   
    ## tratamento de tickets com mais de 20 horas ##########################
    # if($elapsedTime -gt  1200) # 1200 = 20*60
    #     {
    #         $status = "ERROR"
    #         $errorMessage = "SLA breach risk. Ticket awaiting more than 20 hours."
    #     }
    ########################################################################

    $status = $user.Status
    
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
        #Check if have E3
        try
        {
            $upnInfo = Get-MsolUser -UserPrincipalName $UPN

            #Check if user has the current license
            $haveLicense = $false
        
            $userLicense = (Get-MsolUser -UserPrincipalName $UPN).licenses.accountskuid #Get all user licenses
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
                    if ($userLicense[$i] -eq "globalvale:ENTERPRISEPACK")
                    {
                        $haveLicense = $true
                    }
                }
            }

            $userBlocked = Get-MsolUser -UserPrincipalName $UPN | select BlockCredential

            if (($haveLicense -eq $true) -and ($userBlocked.BlockCredential -eq $false))
            {
                $status = "step1"
            }
            else
            {
                $status = "ERROR"
                if($errorMessage -eq "")
                {
                    $errorMessage = "User blocked or dont have SPE_E3 license."
                }
            }        
        }
        catch
        {
            $status = "ERROR"
            if($errorMessage -eq "")
            {
                $errorMessage = "Cannot found user $upn in tenant."
            }
        }

        if($status -eq "step1")
        {
            $i = 0
            $haveSharepoint = $false
            $userDisabledServicePlans = @();
            foreach($servicePlan in $upnInfo.Licenses.ServiceStatus.ServicePlan.ServiceName)
            {
                if(($servicePlan -eq "SHAREPOINTENTERPRISE"))
                {
                    if(($upnInfo.Licenses.ServiceStatus.ProvisioningStatus[$i] -eq "Disabled"))
                    {
                        $haveSharepoint = $false
                    }
                    else
                    {
                        $haveSharepoint = $true
                        $status = "step2"
                    }            
                }
                else
                {
                    if($upnInfo.Licenses.ServiceStatus.ProvisioningStatus[$i] -eq "Disabled")
                    {
                        $userDisabledServicePlans += $servicePlan
                    }
                }        

                $i++
            }
    
            if($haveSharepoint -eq $false)
            {
                try
                {
                    #Create list of ServicePlans that should be disabled
                    #$disabledPlans = (New-MsolLicenseOptions -AccountSkuId "globalvale:ENTERPRISEPACK" -DisabledPlans $userDisabledServicePlans)
                    $disabledPlans = (New-MsolLicenseOptions -AccountSkuId "globalvale:SPE_E3" -DisabledPlans $userDisabledServicePlans)

                    #Provision E3 license and disable list of ServicePlans
                    Set-MsolUserLicense -UserPrincipalName $upnInfo.UserPrincipalName -LicenseOptions $disabledPlans
            
                    $status = "step2"
                }
                catch
                {
                    $status = "ERROR"
                    if($errorMessage -eq "")
                    {
                        $errorMessage = "Cannot update user license with SHAREPOINTENTERPRISE"
                    }
                }
            }
            else
            {
                $status = "step2"
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

sleep(80)


#Baixa informações do banco de dados
$data = GetTickets('GrantOneDrive')

foreach($user in $data)
{
    $WO = $user.WO
    $UPN = $user.UPN_Add    
    [datetime]$sbDate = $user.SubmitDate
    $status = $user.Status
    
    $upnInfo = ""    
    $errorMessage = ""
    $closeNote = ""
    $SPOSite = "" 

    if($status -eq "step2")
    {
        $upnInfo = Get-MsolUser -UserPrincipalName $UPN

        $personalSite = "https://globalvale-my.sharepoint.com/personal/" + $upnInfo.UserPrincipalName.replace('.', '_').replace('@', '_').toLower()

        try
        {
            $SPOSite = Get-SpoSite -Identity $personalSite
        }
        catch
        {
            $SPOSite = $null
        }

        if($SPOSite -eq $null)
        {
            try
            {
                Request-SPOPersonalSite -UserEmails @($upnInfo.UserPrincipalName)
                $status = "Success"
                $closeNote = "RESOLUTION NOTE:|jump||jump||jump|Caro usuário(a)|jump||jump|O OneDrive for Business foi habilitado para sua chave " + $UPN + ". Recomendamos o uso desta nova funcionalidade para armazenar seus arquivos de trabalho na nuvem trabalhar em documentos de forma colaborativa com sua equipe e compartilhar documentos que não sejam versões finais.|jump||jump|Para mais informações vídeos treinamentos guias rápidos e suporte da TI visite nossa página do OneDrive for Business na Intranet copie e cole o link no Internet Explorer (http://intranet.valepub.net/pt/Paginas/tecnologia-da-informacao/ferramentas-de-comunicacao-e-colaboracao/onedrive/onedrive.aspx).|jump||jump|O OneDrive for Business está disponível através do navegador da Web acessando https://portal.office.com/. Você deverá fazer o login no portal usando sua credencial de funcionário Vale.|jump||jump|Att|jump||jump|O365 team.|jump||jump||jump|******************************************************************************************************************************|jump||jump||jump|RESOLUTION NOTE:|jump||jump||jump|Dear user|jump||jump|OneDrive for Business has been enabled for your Vale employee credential " + $UPN + ". We recommend using this new feature to store your documents in the cloud work on documents in collaboration with your team and share documents that are not final versions of projects or teams.|jump||jump|For more information videos trainings quick guides and how to get IT support visit our OneDrive for Business page on Intranet copy and past the link on Internet Explorer (http://intranet.valepub.net/en/Pages/tecnologia-da-informacao/ferramentas-de-comunicacao-e-colaboracao/onedrive/onedrive.aspx).|jump||jump||jump|OneDrive for Business is available via web browser by going to https://portal.office.com/. You must log in using your Vale employee credential.|jump||jump||jump|Regards|jump||jump|O365 team."
            }
            catch
            {
                $status = "ERROR"
                if($errorMessage -eq "")
                {
                    $errorMessage = "Error while requesting PersonalSite"
                }
            }
        }
        else 
        {
            if ($SPOSite.LockState -eq "NoAccess")
            {
                try
                {
                    $SPOSite | Set-SpoSite -LockState Unlock
                    $status = "Success"
                    $closeNote = "RESOLUTION NOTE:|jump||jump||jump|Caro usuário(a)|jump||jump|O OneDrive for Business foi habilitado para sua chave " + $UPN + ". Recomendamos o uso desta nova funcionalidade para armazenar seus arquivos de trabalho na nuvem trabalhar em documentos de forma colaborativa com sua equipe e compartilhar documentos que não sejam versões finais.|jump||jump|Para mais informações vídeos treinamentos guias rápidos e suporte da TI visite nossa página do OneDrive for Business na Intranet copie e cole o link no Internet Explorer (http://intranet.valepub.net/pt/Paginas/tecnologia-da-informacao/ferramentas-de-comunicacao-e-colaboracao/onedrive/onedrive.aspx).|jump||jump|O OneDrive for Business está disponível através do navegador da Web acessando https://portal.office.com/. Você deverá fazer o login no portal usando sua credencial de funcionário Vale.|jump||jump|Att|jump||jump|O365 team.|jump||jump||jump|******************************************************************************************************************************|jump||jump||jump|RESOLUTION NOTE:|jump||jump||jump|Dear user|jump||jump|OneDrive for Business has been enabled for your Vale employee credential " + $UPN + ". We recommend using this new feature to store your documents in the cloud work on documents in collaboration with your team and share documents that are not final versions of projects or teams.|jump||jump|For more information videos trainings quick guides and how to get IT support visit our OneDrive for Business page on Intranet copy and past the link on Internet Explorer (http://intranet.valepub.net/en/Pages/tecnologia-da-informacao/ferramentas-de-comunicacao-e-colaboracao/onedrive/onedrive.aspx).|jump||jump||jump|OneDrive for Business is available via web browser by going to https://portal.office.com/. You must log in using your Vale employee credential.|jump||jump||jump|Regards|jump||jump|O365 team."
                }
                catch
                {
                    $status = "ERROR"
                    if($errorMessage -eq "")
                    {
                        $errorMessage = "Error while unlocking SPOSite"
                    }
                }
            }

            if($SPOSite.LockState -eq "Unlock")
            {
                $status = "Success"
                $closeNote = "RESOLUTION NOTE:|jump||jump||jump|Caro usuário(a)|jump||jump|O OneDrive for Business foi habilitado para sua chave " + $UPN + ". Recomendamos o uso desta nova funcionalidade para armazenar seus arquivos de trabalho na nuvem trabalhar em documentos de forma colaborativa com sua equipe e compartilhar documentos que não sejam versões finais.|jump||jump|Para mais informações vídeos treinamentos guias rápidos e suporte da TI visite nossa página do OneDrive for Business na Intranet copie e cole o link no Internet Explorer (http://intranet.valepub.net/pt/Paginas/tecnologia-da-informacao/ferramentas-de-comunicacao-e-colaboracao/onedrive/onedrive.aspx).|jump||jump|O OneDrive for Business está disponível através do navegador da Web acessando https://portal.office.com/. Você deverá fazer o login no portal usando sua credencial de funcionário Vale.|jump||jump|Att|jump||jump|O365 team.|jump||jump||jump|******************************************************************************************************************************|jump||jump||jump|RESOLUTION NOTE:|jump||jump||jump|Dear user|jump||jump|OneDrive for Business has been enabled for your Vale employee credential " + $UPN + ". We recommend using this new feature to store your documents in the cloud work on documents in collaboration with your team and share documents that are not final versions of projects or teams.|jump||jump|For more information videos trainings quick guides and how to get IT support visit our OneDrive for Business page on Intranet copy and past the link on Internet Explorer (http://intranet.valepub.net/en/Pages/tecnologia-da-informacao/ferramentas-de-comunicacao-e-colaboracao/onedrive/onedrive.aspx).|jump||jump||jump|OneDrive for Business is available via web browser by going to https://portal.office.com/. You must log in using your Vale employee credential.|jump||jump||jump|Regards|jump||jump|O365 team."
            }
        }
    }

    UpdateTicket $WO $Status $errorMessage  $closeNote
}