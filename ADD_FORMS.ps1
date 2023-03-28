#importa funcoes de banco de dados
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

#Baixa informações do banco de dados
$data = GetTickets('ADD_Forms')

foreach ($row in $data) 
{
    $WO = $row.WO
    $UPN = $row.UPN_Add  
    
    
    [datetime] $sbDate = $row.SubmitDate
    [datetime] $ticketDate = $row.BotLastUpdate
    [datetime] $currentDate = Get-Date
    $status = $row.Status

    $elapsedTime =  ($currentDate - $ticketDate).TotalMinutes
   
    ## tratamento de tickets com mais de 20 horas ##########################
    # if($elapsedTime -gt  1200) # 1200 = 20*60
    # {
    #     $status = "ERROR"
    #     $errorMessage = "SLA breach risk. Ticket awaiting more than 20 hours."
    # }
    ########################################################################

    
    
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
        #Check if have E3
        try
        {
            $upnInfo = Get-MsolUser -UserPrincipalName $UPN

            #Declare variable and flags
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
                    #if ($userLicense[$i] -eq "globalvale:ENTERPRISEPACK")
                    if ($userLicense -eq "globalvale:SPE_E3") #check if the only licenses is the E3
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
            $upnInfo = Get-MsolUser -UserPrincipalName $UPN

            $i = 0
            $haveForms = $false
            $userDisabledServicePlans = @();
            foreach($servicePlan in $upnInfo.Licenses.ServiceStatus.ServicePlan.ServiceName)
            {
                if(($servicePlan -eq "FORMS_PLAN_E3"))
                {
                    if(($upnInfo.Licenses.ServiceStatus.ProvisioningStatus[$i] -eq "Disabled"))
                    {
                        $haveForms = $false
                    }
                    else
                    {
                        $haveForms = $true
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
    
            if($haveForms -eq $false)
            {
                try
                {
                    #Create list of ServicePlans that should be disabled
                    #$disabledPlans = (New-MsolLicenseOptions -AccountSkuId "globalvale:ENTERPRISEPACK" -DisabledPlans $userDisabledServicePlans)
                    $disabledPlans = (New-MsolLicenseOptions -AccountSkuId "globalvale:SPE_E3" -DisabledPlans $userDisabledServicePlans)
                    #Provision E3 license and disable list of ServicePlans
                    Set-MsolUserLicense -UserPrincipalName $upnInfo.UserPrincipalName -LicenseOptions $disabledPlans
            
                    $status = "Success"
                    $closeNote = "Caro usuário(a),|jump||jump|Conforme solicitado, a licença de FORMS foi ativada para sua conta.|jump||jump|O FORMS está disponível através do navegador da Web, acessando https://portal.office.com/, você deverá fazer o login no portal usando o seu e-mail e senha.|jump||jump|Caso o icone do FORMS não apareça na primeira pagina, basta clicar em Todos os Aplicativos para visualizar.|jump||jump|Favor aguardar 1 hora após recebimento desse e-mail para acesso.|jump||jump|Att,|jump|O365 Team.|jump||jump|********************************************************************************************************|jump||jump|Dear,|jump||jump|As requested, the FORMS license has been activated for your account.|jump||jump|FORMS is available through the web browser, accessing https://portal.office.com/, you must login to the portal using your email and password.|jump||jump|If the FORMS icon does not appear on the first page, just click on All Applications to view.|jump||jump|Please wait 1 hour after receiving this email to access.|jump||jump|Best regards,|jump|O365 Team.|jump||jump|********************************************************************************************************|jump||jump|Estimado(a),|jump||jump|Como se solicitó, la licencia de FORMS se ha activado para su cuenta.|jump||jump|FORMS está disponible a través del navegador web, accediendo a https://portal.office.com/, debe iniciar sesión en el portal con su correo electrónico y contraseña.|jump||jump|Si el icono de FORMS no aparece en la primera página, simplemente haga clic en Todas las aplicaciones para verlo.|jump||jump|Espere 1 hora después de recibir este correo electrónico para acceder.|jump||jump|Atentamente,|jump|Equipo O365"
                }
                catch
                {
                    $status = "ERROR"
                    if($errorMessage -eq "")
                    {
                        $errorMessage = "Cannot update user license with FORMS_PLAN_E3"
                    }
                }
            }
            else
            {
                $status = "Success"
                $errorMessage = "user already have FORMS_PLAN_E3"
                $closeNote = "Caro usuário(a),|jump||jump|Conforme solicitado, a licença de FORMS foi ativada para sua conta.|jump||jump|O FORMS está disponível através do navegador da Web, acessando https://portal.office.com/, você deverá fazer o login no portal usando o seu e-mail e senha.|jump||jump|Caso o icone do FORMS não apareça na primeira pagina, basta clicar em Todos os Aplicativos para visualizar.|jump||jump|Favor aguardar 1 hora após recebimento desse e-mail para acesso.|jump||jump|Att,|jump|O365 Team.|jump||jump|********************************************************************************************************|jump||jump|Dear,|jump||jump|As requested, the FORMS license has been activated for your account.|jump||jump|FORMS is available through the web browser, accessing https://portal.office.com/, you must login to the portal using your email and password.|jump||jump|If the FORMS icon does not appear on the first page, just click on All Applications to view.|jump||jump|Please wait 1 hour after receiving this email to access.|jump||jump|Best regards,|jump|O365 Team.|jump||jump|********************************************************************************************************|jump||jump|Estimado(a),|jump||jump|Como se solicitó, la licencia de FORMS se ha activado para su cuenta.|jump||jump|FORMS está disponible a través del navegador web, accediendo a https://portal.office.com/, debe iniciar sesión en el portal con su correo electrónico y contraseña.|jump||jump|Si el icono de FORMS no aparece en la primera página, simplemente haga clic en Todas las aplicaciones para verlo.|jump||jump|Espere 1 hora después de recibir este correo electrónico para acceder.|jump||jump|Atentamente,|jump|Equipo O365"
            }
        }
    }
       
    UpdateTicket $WO $Status $errorMessage  $closeNote
}
