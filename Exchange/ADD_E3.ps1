#Lista oficial de países
$CountryHashTable = @{ `
    "Afghanistan" = "AF"; `
    "Åland Islands" = "AX"; `
    "Albania" = "AL"; `
    "Algeria" = "DZ"; `
    "American Samoa" = "AS"; `
    "Andorra" = "AD"; `
    "Angola" = "AO"; `
    "Anguilla" = "AI"; `
    "Antarctica" = "AQ"; `
    "Antigua and Barbuda" = "AG"; `
    "Argentina" = "AR"; `
    "Armenia" = "AM"; `
    "Aruba" = "AW"; `
    "Australia" = "AU"; `
    "Austria" = "AT"; `
    "Azerbaijan" = "AZ"; `
    "Bahamas" = "BS"; `
    "Bahrain" = "BH"; `
    "Bangladesh" = "BD"; `
    "Barbados" = "BB"; `
    "Belarus" = "BY"; `
    "Belgium" = "BE"; `
    "Belize" = "BZ"; `
    "Benin" = "BJ"; `
    "Bermuda" = "BM"; `
    "Bhutan" = "BT"; `
    "Bolivia" = "BO"; `
    "Bonaire, Sint Eustatius and Saba" = "BQ"; `
    "Bosnia and Herzegovina" = "BA"; `
    "Botswana" = "BW"; `
    "Bouvet Island" = "BV"; `
    "Brasil" = "BR"; `
    "Brazil" = "BR"; `
    "British Indian Ocean Territory" = "IO"; `
    "Brunei Darussalam" = "BN"; `
    "Bulgaria" = "BG"; `
    "Burkina Faso" = "BF"; `
    "Burundi" = "BI"; `
    "Cabo Verde" = "CV"; `
    "Cambodia" = "KH"; `
    "Cameroon" = "CM"; `
    "Canada" = "CA"; `
    "Cayman Islands" = "KY"; `
    "Central African Republic" = "CF"; `
    "Chad" = "TD"; `
    "Chile" = "CL"; `
    "China" = "CN"; `
    "Christmas Island" = "CX"; `
    "Cocos (Keeling) Islands" = "CC"; `
    "Colombia" = "CO"; `
    "Comoros" = "KM"; `
    "Congo" = "CG"; `
    "Congo (DRC)" = "CD"; `
    "Cook Islands" = "CK"; `
    "Costa Rica" = "CR"; `
    "Côte d'Ivoire" = "CI"; `
    "Croatia" = "HR"; `
    "Cuba" = "CU"; `
    "Curaçao" = "CW"; `
    "Cyprus" = "CY"; `
    "Czech Republic" = "CZ"; `
    "Denmark" = "DK"; `
    "Djibouti" = "DJ"; `
    "Dominica" = "DM"; `
    "Dominican Republic" = "DO"; `
    "Ecuador" = "EC"; `
    "Egypt" = "EG"; `
    "El Salvador" = "SV"; `
    "Equatorial Guinea" = "GQ"; `
    "Eritrea" = "ER"; `
    "Estonia" = "EE"; `
    "Ethiopia" = "ET"; `
    "Falkland Islands (Malvinas)" = "FK"; `
    "Faroe Islands" = "FO"; `
    "Fiji" = "FJ"; `
    "Finland" = "FI"; `
    "France" = "FR"; `
    "French Guiana" = "GF"; `
    "French Polynesia" = "PF"; `
    "French Southern Territories" = "TF"; `
    "Gabon" = "GA"; `
    "Gambia" = "GM"; `
    "Georgia" = "GE"; `
    "Germany" = "DE"; `
    "Ghana" = "GH"; `
    "Gibraltar" = "GI"; `
    "Greece" = "GR"; `
    "Greenland" = "GL"; `
    "Grenada" = "GD"; `
    "Guadeloupe" = "GP"; `
    "Guam" = "GU"; `
    "Guatemala" = "GT"; `
    "Guernsey" = "GG"; `
    "Guinea" = "GN"; `
    "Guinea-Bissau" = "GW"; `
    "Guyana" = "GY"; `
    "Haiti" = "HT"; `
    "Heard Island and McDonald Islands" = "HM"; `
    "Holy See (Vatican City State)" = "VA"; `
    "Honduras" = "HN"; `
    "Hong Kong" = "HK"; `
    "Hungary" = "HU"; `
    "Iceland" = "IS"; `
    "India" = "IN"; `
    "Indonesia" = "ID"; `
    ### Not currently available as a usage location in Office 365 ### "Iran (the Islamic Republic of)" = "IR"; `
    "Iraq" = "IQ"; `
    "Ireland" = "IE"; `
    "Isle of Man" = "IM"; `
    "Israel" = "IL"; `
    "Italy" = "IT"; `
    "Jamaica" = "JM"; `
    "Japan" = "JP"; `
    "Jersey" = "JE"; `
    "Jordan" = "JO"; `
    "Kazakhstan" = "KZ"; `
    "Kenya" = "KE"; `
    "Kiribati" = "KI"; `
    ### Not currently available as a usage location in Office 365 ### "Korea (the Democratic People's Republic of)" = "KP"; `
    "Korea, Republic of" = "KR"; `
    "Kuwait" = "KW"; `
    "Kyrgyzstan" = "KG"; `
    "Lao People's Democratic Republic" = "LA"; `
    "Latvia" = "LV"; `
    "Lebanon" = "LB"; `
    "Lesotho" = "LS"; `
    "Liberia" = "LR"; `
    "Libya" = "LY"; `
    "Liechtenstein" = "LI"; `
    "Lithuania" = "LT"; `
    "Luxembourg" = "LU"; `
    "Macao" = "MO"; `
    "Macedonia, the former Yugoslav Republic of" = "MK"; `
    "Madagascar" = "MG"; `
    "Malawi" = "MW"; `
    "Malaysia" = "MY"; `
    "Maldives" = "MV"; `
    "Mali" = "ML"; `
    "Malta" = "MT"; `
    "Marshall Islands" = "MH"; `
    "Martinique" = "MQ"; `
    "Mauritania" = "MR"; `
    "Mauritius" = "MU"; `
    "Mayotte" = "YT"; `
    "Mexico" = "MX"; `
    "Micronesia" = "FM"; `
    "Moldova" = "MD"; `
    "Monaco" = "MC"; `
    "Mongolia" = "MN"; `
    "Montenegro" = "ME"; `
    "Montserrat" = "MS"; `
    "Morocco" = "MA"; `
    "Mozambique" = "MZ"; `
    ### Not currently available as a usage location in Office 365 ### "Myanmar" = "MM"; `
    "Namibia" = "NA"; `
    "Nauru" = "NR"; `
    "Nepal" = "NP"; `
    "Netherlands" = "NL"; `
    "The Netherlands" = "NL"; `
    "New Caledonia" = "NC"; `
    "New Zealand" = "NZ"; `
    "Nicaragua" = "NI"; `
    "Niger" = "NE"; `
    "Nigeria" = "NG"; `
    "Niue" = "NU"; `
    "Norfolk Island" = "NF"; `
    "Northern Mariana Islands" = "MP"; `
    "Norway" = "NO"; `
    "Oman" = "OM"; `
    "Pakistan" = "PK"; `
    "Palau" = "PW"; `
    "Palestine, State of" = "PS"; `
    "Panama" = "PA"; `
    "Papua New Guinea" = "PG"; `
    "Paraguay" = "PY"; `
    "Peru" = "PE"; `
    "Philippines" = "PH"; `
    "Pitcairn" = "PN"; `
    "Poland" = "PL"; `
    "Portugal" = "PT"; `
    "Puerto Rico" = "PR"; `
    "Qatar" = "QA"; `
    "Réunion" = "RE"; `
    "Romania" = "RO"; `
    "Russian Federation" = "RU"; `
    "Rwanda" = "RW"; `
    "Saint Barthélemy" = "BL"; `
    "Saint Helena, Ascension and Tristan da Cunha" = "SH"; `
    "Saint Kitts and Nevis" = "KN"; `
    "Saint Lucia" = "LC"; `
    "Saint Martin" = "MF"; `
    "Saint Pierre and Miquelon" = "PM"; `
    "Saint Vincent and the Grenadines" = "VC"; `
    "Samoa" = "WS"; `
    "San Marino" = "SM"; `
    "Sao Tome and Principe" = "ST"; `
    "Saudi Arabia" = "SA"; `
    "Senegal" = "SN"; `
    "Serbia" = "RS"; `
    "Seychelles" = "SC"; `
    "Sierra Leone" = "SL"; `
    "Singapore" = "SG"; `
    "Sint Maarten" = "SX"; `
    "Slovakia" = "SK"; `
    "Slovenia" = "SI"; `
    "Solomon Islands" = "SB"; `
    "Somalia" = "SO"; `
    "South Africa" = "ZA"; `
   "South Georgia and the South Sandwich Islands" = "GS"; `
    "South Sudan " = "SS"; `
    "Spain" = "ES"; `
    "Sri Lanka" = "LK"; `
    "Sudan" = "SD"; `
    "Suriname" = "SR"; `
    "Svalbard and Jan Mayen" = "SJ"; `
    "Swaziland" = "SZ"; `
    "Sweden" = "SE"; `
    "Switzerland" = "CH"; `
    "Syrian Arab Republic" = "SY"; `
    "Taiwan" = "TW"; `
    "Tajikistan" = "TJ"; `
    "Tanzania" = "TZ"; `
    "Thailand" = "TH"; `
    "Timor-Leste" = "TL"; `
    "Togo" = "TG"; `
    "Tokelau" = "TK"; `
    "Tonga" = "TO"; `
    "Trinidad and Tobago" = "TT"; `
    "Tunisia" = "TN"; `
    "Turkey" = "TR"; `
    "Turkmenistan" = "TM"; `
    "Turks and Caicos Islands" = "TC"; `
    "Tuvalu" = "TV"; `
    "Uganda" = "UG"; `
    "Ukraine" = "UA"; `
    "United Arab Emirates" = "AE"; `
    "United Kingdom" = "GB"; `
    "United States" = "US"; `
    "United States Minor Outlying Islands" = "UM"; `
    "Uruguay" = "UY"; `
    "Uzbekistan" = "UZ"; `
    "Vanuatu" = "VU"; `
    "Venezuela, Bolivarian Republic of" = "VE"; `
    "Viet Nam" = "VN"; `
    "Virgin Islands, British" = "VG"; `
    "Virgin Islands, U.S." = "VI"; `
    "Wallis and Futuna" = "WF"; `
    "Western Sahara*" = "EH"; `
    "Yemen" = "YE"; `
    "Zambia" = "ZM"; `
    "Zimbabwe" = "ZW"; `
    };

$GroupErrorLog = @()



$ValidTicketStatus = @("step 0","step0","WaitNextRun","WaitLastRun")

function GetUserSPE3LicenseStatus($upn)
{
    try
    {
        $PSError = ""
        $result = $false
        $LicenseDetails = (Get-MsolUser -UserPrincipalName $upn).Licenses
        if ($LicenseDetails.AccountSkuID -contains "globalvale:SPE_E3")
        {
            $result = $true
        }
        return $result
     }
     catch
     {
        Write-Host ($PSError)
     }
}


#importa funcoes de banco de dados
Import-Module "C:\Automation\TTR_Vale\PowerShell\D_Database_Modules.psm1"

#Baixa informações do banco de dados
$data = GetTickets('ADD_E3')


##Groups ###############################################################
$Vale_Lic_M365E3_Standard = "bea5756a-862a-43ad-81fe-11f1ee463a85"
$Vale_Lic_M365E3_SPO = "e1fe16ca-f96c-45e6-b397-e1c1f35b962a"
#$Vale_Talk360 = "24abbc28-9b53-41a8-b2ad-bd6d0b49a28b"
#$Vale_Safety_and_Operational_Risk = "4f9ccd36-7ba3-4815-b6c3-cfea5610ef5a"
$Vale_Jornal = "a93f0ba3-b384-46ae-a454-dc3ad0572a99"
#########################################################################

$BrazilGroups = @(
    $Vale_Lic_M365E3_Standard,
    $Vale_Lic_M365E3_SPO,
    #$Vale_Talk360,
    #$Vale_Safety_and_Operational_Risk,
    $Vale_Jornal)


$NcIdOmZaGroups = @(
    $Vale_Lic_M365E3_Standard,
    #$Vale_Talk360,
    #$Vale_Safety_and_Operational_Risk,
    $Vale_Jornal )

$GeneralGroups =@(
    $Vale_Lic_M365E3_Standard ,
    $Vale_Lic_M365E3_SPO ,
    #$Vale_Talk360,
    #$Vale_Safety_and_Operational_Risk,
    $Vale_Jornal  )

foreach ($ticket in $data)
{
    $WO = $ticket.WO
    $UPN = $ticket.UPN_Add    
   
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
    #        $PSErrorMessage = "SLA breach risk. Ticket awaiting more than 20 hours."
    #        UpdateTicket $WO $Status $PSErrorMessage  $closeNote
    #        continue
    #    }
    ###############################################################################

    

        $upnInfo = ""
        [datetime]$nowDate = Get-Date
        $PSErrorMessage = ""

        if($WO -eq "")
        {
            $WO = "ERROR"
            $UPN = "ERROR"
            $status = "ERROR"
            $upnInfo = "ERROR"
            $PSErrorMessage = "ERROR"
            $closeNote = "ERROR"
        }

        #if (($status -eq "step 0") -or ($status -eq "step0") -or ($status -eq "WaitNextRun") -or ($status -eq "WaitLastRun") )
        if ($ValidTicketStatus -contains $status  )
        {
                #Check if user exists
                try
                {
                    $upnInfo = Get-Mailbox -Identity $UPN -ResultSize 5
                    $UPN = $upnInfo.userprincipalname

                    if (($UPN -ne '') -and ($UPN -ne $null))
                    {
                        $userINFO = Get-MsolUser -UserPrincipalName $UPN
                        $userGUID = $userINFO.Objectid.Guid

                        $status = "step1"
                    }
                    else
                    {
                        if ($status -eq "WaitNextRun")
                        {
                            $status = "WaitLastRun"
                            if($PSErrorMessage -eq "")
                            {
                                $PSErrorMessage = "User still havent synchronized, waiting last run."
                            }
                        }
                        elseif ($status -eq "WaitLastRun")
                        {
                            $status = "ERROR"
                            if($PSErrorMessage -eq "")
                            {
                                $PSErrorMessage = "Cannot found user $upn in tenant."
                            }
                        }
                        else
                        {
                            $status = "WaitNextRun"
                            if($PSErrorMessage -eq "")
                            {
                                $PSErrorMessage = "User probably aint synchronized, waiting next run."
                            }
                        }

                        UpdateTicket $WO $Status $PSErrorMessage  $closeNote
                    }
                }
                catch
                {
                    if($nowDate -le $sbDate.AddHours(5))
                    {
                        $status = "WaitNextRun"
                        if($PSErrorMessage -eq "")
                        {
                            $PSErrorMessage = "User probably not synchronized, waiting next run."
                        }
                    }
                    else
                    {
                        $status = "ERROR"
                        if($PSErrorMessage -eq "")
                        {
                            $PSErrorMessage = "Cannot find user $upn in tenant."
                        }
                    }
                }
            #}

            #If user exists in tenant, check and set user location
            if($status -eq "step1")
            {
                try
                {
                    ###Estrutura para obter GUID do usuári via AAD (Get-MsolUser -UserPrincipalName AdeleV@fmcao365.tk | fl Objectid )
                    #$userGUID = Get-MsolUser -UserPrincipalName $UPN | fl Objectid
                    # trying to match the country value with a two letter code country, skiping the user if no match was found
                    if ($upnInfo.usagelocation -eq $null)
                    {
                        Set-MsolUser -UserPrincipalName $upnInfo.UserPrincipalName -UsageLocation($CountryHashTable.Item($upnInfo.CustomAttribute7))
                    }
            
                    $status = "step2"
                }
                catch
                {
                    $status = "ERROR"
                    if($PSErrorMessage -eq "")
                    {
                        $PSErrorMessage = "Failed to get/set usage location"
                    }
                }

                try
                {
                    $LicenseDetails = (Get-MsolUser -UserPrincipalName $UPN).Licenses
    
                    if ($LicenseDetails.AccountSkuID -contains "globalvale:TEAMS_COMMERCIAL_TRIAL")
                    {
                        If ($LicenseDetails.AccountSkuID -contains "globalvale:MS_TEAMS_IW")
                        {
                            Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "globalvale:MS_TEAMS_IW" 
                            Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "globalvale:TEAMS_COMMERCIAL_TRIAL"
                            $status = "step2.1"
                        }
                        else
                        {
                            Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "globalvale:TEAMS_COMMERCIAL_TRIAL"
                            $status = "step2.1"
                        }
                    }
                    else
                    {
                        If ($LicenseDetails.AccountSkuID -contains "globalvale:MS_TEAMS_IW")
                        {
                            Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses "globalvale:MS_TEAMS_IW"
                            $status = "step2.1" 
                        }
                    }

                    $status = "step2.1" 
                }
                catch
                {
                    $status = "ERROR"
                    if($PSErrorMessage -eq "")
                    {
                        $PSErrorMessage = "Failed to remove teams trials"
                    }
                }
            }
    
            if($status -eq "step2.1")
            {
                try
                {
                    $Error.Clear()
                    if($upnInfo.CustomAttribute7 -eq "Brazil")
                    {
                        
                        foreach($group in $BrazilGroups)
                        {
                            try
                            {
                                Add-AzureADGroupMember -ObjectId $group -RefObjectId $userGUID
                            }
                            catch
                            {
                                if ( $error[0].Exception.Message.Contains("already exist"))
                                {
                                    continue
                                }
                            }
                        }
                 
                    }

                    elseif($upnInfo.CustomAttribute7 -eq "New Caledonia" -or  $upnInfo.CustomAttribute7 -eq "Indonesia" -or $upnInfo.CustomAttribute7 -eq "Oman" -or $upnInfo.CustomAttribute7 -eq "South Africa")
                    {
                        $Error.Clear()

                        foreach($group in $NcIdOmZaGroups)
                        {
                            try
                            {
                                Add-AzureADGroupMember -ObjectId $group -RefObjectId $userGUID
                            }
                            catch
                            {
                                if ( $error[0].Exception.Message.Contains("already exist"))
                                {
                                    continue
                                }
                            }
                        }
                       
                    }
                    else
                    {
                        foreach($group in $GeneralGroups)
                        {
                            try
                            {
                                Add-AzureADGroupMember -ObjectId $group -RefObjectId $userGUID
                            }
                            catch
                            {
                                if ( $error[0].Exception.Message.Contains("already exist"))
                                {
                                    continue
                                }
                            }
                        }
                    }
                   
                    $status = "step5"
                    UpdateTicket $WO $Status $PSErrorMessage  $closeNote
                }
                catch
                {
                  
                    $status = "ERROR"
                    if($PSErrorMessage -eq "")
                    {
                        $PSErrorMessage = "Failed to set one or more groups"

                    }
                }
            }
        }
}


Start-Sleep(60) 

#Baixa informações do banco de dados
$data = GetTickets('ADD_E3')

foreach ($user in $data)
{
    $WO = $user.WO
    $UPN = $user.UPN_Add
    $status = $user.Status
    $PSErrorMessage = $user.ErrorMessage
    $upnName = ""
    $displayName = ""
    $upnAlias = ""
    $closeNote = ""
    $state = $user.state

    if ($status -ne "Step5")
    {
        continue
    }
        try
        {
            $userData = Get-Mailbox -Identity $UPN
            $UPN = $userData.UserPrincipalName
            $userLocation = $CountryHashTable.Item($userData.CustomAttribute7)
            
           ##  verifica se a licença aplicada está ativa ########################################
            $UserHasSPEE3 = GetUserSPE3LicenseStatus($UPN)
            if($UserHasSPEE3 -eq $false)
            {
                $status = "ERROR"
                $PSErrorMessage = "User added to groups but no MS365 E3 license has been applied"
                UpdateTicket $WO $Status $PSErrorMessage  $closeNote
                continue
            }
            #####################################################################################
            
                if(($userLocation -ne "OM") -and ($userLocation -ne "NC") -and ($userLocation -ne "MZ") -and ($state -ne "Minas Gerais"))
                {
                    Grant-CsTeamsMeetingPolicy -Identity $UPN -PolicyName "Teams - Covid19 Crisis"
                }
                elseif($state -eq "Minas Gerais")
                {
                    Grant-CsTeamsMeetingPolicy -Identity $UPN -PolicyName "Teams – Covid 19 Crisis 1Mbps"
                }
                else
                {
                    Grant-CsTeamsMeetingPolicy -Identity $UPN -PolicyName "Teams - Global Policy"
                }
                
                Grant-CsTeamsMessagingPolicy -Identity $UPN -PolicyName ""
                Grant-CsTeamsAppSetupPolicy -Identity $UPN -PolicyName "Standard"
                Grant-CsTeamsAppPermissionPolicy -Identity $UPN -PolicyName "Standard"

                Set-Mailbox -identity $UPN -AuditEnabled $true
                Set-CASMailbox -identity $UPN -POPEnabled $false -ImapEnabled $false
                Set-CASMailbox -Identity $UPN -ActiveSyncEnabled $false

                $status = "Step6" #[MAU]status criado para executar o powershell Enable_E3_SharePoint.ps1 que é complemento do provisionamento de E3

                $upnName = (Get-MsolUser -UserPrincipalName $UPN).FirstName
                $displayName = (Get-MsolUser -UserPrincipalName $UPN).DisplayName
                $upnAlias = (Get-Mailbox -Identity $UPN).Alias
                
                $aliasTag = "Alias"
                $displayNameTag = "DisplayName"
                $closeNote = "Sr(a) " + $upnName + "|jump||jump|" + "Conforme solicitado no chamado " + $WO + "; foi  criada chave de acesso ao Correio, conforme abaixo:|jump||jump|Name: " + $upnName + "|jump|"+ $displaynameTag +": " + $displayName + "|jump|UserPrincipalName: " + $UPN + "|jump|"+$aliasTag+": " + $upnAlias + "|jump||jump|-> Para acessar ao Outlook; executar os seguintes procedimentos:|jump|01.  Inicie o aplicativo Outlook;|jump|02.  Uma tela de boas-vindas irá aparecer. clique em próximo;|jump|03.  Suas credenciais devem aparecer automaticamente. Clique em próximo.|jump|04.  Uma tela de confimação irá aparecer. Clique em concluir|jump|05.  Entre no Outlook e use-o normalmente.|jump||jump|-> Para acesso ao Outlook via web, seguir os seguintes passos :|jump|01.   Abra o Internet Explorer e acesse o link: outlook.office365.com\owa|jump|02.   No campo Dominio\Nome de usuário, digite o seu e-mail e aperte a tecla tab|jump|03.   Uma nova tela de autenticação será aberta. Insira sua matrícula no campo usuário e sua senha no campo senha|jump|04.   Clique em Entrar.|jump|Favor aguardar 1 hora após recebimento desse email para acesso ao novo e-mail. Em relação ao OneDrive For Business, a Microsoft pede até 48 horas para ativação completa.|jump|Caso tenha alguma dúvida ou problema, por favor contate o Service Desk.|jump||jump|O365 Team|jump||jump|Mr. /Mrs." + $upnName+ ";|jump||jump|" + "As requested on ticket " + $WO + "; the mail access key for the user was created:|jump||jump|Name: " + $upnName + "|jump|DisplayName: " + $displayName + "|jump|UserPrincipalName: " + $UPN + "|jump|Alias: " + $upnAlias + "|jump||jump|Please log off and log in before testing the access.|jump|-> To access Outlook; follow the procedure below:|jump|01. Execute the outlook application|jump|02. A welcome screen will appear. Press next button|jump|03. Your e-mail and credentials should appear automatically. Then click next.|jump|04. A confirmation screen will appear. Click finish.|jump|05. Open your outlook and use your e-mail normally.|jump||jump|-> To access Outlook via web; follow the steps below:|jump|01. Launch Internet Explorer and access the link: outlook.office365.com\owa|jump|02. In Domain \ User name  type your e-mail and press tab button on the keyboard|jump|03. Another authentication screen will appear. In the fields username and password insert your Vale ID and password respectively;|jump|04. Click Login.|jump||jump|Please wait 1 hour after receiving this email to access new mail. Related to the OneDrive For Business, Microsoft ask until 48 hours to have it complete active.|jump||jump|If you have any questions or problems; please contact the Service Desk.|jump||jump|O365 Team"
                $closeNote = $closeNote.Replace("'", " ")
                $closeNote = $closeNote.Replace(",", " ")
        }
        catch
        {
            $status = "ERROR"
            if($PSErrorMessage -eq "")
            {
                $PSErrorMessage = "Failed to set policies"
            }
        }
    #}

    UpdateTicket $WO $Status $PSErrorMessage  $closeNote
}