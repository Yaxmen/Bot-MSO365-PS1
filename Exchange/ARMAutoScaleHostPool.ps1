param (
	[Parameter(mandatory = $false)] 
    [object]$WebHookData
)
# If runbook was called from Webhook, WebhookData will not be null.
if ($WebHookData) {
	# Collect properties of WebhookData
	$WebhookName = $WebHookData.WebhookName
	$WebhookHeaders = $WebHookData.RequestHeader
	$WebhookBody = $WebHookData.RequestBody
	# Collect individual headers. Input converted from JSON.
	$From = $WebhookHeaders.From
	$Json = $(ConvertFrom-Json -InputObject $WebhookBody)
} else {
	Write-Error -Message 'Runbook was not started from Webhook' -ErrorAction stop
}

$HostpoolName = $Json.hostpoolname
$BeginPeakTime = $Json.BeginPeakTime
$EndPeakTime = $Json.EndPeakTime
$TimeDifference = $Json.TimeDifference
$SessionThresholdPerCPU = $Json.SessionThresholdPerCPU
$MinimumNumberOfRDSH = $Json.MinimumNumberOfRDSH
$LimitSecondsToForceLogOffUser = $Json.LimitSecondsToForceLogOffUser
$LogOffMessageTitle = $Json.LogOffMessageTitle
$LogOffMessageBody = $Json.LogOffMessageBody
$MaintenanceTagName = $Json.MaintenanceTagName
$ResourceGroupName = $Json.ResourceGroupName
$ResourceGroupNameAutomation = $Json.ResourceGroupNameAutomation
$AutomationAccountName = $Json.AutomationAccountName
$ConnectionAssetName = $Json.ConnectionAssetName
$RunbookLogoffShutdown = $Json.RunbookLogoffShutdown

Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope Process -Force -Confirm:$false
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -Force -Confirm:$false
# Setting ErrorActionPreference to stop script execution when error occurs
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Function to convert from UTC to Local time
function Convert-UTCtoLocalTime {
	param (
		$TimeDifferenceInHours,
        $UniversalTime=$null
	)
    
    if($UniversalTime -eq $null) {
	    $UniversalTime = (Get-Date).ToUniversalTime()
    }
	$TimeDifferenceMinutes = 0
	if ($TimeDifferenceInHours -match ":") {
		$TimeDifferenceHours = $TimeDifferenceInHours.Split(":")[0]
		$TimeDifferenceMinutes = $TimeDifferenceInHours.Split(":")[1]
	} else {
		$TimeDifferenceHours = $TimeDifferenceInHours
	}
	#Azure is using UTC time, justify it to the local time
	$ConvertedTime = $UniversalTime.AddHours($TimeDifferenceHours).AddMinutes($TimeDifferenceMinutes)
	return $ConvertedTime
}

# Set AllowNewSession in Session Host
function Set-AllowNewSession {
	param (
        [string]$SessionHostName,
		[boolean]$AllowNewSession,
        [switch]$Count
	)

    try {					
        $KeepDrainMode = Update-AzWvdSessionHost -HostPoolName $HostPoolName -ResourceGroupName $ResourceGroupName -Name $SessionHostName -AllowNewSession:$AllowNewSession -ErrorAction Stop
        
        if($Count) {
            if($AllowNewSession) {
                $Script:TotalUnblocked+=1
            } else {
                $Script:TotalBlocked+=1
            }
        }
    } catch {
	    Write-Error "Unable to set AllowNewSession to $AllowNewSession on $SessionHostName with error: $($_.exception.message)" -ErrorAction Continue
    }
}

# Start the Session Host 
function Start-SessionHost {
	param (
		[string]$VMName
	)
	try {
		Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName | Start-AzVM -AsJob | Out-Null
        $Script:TotalStarted+=1
	} catch {
		Write-Error "Failed to start Azure VM: $($VMName) with error: $($_.exception.message)" -ErrorAction Continue
	}

}

# Stop the Session Host
function Stop-SessionHost {
	param (
		[string]$VMName
	)
	try {
		Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName | Stop-AzVM -Force -AsJob | Out-Null
        $Script:TotalStopped+=1
	} catch {
		Write-Error "Failed to stop Azure VM: $($VMName) with error: $($_.exception.message)" -ErrorAction Continue
	}
}

# Restart Session Host
function Restart-SessionHost {
	param (
		[string]$VMName
	)
	try {
		Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName | Restart-AzVM -NoWait | Out-Null
	} catch {
		Write-Error "Failed to restart Azure VM: $($VMName) with error: $($_.exception.message)" -ErrorAction Continue
	}
}

#Converting date time from UTC to Local
$CurrentDateTime = Convert-UTCtoLocalTime -TimeDifferenceInHours $TimeDifference

$BeginPeakDateTime = [datetime]::Parse($CurrentDateTime.ToShortDateString() + ' ' + $BeginPeakTime)
$EndPeakDateTime = [datetime]::Parse($CurrentDateTime.ToShortDateString() + ' ' + $EndPeakTime)
$IsPeakDataTime = $CurrentDateTime -ge $BeginPeakDateTime -and $CurrentDateTime -le $EndPeakDateTime

$DayOfWeek = (Get-Date).DayOfWeek

#check the calculated end time is later than begin time in case of time zone
if ($EndPeakDateTime -lt $BeginPeakDateTime) {
	$EndPeakDateTime = $EndPeakDateTime.AddDays(1)
}

[int]$NumberOfRunningHost = 0
[int]$NumberOfAvailableHost = 0
[int]$NumberOfBlockedHost = 0
[int]$NumberOfHostSessions = 0
[int]$AvailableSessionCapacity = 0 
[int]$TotalOfHostSessions = 0
[int]$TotalOfSessionCapacity = 0
[int]$TotalStarted = 0
[int]$TotalStopped = 0
[int]$TotalBlocked = 0
[int]$TotalUnblocked = 0

# #Collect the credentials from Azure Automation Account Assets
# $Connection = Get-AutomationConnection -Name $ConnectionAssetName

#Authenticating to Azure
Clear-AzContext -Force

# Change autentication method to MSI
Connect-AzAccount -Identity

# $AZAuthentication = Connect-AzAccount -ApplicationId $Connection.ApplicationId -TenantId $Connection.TenantId  -CertificateThumbprint $Connection.CertificateThumbprint -ServicePrincipal

# if ($AZAuthentication -eq $null) {
# 	Write-Output "Failed to authenticate Azure: $($_.exception.message)"
# 	exit
# } else {
# 	$AzObj = $AZAuthentication | Out-String
# 	Write-Output "Authenticating as service principal for Azure. Result: `n$AzObj"
# }

#Set the Azure context with Subscription
#$AzContext = Set-AzContext -SubscriptionId $Connection.SubscriptionID

# if ($AzContext -eq $null) {
# 	Write-Error "Please provide a valid subscription"
# 	exit
# } else {
# 	$AzSubObj = $AzContext | Out-String
# 	Write-Output "Sets the Azure subscription. Result: `n$AzSubObj"
# }

#Checking given host pool name exists in Tenant, and get a MaxSessionLimit
$HostpoolInfo = Get-AzWvdHostPool -ResourceGroupName $ResourceGroupName -Name $HostpoolName
if ($HostpoolInfo -eq $null) {
	Write-Output "Hostpoolname '$HostpoolName' does not exist in resource group '$ResourceGroupName'. Ensure that you have entered the correct values."
	exit
} else {
    [int]$MaxSessionLimitValue = $HostpoolInfo.MaxSessionLimit
}

Write-Output "Starting WVD tenant hosts scale optimization: Current Date Time is: $CurrentDateTime"
Write-Output "Processing HostPoolName: $HostpoolName"

# Get user sessions of HostPool
try {
    $HostPoolUserSessions = Get-AzWvdUserSession -ResourceGroupName $ResourceGroupName -HostPoolName $HostpoolName
} catch {
	Write-Error "Failed to retrieve user sessions in hostpool: $HostpoolName with error: $($_.exception.message)"
	exit
}

if($CurrentDateTime.Hour -gt 12) {
    if($DayOfWeek -eq 'Saturday') {
        $Day = $CurrentDateTime.DayOfWeek.value__
    } else {
        $Day = $CurrentDateTime.DayOfWeek.value__+1
    }
} else {
    $Day = $CurrentDateTime.DayOfWeek.value__
}

$ListOfSessionHosts = Get-AzWvdSessionHost -ResourceGroupName $ResourceGroupName -HostPoolName $HostpoolName |Sort-Object Status

if ($ListOfSessionHosts -eq $null) {
	Write-Output "Session hosts does not exist in the Hostpool of $HostpoolName."
	exit
}

# Processing all Azure VM 
$IndexOfAzVM = @{}
$IndexOfAzVmSize = @{}
$ListOfAzVM = Get-AzVM -ResourceGroupName $ResourceGroupName -Status

foreach ($vm in $ListOfAzVM) {
    $IndexOfAzVM.Add($vm.Name,$vm)
    if (!$IndexOfAzVmSize.ContainsKey($vm.HardwareProfile.VmSize)) {
        $VMSize = Get-AzVMSize -Location $vm.Location | ?{$_.Name -eq $vm.HardwareProfile.VmSize}
        $IndexOfAzVmSize.add($vm.HardwareProfile.VmSize,$VMSize)
    }
}

$SkipSessionhosts = @()
foreach ($SessionHost in $ListOfSessionHosts) {
	$SessionHostName = $SessionHost.Name.Split("/")[1]
	$VMName = $SessionHostName.Split(".")[0]
	$AzureVM = $IndexOfAzVM.$VMName
    $RoleSize = $IndexOfAzVmSize.$($AzureVM.HardwareProfile.VmSize)

	# Check the session host is in maintenance
	if ($AzureVM.Tags.Keys -contains $MaintenanceTagName) {
		Write-Output "Skipping, session host is in maintenance: $VMName"
		$SkipSessionhosts += $SessionHost
		continue
	}

    Write-Output "Checking session host: $SessionHostName  of sessions:$($SessionHost.Session) and status:$($SessionHost.Status)"

	if ($SessionHostName.ToLower().Contains($AzureVM.Name.ToLower())) {
		# Check if the Azure VM is running
		if ($AzureVM.PowerState -match "running") {
            #Forcing logoff if status not equal "Available", block new sessions and add maintenance tag
            if ($SessionHost.Status -ne "Available" -and $SessionHost.Status -ne "NeedsAssistance") {
                Write-Output "Skipping, session host is unavailable: $SessionHostName"
                $SkipSessionhosts += $SessionHost	
                
                #Nao contabiilza o AllowNewSession pois o host pode estar unavailable
                Write-Output "Setting AllowNewSession to FALSE on $SessionHostName"	
                Update-AzWvdSessionHost -HostPoolName $HostPoolName -ResourceGroupName $ResourceGroupName -Name $SessionHostName -AllowNewSession:$false -ErrorAction Continue

                Write-Output "Restarting Azure VM: $VMName"
				Restart-SessionHost -VMName $VMName   
                
                <# 
                LOGICA DE LOGOFF E STOP ALTERADA PARA RESTART DA VM QUANDO ESTIVER UNAVAILABLE

                #Write-Output "Tagging resource with maintenance tag name ($MaintenanceTagName): $VMName"
                #$AzureVM | %{Set-AzResource -ResourceGroupName $_.ResourceGroupName -Name $_.Name -ResourceType $_.Type -Tag @{"$MaintenanceTagName"="$MaintenanceTagName"} -Force -AsJob}

                Write-Output "Forcing logoff on: $SessionHostName"
                #### ----> OLHAR AQUI!!!! PARÂMETROS...
                $SessionHostUserSessions = $HostPoolUserSessions | ?{$_.Name.Contains($SessionHostName) -and $_.SessionState -eq "Active"}
                if ($SessionHostUserSessions.count) {
                    $SessionHostUserSessions | %{
                        #### ----> OLHAR AQUI!!!! PARÂMETROS...
                        Disconnect-AzWvdUserSession -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -SessionHostName $SessionHostName -Id $_.Name.Split("/")[2] -Confirm:$false -ErrorAction SilentlyContinue
                        Remove-AzWvdUserSession -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -SessionHostName $SessionHostName -Id $_.Name.Split("/")[2] -Confirm:$false -ErrorAction SilentlyContinue
                    }
                }
                
                Write-Output "Restarting Azure VM: $VMName"
				Stop-SessionHost -VMName $VMName   
                #>

            } 
            elseif ($SessionHost.AllowNewSession -eq $true) {
                [int]$NumberOfRunningHost+=1
                [int]$NumberOfAvailableHost+=1
                [int]$AvailableSessionCapacity += $($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
                [int]$NumberOfHostSessions += $SessionHost.Session
                [int]$TotalOfHostSessions += $SessionHost.Session
                [int]$TotalOfSessionCapacity += $($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
            } 
            else {
                [int]$NumberOfRunningHost+=1
                [int]$NumberOfBlockedHost+=1
                [int]$TotalOfHostSessions += $SessionHost.Session
                [int]$TotalOfSessionCapacity += $($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
            }
		}
	}
}

Write-Output "Current status of HostpoolName: $HostpoolName"
Write-Output "Current SkipSessionHosts: $($SkipSessionhosts.count)"
Write-Output "Current NumberOfRunningHosts: $NumberOfRunningHost"
Write-Output "Current NumberOfBlockedHosts: $NumberOfBlockedHost"
Write-Output "Current NumberOfAvailableHost: $NumberOfAvailableHost"
Write-Output "Current TotalOfSessionCapacity: $TotalOfSessionCapacity"
Write-Output "Current TotalOfHostSessions: $TotalOfHostSessions"
Write-Output "------------------------------------------------------------"

# Maintenance VMs skipped and stored into a variable
$AllSessionHosts = $ListOfSessionHosts | Where-Object { $SkipSessionhosts -notcontains $_ }

# Sort Host List to START-STOP VM
if($Day%2 -eq 0) {
    $AllSessionHostsStart = $AllSessionHosts | Sort-Object SessionHostName
    $AllSessionHostsStop = $AllSessionHosts | Sort-Object SessionHostName -Descending
} else {
    $AllSessionHostsStart = $AllSessionHosts | Sort-Object SessionHostName -Descending
    $AllSessionHostsStop = $AllSessionHosts | Sort-Object SessionHostName
}

# Grant minimum of session hosts as needed, independently of time
# Check if needed START VM to meet the minimum requirement      
if ($NumberOfAvailableHost -lt $MinimumNumberOfRDSH) {
	Write-Output "Current number of available session hosts ($NumberOfAvailableHost) is less than minimum requirements ($MinimumNumberOfRDSH), start session host ..."
    foreach ($SessionHost in $AllSessionHostsStart) {
        $SessionHostName = $SessionHost.Name.Split("/")[1]
        $VMName = $SessionHostName.Split(".")[0]
        $AzureVM = $IndexOfAzVM.$VMName
        $RoleSize = $IndexOfAzVmSize.$($AzureVM.HardwareProfile.VmSize)

        # Check whether the number of running VMs meets the minimum or not
		if ($NumberOfAvailableHost -lt $MinimumNumberOfRDSH) {
            # Check if the Azure VM is healthy
            if ($SessionHost.UpdateState -eq "Succeeded") {
                # Check if the Azure VM is running
                if ($AzureVM.PowerState -notmatch "running") {
				    # Check if necessary configure session host with AllowNewSession TRUE
                    if ($SessionHost.AllowNewSession -eq $false) {
                        Write-Output "Unblock session host, setting AllowNewSession to TRUE on $SessionHostName"
                        Set-AllowNewSession -SessionHostName $SessionHostName -AllowNewSession $true
                    }
						
                    # Start the Az VM
				    Write-Output "Starting Azure VM: $VMName"
				    Start-SessionHost -VMName $VMName

				    # Calculate available capacity of sessions
                    [int]$NumberOfRunningHost+=1
                    [int]$NumberOfAvailableHost+=1
                    [int]$AvailableSessionCapacity = $AvailableSessionCapacity + ($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
		        } elseif(($SessionHost.Status -eq "Available" -or $SessionHost.Status -eq "NeedsAssistance") -and $SessionHost.AllowNewSession -eq $false) {
                    Write-Output "Unblock session host, setting AllowNewSession to TRUE on $SessionHostName"	
                    Set-AllowNewSession -SessionHostName $SessionHostName -AllowNewSession $true -Count
                    # Calculate available capacity of sessions
                    [int]$NumberOfBlockedHost-=1
                    [int]$NumberOfAvailableHost+=1
                    [int]$AvailableSessionCapacity = $AvailableSessionCapacity + ($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
                    [int]$NumberOfHostSessions += $SessionHost.Session
                } elseif($SessionHost.Status -eq "Available" -or $SessionHost.Status -eq "NeedsAssistance") {
                    Write-Output "Skipped Azure VM, it's already available: $VMName"
                }
            } else {
                Write-Output "Skipped Azure VM, it's not healthy: $VMName"
            }
        } else {
            Write-Output "Current number of running session hosts is greater than or equal minimum requirements"
            break;
        }
	}
}

# Check if it is off-peak time or weekend to STOP machines
if ($IsPeakDataTime -eq $false -or $DayOfWeek -eq 'Saturday' -or $DayOfWeek -eq 'Sunday') {
	Write-Output "Starting to scale down WVD session hosts ..."
	# Breadth first session hosts shutdown in off peak hours, if the capacity is OK
	if ($NumberOfRunningHost -gt $MinimumNumberOfRDSH -and $AvailableSessionCapacity -gt $NumberOfHostSessions) {
		foreach ($SessionHost in $AllSessionHostsStop) {
            $SessionHostName = $SessionHost.Name.Split("/")[1]
			$VMName = $SessionHostName.Split(".")[0]
            $AzureVM = $IndexOfAzVM.$VMName
            $RoleSize = $IndexOfAzVmSize.$($IndexOfAzVM.$VMName.HardwareProfile.VmSize)

            if ($AzureVM.PowerState -match "running") {
                $DiffSessions = $AvailableSessionCapacity - $NumberOfHostSessions
			    if (($NumberOfAvailableHost -gt $MinimumNumberOfRDSH -and $DiffSessions -gt $MaxSessionLimitValue) -or $SessionHost.AllowNewSession -eq $false) {
                    # Get the user sessions in the session host
                    #### ----> OLHAR AQUI!!!! PARÂMETROS...
                    $SessionHostUserSessions = $HostPoolUserSessions | ?{$_.Name.Contains($SessionHostName)}
                    #Para fins de DEBUG [Isaac]
                    #Write-Output "[DEBUG] HostPoolUserSessions: $HostPoolUserSessions"
                    #Write-Output "[DEBUG] SessionHostUserSessions: $SessionHostUserSessions"
                    
				    $ActiveSessionCount = ($SessionHostUserSessions | ?{$_.SessionState -eq "Active"}).count
                    Write-Output "Counting the current active sessions on the host $SessionHostName :$ActiveSessionCount"

                    # Ensure the running Azure VM is set as drain mode
                    if ($SessionHost.AllowNewSession -eq $false) {
                        $blocked = $false	
                        Write-Output "Skipped, AllowNewSession already is FALSE on $SessionHostName"
                    } else {
                        $blocked = $true

                        Write-Output "Setting AllowNewSession to FALSE on $SessionHostName"	
                        Set-AllowNewSession -SessionHostName $SessionHostName -AllowNewSession $false
                            
                        [int]$NumberOfHostSessions -= $SessionHost.Session
                        [int]$AvailableSessionCapacity = $AvailableSessionCapacity - ($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
                        [int]$NumberOfAvailableHost-=1
                    }
                    #Para fins de DEBUG [Isaac]
                    #Write-Output "[DEBUG] $SessionHostName :$ActiveSessionCount"
                    # Shutdown the Azure VM, which session host have 0 active sessions
				    if ($ActiveSessionCount -eq 0) {
                        Write-Output "Forcing logoff of inactive session on: $VMName"
                        if ($SessionHostUserSessions.count) {
                            $SessionHostUserSessions | %{
                                Disconnect-AzWvdUserSession -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -SessionHostName $SessionHostName -Id $_.Name.Split("/")[2] -Confirm:$false -ErrorAction SilentlyContinue
                                Remove-AzWvdUserSession -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -SessionHostName $SessionHostName -Id $_.Name.Split("/")[2] -Confirm:$false -ErrorAction SilentlyContinue
                            }
                        }

                        Write-Output "Stopping Azure VM: $VMName"
					    Stop-SessionHost -VMName $VMName   

                        [int]$NumberOfRunningHost-=1
                        
                        if (!$blocked) { [int]$NumberOfBlockedHost-=1 }
				    } else {
					    # Notify user to log off session is LimitSecondsToForceLogOffUser is greater than zero
                        if ($LimitSecondsToForceLogOffUser -gt 0) {
                            Write-Output "Starting a runbook for logoff and shutdown on $SessionHostName"	

                            $params = @{
                                "HostpoolName" = $Hostpoolname;
                                "SessionHostName" = $SessionHostName;
                                "LimitSecondsToForceLogOffUser" = $LimitSecondsToForceLogOffUser;
                                "LogOffMessageTitle" = $LogOffMessageTitle;
                                "LogOffMessageBody" = $LogOffMessageBody;
                                "ConnectionAssetName" = $ConnectionAssetName;
                                "ResourceGroupName" = $ResourceGroupName;
                                "VMName" = $VMName
                            }
                            #Write-Output "[DEBUG]: $RunbookLogoffShutdown"
                            #Write-Output "[DEBUG]: $AutomationAccountName"
                            #Write-Output "[DEBUG]: $ResourceGroupNameAutomation"
                            #Write-Output "[DEBUG]: $AzContext"
                            #Write-Output "[DEBUG]: $params"
                            Start-AzAutomationRunbook –Name $RunbookLogoffShutdown –AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupNameAutomation -AzContext $AzContext –Parameters $params

                            [int]$TotalStopped+=1
        				    [int]$NumberOfRunningHost-=1
                            
                            if(!$blocked) { [int]$NumberOfBlockedHost-=1 }

                            sleep -Seconds 5
                        } elseif ($blocked) {
                            [int]$TotalBlocked+=1
                            [int]$NumberOfBlockedHost+=1
                        }
				    }
			    } else {
                    Write-Output "Maintaining Azure VM: $VMName"
                
                    # Ensure the running Azure VM is not set as drain mode
                    if ($SessionHost.AllowNewSession -eq $false -and $SessionHost.Session -le $MaxSessionLimitValue) {
                        Write-Output "Resetting AllowNewSession to TRUE on $SessionHostName"						
                        Set-AllowNewSession -SessionHostName $SessionHostName -AllowNewSession $true -Count
                        # Calculate available capacity of sessions
                        [int]$NumberOfBlockedHost-=1
                        [int]$NumberOfAvailableHost+=1
                        [int]$AvailableSessionCapacity = $AvailableSessionCapacity + ($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
                        [int]$NumberOfHostSessions += $SessionHost.Session
                   }
                }
            } else {
                Write-Output "Skipped Azure VM, it's NOT running: $VMName"
            }
		}
	} else {
        Write-Output "Not needed start scale down WVD session hosts ..."
    }
}

# Check if it is off-peak time or weekend to START machines
if ($IsPeakDataTime -eq $false -or $DayOfWeek -eq 'Saturday' -or $DayOfWeek -eq 'Sunday') {
	Write-Output "Starting session hosts as needed based on current workloads."
	
    # Check if needed start VM to meet the minimum of session capacity    
	if ($NumberOfHostSessions -ge $AvailableSessionCapacity) {
		Write-Output "Current available session capacity is less than demanded user sessions, starting session host"
		# Running out of capacity, we need to start more VMs if there are any 
		foreach ($SessionHost in $AllSessionHostsStart) {
            $SessionHostName = $SessionHost.Name.Split("/")[1]
            $VMName = $SessionHostName.Split(".")[0]
            $AzureVM = $IndexOfAzVM.$VMName
            $RoleSize = $IndexOfAzVmSize.$($AzureVM.HardwareProfile.VmSize)

            if ($NumberOfHostSessions -ge $AvailableSessionCapacity) {
                # Check if the Azure VM is healthy
                if ($SessionHost.UpdateState -eq "Succeeded") {
                    # Check if the Azure VM is running
		            if ($AzureVM.PowerState -notmatch "running") {
						# Check if necessary configure session host with AllowNewSession TRUE
                        if ($SessionHost.AllowNewSession -eq $false) {
                            Write-Output "Unblock session host, setting AllowNewSession to TRUE on $SessionHostName"	
                            Set-AllowNewSession -SessionHostName $SessionHostName -AllowNewSession $true
                        }
							
                        # Start the Az VM
						Write-Output "Starting Azure VM: $VMName"
						Start-SessionHost -VMName $VMName

						# Calculate available capacity of sessions
                        [int]$NumberOfRunningHost+=1
                        [int]$NumberOfAvailableHost+=1
                        [int]$AvailableSessionCapacity = $AvailableSessionCapacity + ($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
					} elseif (($SessionHost.Status -eq "Available" -or $SessionHost.Status -eq "NeedsAssistance") -and $SessionHost.AllowNewSession -eq $false) {
                        Write-Output "Unblock session host, setting AllowNewSession to TRUE on $SessionHostName"	
                        Set-AllowNewSession -SessionHostName $SessionHostName -AllowNewSession $true -Count

                        # Calculate available capacity of sessions
                        [int]$NumberOfBlockedHost-=1
                        [int]$NumberOfAvailableHost+=1
                        [int]$AvailableSessionCapacity = $AvailableSessionCapacity + ($RoleSize.NumberOfCores * $SessionThresholdPerCPU)
                        [int]$NumberOfHostSessions += $SessionHost.Session
                    } elseif ($SessionHost.Status -eq "Available" -or $SessionHost.Status -eq "NeedsAssistance") {
                        Write-Output "Skipped Azure VM, it's already available: $VMName"
                    }
                } else {
                    Write-Output "Skipped Azure VM, it's not healthy: $VMName"
                }
            } else {
                Write-Output "Current available session capacity is greater than or equal demanded user sessions"
                break
            }
		}
	} else {
        Write-Output "No session hosts needed based on current workloads."
    }
}

Write-Output "Total VM STARTED in this execution: $TotalStarted"
Write-Output "Total VM STOPPED in this execution: $TotalStopped"
Write-Output "Total VM BLOCKED in this execution: $TotalBlocked"
Write-Output "Total VM UNBLOCKED in this execution: $TotalUnblocked"
Write-Output "------------------------------------------------------------"
Write-Output "Final status of Hostpool: $HostpoolName"
Write-Output "Final number of running hosts: $NumberOfRunningHost"
Write-Output "Final number of blocked hosts: $NumberOfBlockedHost"
Write-Output "Final number of available hosts: $NumberOfAvailableHost"
Write-Output "Final number of availabe capacity: $AvailableSessionCapacity"
Write-Output "Final sessions on available Hosts: $NumberOfHostSessions"
Write-Output "End WVD tenant scale optimization."