#
# AzureStack HCI Registration and Unregistration Powershell Cmdlets.
#

$ErrorActionPreference = 'Stop'

$GAOSBuildNumber = 17784
$GAOSUBR = 1374
$V2OSBuildNumber = 20348
$V2OSUBR = 288
#region User visible strings

$AdminConsentWarning = "You need additional Azure Active Directory permissions to register in this Azure subscription. Contact your Azure AD administrator to grant consent to AAD application identity {0} at {1}. Then, run Register-AzStackHCI again with same parameters to complete registration."
$NoClusterError = "Computer {0} is not part of an Azure Stack HCI cluster. Use the -ComputerName parameter to specify an Azure Stack HCI cluster node and try again."
$CloudResourceDoesNotExist = "The Azure resource with ID {0} doesn't exist. Unregister the cluster using Unregister-AzStackHCI and then try again."
$RegisteredWithDifferentResourceId = "Azure Stack HCI is already registered with Azure resource ID {0}. To register or change registration, first unregister the cluster using Unregister-AzStackHCI, then try again."
$RegistrationInfoNotFound = "Additional parameters are required to unregister. Run 'Get-Help Unregister-AzStackHCI -Full' for more information."
$RegionNotSupported = "Azure Stack HCI is not yet available in region {0}. Please choose one of these regions: {1}."
$CertificateNotFoundOnNode = "Certificate with thumbprint {0} not found on node(s) {1}. Make sure the certificate has been added to the certificate store on every clustered node."
$SettingCertificateFailed = "Failed to register. Couldn't generate self-signed certificate on node(s) {0}. Couldn't set and verify registration certificate on node(s) {1}. Make sure every clustered node is up and has Internet connectivity (at least outbound to Azure)."
$InstallLatestVersionWarning = "Newer version of the Az.StackHCI module is available. Update from version {0} to version {1} using Update-Module."
$NotAllTheNodesInClusterAreGA = "Update the operating system on node(s) {0} to version $GAOSBuildNumber.$GAOSUBR or later to continue."
$NoExistingRegistrationExistsErrorMessage = "Can't repair registration because the cluster isn't registered yet. Register the cluster using Register-AzStackHCI without the -RepairRegistration option."
$UserCertValidationErrorMessage = "Can't use certificate with thumbprint {0} because it expires in less than 60 days, on {1}. Certificates must be valid for at least 60 days."
$FailedToRemoveRegistrationCertWarning = "Couldn't clean up Azure Stack HCI registration certificate from node(s) {0}. You can ignore this message or clean up the certificate yourself (optional)."
$UnregistrationSuccessDetailsMessage = "Azure Stack HCI is successfully unregistered. The Azure resource representing Azure Stack HCI has been deleted. Azure Stack HCI can't sync with Azure until you register again."
$RegistrationSuccessDetailsMessage = "Azure Stack HCI is successfully registered. An Azure resource representing Azure Stack HCI has been created in your Azure subscription to enable an Azure-consistent monitoring, billing, and support experience."
$CouldNotGetLatestModuleInformationWarning = "Can't connect to the PowerShell Gallery to verify module version. Make sure you have the latest Az.StackHCI module with major version {0}.*."
$ConnectingToCloudBillingServiceFailed = "Can't reach Azure from node(s) {0}. Make sure every clustered node has network connectivity to Azure. Verify that your network firewall allows outbound HTTPS from port 443 to all the well-known Azure IP addresses and URLs required by Azure Stack HCI. Visit aka.ms/hcidocs for details."
$ResourceExistsInDifferentRegionError = "There is already an Azure Stack HCI resource with the same resource ID in region {0}, which is different from the input region {1}. Either specify the same region or delete the existing resource and try again."
$ArcCmdletsNotAvailableError = "Azure Arc integration isn't available for the version of Azure Stack HCI installed on node(s) {0} yet. Check the documentation for details. You may need to install an update or join the Preview channel."
$ArcRegistrationDisableInProgressError = "Unregister of Azure Arc integration is in progress. Try Unregister-AzStackHCI to finish unregistration and then try Register-AzStackHCI again."
$ArcIntegrationNotAvailableForCloudError = "Azure Arc integration is not available in {0}. Specify '-EnableAzureArcServer:`$false' in Register-AzStackHCI Cmdlet to register without Arc integration."

$FetchingRegistrationState = "Checking whether the cluster is already registered"
$ValidatingParametersFetchClusterName = "Validating cmdlet parameters"
$ValidatingParametersRegisteredInfo = "Validating the parameters and checking registration information"
$RegisterProgressActivityName = "Registering Azure Stack HCI with Azure..."
$UnregisterProgressActivityName = "Unregistering Azure Stack HCI from Azure..."
$InstallAzResourcesMessage = "Installing required PowerShell module: Az.Resources"
$InstallAzureADMessage = "Installing required PowerShell module: AzureAD"
$InstallRSATClusteringMessage = "Installing required Windows feature: RSAT-Clustering-PowerShell"
$LoggingInToAzureMessage = "Logging in to Azure"
$ConnectingToAzureAD = "Connecting to Azure Active Directory"
$RegisterAzureStackRPMessage = "Registering Microsoft.AzureStackHCI provider to Subscription"
$CreatingAADAppMessage = "Creating AAD application {0} in Azure AD directory {1}"
$CreatingResourceGroupMessage = "Creating Azure Resource Group {0}"
$CreatingCloudResourceMessage = "Creating Azure Resource {0} representing Azure Stack HCI by calling Microsoft.AzureStackHCI provider"
$GrantingAdminConsentMessage = "Trying to grant admin consent for the required permissions needed for Azure AD application identity {0}"
$GettingCertificateMessage = "Getting new certificate from on-premises cluster to use as application credential"
$AddAppCredentialMessage = "Adding certificate as application credential for the Azure AD application {0}"
$RegisterAndSyncMetadataMessage = "Registering Azure Stack HCI cluster and syncing cluster census information from the on-premises cluster to the cloud"
$UnregisterHCIUsageMessage = "Unregistering Azure Stack HCI cluster and cleaning up registration state on the on-premises cluster"
$DeletingAADApplicationMessage = "Deleting Azure AD application identity {0}"
$DeletingCloudResourceMessage = "Deleting Azure resource with ID {0} representing the Azure Stack HCI cluster"
$DeletingArcCloudResourceMessage = "Deleting Azure resource with ID {0} representing the Azure Stack HCI cluster Arc integration"
$DeletingExtensionMessage = "Deleting extension {0} on cluster {1}"
$DeletingCertificateFromAADApp = "Deleting certificate with KeyId {0} from Azure Active Directory"
$SkippingDeleteCertificateFromAADApp = "Certificate with KeyId {0} is still being used by Azure Active Directory and won't be deleted"
$RegisterArcMessage = "Arc for servers registration triggered"
$UnregisterArcMessage = "Arc for servers unregistration triggered"

$RegisterArcProgressActivityName = "Registering Azure Stack HCI with Azure Arc..."
$UnregisterArcProgressActivityName = "Unregistering Azure Stack HCI with Azure Arc..."
$RegisterArcRPMessage = "Registering Microsoft.HybridCompute and Microsoft.GuestConfiguration resource providers to subscription"
$SetupArcMessage = "Initializing Azure Stack HCI integration with Azure Arc"
$StartingArcAgentMessage = "Enabling Azure Arc integration on every clustered node"
$WaitingUnregisterMessage = "Disabling Azure Arc integration on every clustered node"
$CleanArcMessage = "Cleaning up Azure Arc integration"

$ArcAgentRolesInsufficientPreviligeMessage = "Failed to assign required roles for Azure Arc integration. Your Azure AD account must be an Owner or User Access Administrator in the subscription to enable Azure Arc integration."
$ArcRolesCleaningWarningMessage = "Couldn't clean up Azure AD application identity {0} used by Azure Arc integration. You can ignore this message or clean it up yourself through the Azure portal (optional)."
$RegisterArcFailedWarningMessage = "Some clustered nodes couldn't be Arc-enabled right now. This can happen if some of the nodes are down. We'll automatically try again in an hour. In the meantime, you can use Get-AzureStackHCIArcIntegration to check status on each node."
$UnregisterArcFailedError = "Couldn't disable Azure Arc integration on Node {0}. Try running Disable-AzureStackHCIArcIntegration Cmdlet on the node. If the node is in a state where Disable-AzureStackHCIArcIntegration Cmdlet could not be run, remove the node from the cluster and try Unregister-AzStackHCI Cmdlet again."
$ArcExtensionCleanupFailedError = "Couldn't delete Arc extension {0} on cluster nodes. You can try the extension uninstallation steps listed at https://docs.microsoft.com/en-us/azure/azure-arc/servers/manage-agent for removing the extension and try Unregister-AzStackHCI again. If the node is in a state where extension uninstallation could not succeed, try Unregister-AzStackHCI with -Force switch."
$ArcExtensionCleanupFailedWarning = "Couldn't delete Arc extension {0} on cluster nodes. Extension may continue to run even after unregistration."

$SetProgressActivityName = "Setting properties for the Azure Stack HCI resource in Azure..."
$SetProgressStatusGathering = "Gathering information"
$SetProgressStatusGetAzureResource = "Getting the Azure Stack HCI resource"
$SetProgressStatusOpSwitching = "Switching to the subscription ID {0}"
$SetProgressStatusUpdatingProps = "Updating the resource properties"
$SetProgressStatusSyncCluster = "Syncing the Azure Stack HCI cluster with Azure"
$SetAzResourceClusterNotRegistered = "The cluster is not registered with Azure. Register the cluster using Register-AzStackHCI and then try again."
$SetAzResourceClusterNodesDown = "One or more servers in your cluster are offline. Check that all your servers are up and then try again."
$SetAzResourceSuccessWSSE = "Successfully enabled Windows Server Subscription."
$SetAzResourceSuccessWSSD = "Successfully disabled Windows Server Subscription."
$SetAzResourceSuccessDiagLevel = "Successfully configured the Azure Stack HCI diagnostic level to {0}."
$SetProgressShouldProcess = "Update the resource properties to change Windows Server Subscription or Azure Stack HCI diagnostic level"
$SetProgressShouldContinue = "This will enable or disable billing for Windows Server guest licenses through your Azure subscription."
$SetProgressShouldContinueCaption = "Configure Windows Server Subscription"
$SetProgressWarningDiagnosticOff = "Setting diagnostic level to Off will prevent Microsoft from collecting important diagnostic information that helps improve Azure Stack HCI."
$SetProgressWarningWSSD = "Windows Server Subscription will no longer activate your Windows Server VMs. Please check that your VMs are being activated another way."
$SecondaryProgressBarId = 2
$EnableAzsHciImdsActivity = "Enable Azure Stack HCI IMDS Attestation..."
$ConfirmEnableImds = "Enabling IMDS Attestation configures your cluster to use workloads that are exclusively available on Azure."
$ConfirmDisableImds = "Disabling IMDS Attestation will remove the ability for some exclusive Azure workloads to function."
$ImdsClusterNotRegistered = "The cluster is not registered with Azure. Register the cluster using Register-AzStackHCI and then try again."
$DisableAzsHciImdsActivity = "Disable Azure Stack HCI IMDS Attestation..."
$AddAzsHciImdsActivity = "Add Virtual Machines to Azure Stack HCI IMDS Attestation..."
$RemoveAzsHciImdsActivity = "Remove Virtual Machines from Azure Stack HCI IMDS Attestation..."
$ShouldContinueHyperVInstall = "The Hyper-V Powershell management tools are required to be installed on {0} to continue. Install RSAT-Hyper-V-Tools and continue?"
$DiscoveringClusterNodes = "Discovering cluster nodes..."
$AllClusterNodesAreNotOnline = "One or more servers in your cluster are offline. Check that all your servers are up and then try again."
$CheckingClusterNode = "Checking AzureStack HCI IMDS Attestation on {0}"
$ConfiguringClusterNode = "Configuring AzureStack HCI IMDS Attestation on {0}"
$DisablingIMDSOnNode = "Disabling AzureStack HCI IMDS Attestation on {0}"
$RemovingVmImdsFromNode = "Removing AzureStack HCI IMDS Attestation from guests on {0}"
$AttestationNotEnabled = "The IMDS Service on {0} needs to be activated. This is required before guests can be configured. Run Enable-AzStackHCIAttestation cmdlet."
$ErrorAddingAllVMs = "Did not add all guests. Try running Add-AzStackHCIVMAttestation on each node manually."
#endregion

#region Constants

$UsageServiceFirstPartyAppId = "1322e676-dee7-41ee-a874-ac923822781c"
$MicrosoftTenantId = "72f988bf-86f1-41af-91ab-2d7cd011db47"

$MSPortalDomain = "https://ms.portal.azure.com/"
$AzureCloudPortalDomain = "https://portal.azure.com/"
$AzureChinaCloudPortalDomain = "https://portal.azure.cn/"
$AzureUSGovernmentPortalDomain = "https://portal.azure.us/"
$AzureGermanCloudPortalDomain = "https://portal.microsoftazure.de/"
$AzurePPEPortalDomain = "https://df.onecloud.azure-test.net/"
$AzureCanaryPortalDomain = "https://portal.azure.com/"

$AzureCloud = "AzureCloud"
$AzureChinaCloud = "AzureChinaCloud"
$AzureUSGovernment = "AzureUSGovernment"
$AzureGermanCloud = "AzureGermanCloud"
$AzurePPE = "AzurePPE"
$AzureCanary = "AzureCanary"

$PortalCanarySuffix = '?feature.armendpointprefix={0}'
$PortalAADAppPermissionUrl = '#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/CallAnAPI/appId/{0}/isMSAApp/'
$PortalHCIResourceUrl = '#@{0}/resource/subscriptions/{1}/resourceGroups/{2}/providers/Microsoft.AzureStackHCI/clusters/{3}/overview'

# GUID's of the scopes generated in first party portal
$ClusterReadPermission = "2344a320-6a09-4530-bed7-c90485b5e5e2"
$ClusterReadWritePermission = "493bd689-9082-40db-a506-11f40b68128f"
$ClusterNodeReadPermission = "8fa5445e-80fb-4c71-a3b1-9a16a81a1966"
$ClusterNodeReadWritePermission = "bbe8afc9-f3ba-4955-bb5f-1cfb6960b242"

# Deprecated scopes. Will be deleted during RepairRegistration
$BillingSyncPermission = "e4359fc6-82ee-4411-9a4d-edfc7812cf24"
$CensusSyncPermission = "8c83ab0a-0f96-40e9-940b-20cc5c5ecca9"

$PermissionIds = New-Object System.Collections.Generic.List[string]

$PermissionIds.Add($ClusterReadPermission)
$PermissionIds.Add($ClusterReadWritePermission)
$PermissionIds.Add($ClusterNodeReadPermission)
$PermissionIds.Add($ClusterNodeReadWritePermission)

$Region_EASTUSEUAP = 'eastus2euap'

[hashtable] $ServiceEndpointsAzureCloud = @{
        $Region_EASTUSEUAP = 'https://eus2euap-azurestackhci-usage.azurewebsites.net';
        }

$ServiceEndpointAzureCloudFrontDoor = "https://azurestackhci.azurefd.net"
$ServiceEndpointAzureCloud = $ServiceEndpointAzureCloudFrontDoor

$AuthorityAzureCloud = "https://login.microsoftonline.com"
$BillingServiceApiScopeAzureCloud = "https://azurestackhci-usage.trafficmanager.net/.default"
$GraphServiceApiScopeAzureCloud = "https://graph.microsoft.com/.default"

$ServiceEndpointAzurePPE = "https://azurestackhci-df.azurefd.net"
$AuthorityAzurePPE = "https://login.windows-ppe.net"
$BillingServiceApiScopeAzurePPE = "https://azurestackhci-usage-df.azurewebsites.net/.default"
$GraphServiceApiScopeAzurePPE = "https://graph.ppe.windows.net/.default"

$ServiceEndpointAzureChinaCloud = "https://dp.stackhci.azure.cn"
$AuthorityAzureChinaCloud = "https://login.partner.microsoftonline.cn"
$BillingServiceApiScopeAzureChinaCloud = "$UsageServiceFirstPartyAppId/.default"
$GraphServiceApiScopeAzureChinaCloud = "https://microsoftgraph.chinacloudapi.cn/.default"

$ServiceEndpointAzureUSGovernment = "https://dp.azurestackhci.azure.us"
$AuthorityAzureUSGovernment = "https://login.microsoftonline.us"
$BillingServiceApiScopeAzureUSGovernment = "https://dp.azurestackhci.azure.us/.default"
$GraphServiceApiScopeAzureUSGovernment = "https://graph.windows.net/.default"

$ServiceEndpointAzureGermanCloud = "https://azurestackhci-usage.trafficmanager.de"
$AuthorityAzureGermanCloud = "https://login.microsoftonline.de"
$BillingServiceApiScopeAzureGermanCloud = "https://azurestackhci-usage.azurewebsites.de/.default"
$GraphServiceApiScopeAzureGermanCloud = "https://graph.cloudapi.de/.default"

$RPAPIVersion = "2021-09-01";
$HCIArcAPIVersion = "2021-09-01"
$HCIArcInstanceName = "/arcSettings/default"
$HCIArcExtensions = "/Extensions"

$OutputPropertyResult = "Result"
$OutputPropertyResourceId = "AzureResourceId"
$OutputPropertyPortalResourceURL = "AzurePortalResourceURL"
$OutputPropertyPortalAADAppPermissionsURL = "AzurePortalAADAppPermissionsURL"
$OutputPropertyDetails = "Details"
$OutputPropertyTest = "Test"
$OutputPropertyEndpointTested = "EndpointTested"
$OutputPropertyIsRequired = "IsRequired"
$OutputPropertyFailedNodes = "FailedNodes"
$OutputPropertyErrorDetail = "ErrorDetail"

$ConnectionTestToAzureHCIServiceName = "Connect to Azure Stack HCI Service"

$ResourceGroupCreatedByName = "CreatedBy"
$ResourceGroupCreatedByValue = "4C02703C-F5D0-44B0-ADC3-4ED5C2839E61"

$HealthEndpointPath = "/health"


$MainProgressBarId = 1
$ArcProgressBarId = 2
$IndefinitelyYears = 300

$AzureConnectedMachineOnboardingRole = "Azure Connected Machine Onboarding"
$AzureConnectedMachineResourceAdministratorRole = "Azure Connected Machine Resource Administrator"
$ArcRegistrationTaskName = "ArcRegistrationTask"
$LogFileDir = '\Tasks\ArcForServers'

$ClusterScheduledTaskWaitTimeMinutes = 15
$ClusterScheduledTaskSleepTimeSeconds = 3
$ClusterScheduledTaskRunningState = "Running"
$ClusterScheduledTaskReadyState = "Ready"

$ArcSettingsDisableInProgressState = "DisableInProgress"

enum DiagnosticLevel
{
    Off;
    Basic;
    Enhanced
}
enum ArcStatus
{
    Unknown;
    Enabled;
    Disabled;
    DisableInProgress;
}

enum RegistrationStatus
{
    Registered;
    NotYet;
    OutOfPolicy;
}

enum CertificateManagedBy
{
    Invalid;
    User;
    Cluster;
}
enum VMAttestationStatus
{
    Unknown;
    Connected;
    Disconnected;
}
enum ImdsAttestationNodeStatus
{
    Inactive;
    Active;
    Expired;
    Error;
}

$registerArcScript = {
    try
    {
        # Params for Enable-AzureStackHCIArcIntegration 
        $AgentInstaller_WebLink                  = 'https://aka.ms/AzureConnectedMachineAgent'
        $AgentInstaller_Name                     = 'AzureConnectedMachineAgent.msi'
        $AgentInstaller_LogFile                  = 'ConnectedMachineAgentInstallationLog.txt'
        $AgentExecutable_Path                    =  $Env:Programfiles + '\AzureConnectedMachineAgent\azcmagent.exe'

        $DebugPreference = 'Continue'

        # Setup Directory.
        $LogFileDir = $env:windir + '\Tasks\ArcForServers'
        if (-Not $(Test-Path $LogFileDir))
        {
            New-Item -Type Directory -Path $LogFileDir
        }

        # Delete log files older than 15 days
        Get-ChildItem -Path $LogFileDir -Recurse | Where-Object {($_.LastWriteTime -lt (Get-Date).AddDays(-15))} | Remove-Item

        # Setup Log file name.
        $date = Get-Date
        $datestring = '{0}{1:d2}{2:d2}' -f $date.year,$date.month,$date.day
        $LogFileName = $LogFileDir + '\RegisterArc_' + $datestring + '.log'
    
        Start-Transcript -LiteralPath $LogFileName -Append | Out-Null

        Write-Information 'Triggering Arc For Servers registration cmdlet'
        $arcStatus = Get-AzureStackHCIArcIntegration

        if ($arcStatus.ClusterArcStatus -eq 'Enabled')
        {
            $nodeStatus = $arcStatus.NodesArcStatus
    
            if ($nodeStatus.Keys -icontains ($env:computername))
            {
                if ($nodeStatus[$env:computername.ToLowerInvariant()] -ne 'Enabled')
                {
                    Write-Information 'Registering Arc for servers.'
                    Enable-AzureStackHCIArcIntegration -AgentInstallerWebLink $AgentInstaller_WebLink -AgentInstallerName $AgentInstaller_Name -AgentInstallerLogFile $AgentInstaller_LogFile -AgentExecutablePath $AgentExecutable_Path
                    Sync-AzureStackHCI
                }
                else
                {
                    Write-Information 'Node is already registered.'
                }
            }
            else
            {
                # New node added case.
                Write-Information 'Registering Arc for servers.'
                Enable-AzureStackHCIArcIntegration -AgentInstallerWebLink $AgentInstaller_WebLink -AgentInstallerName $AgentInstaller_Name -AgentInstallerLogFile $AgentInstaller_LogFile -AgentExecutablePath $AgentExecutable_Path
                Sync-AzureStackHCI
            }
        }
        else
        {
            Write-Information ('Cluster Arc status is not enabled. ClusterArcStatus:' + $arcStatus.ClusterArcStatus.ToString())
        }
    }
    catch
    {
        Write-Error -Exception $_.Exception -Category OperationStopped
        # Get script line number, offset and Command that resulted in exception. Write-Error with the exception above does not write this info.
        $positionMessage = $_.InvocationInfo.PositionMessage
        Write-Error ('Exception occurred in RegisterArcScript : ' + $positionMessage) -Category OperationStopped
    }
    finally
    {
        try{ Stop-Transcript } catch {}
    }
}

#endregion

function Setup-Logging{
param(
    [string] $LogFilePrefix
    )
    
    $date = Get-Date
    $datestring = "{0}{1:d2}{2:d2}-{3:d2}{4:d2}" -f $date.year,$date.month,$date.day,$date.hour,$date.minute
    $LogFileName = $LogFilePrefix + "_" + $datestring + ".log"

    Start-Transcript -LiteralPath $LogFileName -Append | out-null
}

function Show-LatestModuleVersion{
    
    $latestModule = Find-Module -Name Az.StackHCI -ErrorAction Ignore
    $installedModule = Get-Module -Name Az.StackHCI | Sort-Object  -Property Version -Descending | Select-Object -First 1

    if($Null -eq $latestModule)
    {
        $CouldNotGetLatestModuleInformationWarningMsg = $CouldNotGetLatestModuleInformationWarning -f $installedModule.Version.Major
        Write-Warning $CouldNotGetLatestModuleInformationWarningMsg
    }
    else
    {
        if($latestModule.Version.GetType() -eq [string])
        {
            $latestModuleVersion = [System.Version]::Parse($latestModule.Version)
        }
        else
        {
            $latestModuleVersion = $latestModule.Version
        }

        if(($latestModuleVersion.Major -eq $installedModule.Version.Major) -and ($latestModuleVersion -gt $installedModule.Version))
        {
            $InstallLatestVersionWarningMsg = $InstallLatestVersionWarning -f $installedModule.Version, $latestModuleVersion
            Write-Warning $InstallLatestVersionWarningMsg
        }
    }
}

function Retry-Command {
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [scriptblock] $ScriptBlock,
        [int]  $Attempts                   = 8,
        [int]  $MinWaitTimeInSeconds       = 5,
        [int]  $MaxWaitTimeInSeconds       = 60,
        [int]  $BaseBackoffTimeInSeconds   = 2,
        [bool] $RetryIfNullOutput          = $true
        )

    $attempt = 0
    $completed = $false
    $result = $null

    if($MaxWaitTimeInSeconds -lt $MinWaitTimeInSeconds)
    {
        throw "MaxWaitTimeInSeconds($MaxWaitTimeInSeconds) is less than MinWaitTimeInSeconds($MinWaitTimeInSeconds)"
    }

    while (-not $completed) {
        try
        {
            $attempt = $attempt + 1
            $result = Invoke-Command -ScriptBlock $ScriptBlock

            if($RetryIfNullOutput)
            {
                if($result -ne $null)
                {
                    Write-Verbose ("Command [{0}] succeeded. Non null result received." -f $ScriptBlock)
                    $completed = $true
                }
                else
                {
                    throw "Null result received."
                }
            }
            else
            {
                Write-Verbose ("Command [{0}] succeeded." -f $ScriptBlock)
                $completed = $true
            }
        }
        catch
        {
            $exception = $_.Exception

            if([int]$exception.ErrorCode -eq [int][system.net.httpstatuscode]::Forbidden)
            {
                Write-Verbose ("Command [{0}] failed Authorization. Attempt {1}. Exception: {2}" -f $ScriptBlock, $attempt,$exception.Message)
                throw
            }
            else
            {
                if ($attempt -ge $Attempts)
                {
                    Write-Verbose ("Command [{0}] failed the maximum number of {1} attempts. Exception: {2}" -f $ScriptBlock, $attempt,$exception.Message)
                    throw
                }
                else
                {
                    $secondsDelay = $MinWaitTimeInSeconds + [int]([Math]::Pow($BaseBackoffTimeInSeconds,($attempt-1)))

                    if($secondsDelay -gt $MaxWaitTimeInSeconds)
                    {
                        $secondsDelay = $MaxWaitTimeInSeconds
                    }

                    Write-Verbose ("Command [{0}] failed. Retrying in {1} seconds. Exception: {2}" -f $ScriptBlock, $secondsDelay,$exception.Message)
                    Start-Sleep $secondsDelay
                }
            }
        }
    }

    return $result
}

function Get-PortalDomain{
param(
    [string] $TenantId,
    [string] $EnvironmentName,
    [string] $Region
    )

    if($EnvironmentName -eq $AzureCloud -and $TenantId -eq $MicrosoftTenantId)
    {
        return $MSPortalDomain;
    }
    elseif($EnvironmentName -eq $AzureCloud)
    {
        return $AzureCloudPortalDomain;
    }
    elseif($EnvironmentName -eq $AzureChinaCloud)
    {
        return $AzureChinaCloudPortalDomain;
    }
    elseif($EnvironmentName -eq $AzureUSGovernment)
    {
        return $AzureUSGovernmentPortalDomain;
    }
    elseif($EnvironmentName -eq $AzureGermanCloud)
    {
        return $AzureGermanCloudPortalDomain;
    }
    elseif($EnvironmentName -eq $AzurePPE)
    {
        return $AzurePPEPortalDomain;
    }
    elseif($EnvironmentName -eq $AzureCanary)
    {
        $PortalCanarySuffixWithRegion = $PortalCanarySuffix -f $Region
        return ($AzureCanaryPortalDomain + $PortalCanarySuffixWithRegion);
    }
}

function Get-DefaultRegion{
param(
    [string] $EnvironmentName
    )

    $defaultRegion = "eastus";

    if($EnvironmentName -eq $AzureCloud)
    {
        $defaultRegion = "eastus"
    }
    elseif($EnvironmentName -eq $AzureChinaCloud)
    {
        $defaultRegion = "chinaeast2"
    }
    elseif($EnvironmentName -eq $AzureUSGovernment)
    {
        $defaultRegion = "usgovvirginia"
    }
    elseif($EnvironmentName -eq $AzureGermanCloud)
    {
        $defaultRegion = "germanynortheast"
    }
    elseif($EnvironmentName -eq $AzurePPE)
    {
        $defaultRegion = "westus"
    }
    elseif($EnvironmentName -eq $AzureCanary)
    {
        $defaultRegion = "eastus2euap"
    }

    return $defaultRegion
}

function Get-GraphAccessToken{
param(
    [string] $TenantId,
    [string] $EnvironmentName
    )

    # Below commands ensure there is graph access token in cache
    Get-AzADApplication -DisplayName SomeApp1 -ErrorAction Ignore | Out-Null

    $graphTokenItemResource = (Get-AzContext).Environment.GraphUrl

    $authFactory = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory
    $azContext = Get-AzContext
    $graphTokenItem = $authFactory.Authenticate($azContext.Account, $azContext.Environment, $azContext.Tenant.Id, $null, [Microsoft.Azure.Commands.Common.Authentication.ShowDialog]::Never, $null, $graphTokenItemResource)
    return $graphTokenItem.AccessToken
}

function Get-EnvironmentEndpoints{
param(
    [string] $EnvironmentName,
    [ref] $ServiceEndpoint,
    [ref] $Authority,
    [ref] $BillingServiceApiScope,
    [ref] $GraphServiceApiScope
    )

    if(($EnvironmentName -eq $AzureCloud) -or ($EnvironmentName -eq $AzureCanary))
    {
        $ServiceEndpoint.Value = $ServiceEndpointAzureCloud
        $Authority.Value = $AuthorityAzureCloud
        $BillingServiceApiScope.Value = $BillingServiceApiScopeAzureCloud
        $GraphServiceApiScope.Value = $GraphServiceApiScopeAzureCloud
    }
    elseif($EnvironmentName -eq $AzureChinaCloud)
    {
        $ServiceEndpoint.Value = $ServiceEndpointAzureChinaCloud
        $Authority.Value = $AuthorityAzureChinaCloud
        $BillingServiceApiScope.Value = $BillingServiceApiScopeAzureChinaCloud
        $GraphServiceApiScope.Value = $GraphServiceApiScopeAzureChinaCloud
    }
    elseif($EnvironmentName -eq $AzureUSGovernment)
    {
        $ServiceEndpoint.Value = $ServiceEndpointAzureUSGovernment
        $Authority.Value = $AuthorityAzureUSGovernment
        $BillingServiceApiScope.Value = $BillingServiceApiScopeAzureUSGovernment
        $GraphServiceApiScope.Value = $GraphServiceApiScopeAzureUSGovernment
    }
    elseif($EnvironmentName -eq $AzureGermanCloud)
    {
        $ServiceEndpoint.Value = $ServiceEndpointAzureGermanCloud
        $Authority.Value = $AuthorityAzureGermanCloud
        $BillingServiceApiScope.Value = $BillingServiceApiScopeAzureGermanCloud
        $GraphServiceApiScope.Value = $GraphServiceApiScopeAzureGermanCloud
    }
    elseif($EnvironmentName -eq $AzurePPE)
    {
        $ServiceEndpoint.Value = $ServiceEndpointAzurePPE
        $Authority.Value = $AuthorityAzurePPE
        $BillingServiceApiScope.Value = $BillingServiceApiScopeAzurePPE
        $GraphServiceApiScope.Value = $GraphServiceApiScopeAzurePPE
    }
}

function Get-PortalAppPermissionsPageUrl{
param(
    [string] $AppId,
    [string] $TenantId,
    [string] $EnvironmentName,
    [string] $Region
    )

    $portalBaseUrl = Get-PortalDomain -TenantId $TenantId -EnvironmentName $EnvironmentName -Region $Region
    $portalAADAppRelativeUrl = $PortalAADAppPermissionUrl -f $AppId
    return $portalBaseUrl + $portalAADAppRelativeUrl
}

function Get-PortalHCIResourcePageUrl{
param(
    [string] $TenantId,
    [string] $EnvironmentName,
    [string] $SubscriptionId,
    [string] $ResourceGroupName,
    [string] $ResourceName,
    [string] $Region
    )

    $portalBaseUrl = Get-PortalDomain -TenantId $TenantId -EnvironmentName $EnvironmentName -Region $Region
    $portalHCIResourceRelativeUrl = $PortalHCIResourceUrl -f $TenantId, $SubscriptionId, $ResourceGroupName, $ResourceName
    return $portalBaseUrl + $portalHCIResourceRelativeUrl
}

function Get-ResourceId{
param(
    [string] $ResourceName,
    [string] $SubscriptionId,
    [string] $ResourceGroupName
    )

    return "/Subscriptions/" + $SubscriptionId + "/resourceGroups/" + $ResourceGroupName + "/providers/Microsoft.AzureStackHCI/clusters/" + $ResourceName
}

function Get-RequiredResourceAccess{
    $requiredResourcesAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess] 

    $requiredAccess = New-Object Microsoft.Open.AzureAD.Model.RequiredResourceAccess
    $requiredAccess.ResourceAppId = $UsageServiceFirstPartyAppId
    $requiredAccess.ResourceAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]

    Foreach ($permId in $PermissionIds)
    {
        $permAccess = New-Object Microsoft.Open.AzureAD.Model.ResourceAccess
        $permAccess.Type = "Role"
        $permAccess.Id = $permId 
        $requiredAccess.ResourceAccess.Add($permAccess)
    }

    $requiredResourcesAccess.Add($requiredAccess)
    return $requiredResourcesAccess
}

# Called during repair registration.
function AddRequiredPermissionsIfNotPresent{
param(
    [string] $AppId
    )

    Write-Verbose "Adding the required permissions to AAD Application $AppId if not already added"
    $app = Retry-Command -ScriptBlock { Get-AzureADApplication -Filter "AppId eq '$AppId'"}
    $shouldAddRequiredPerms = $false

    if($app.RequiredResourceAccess -eq $Null)
    {
        $shouldAddRequiredPerms = $true
    }
    else
    {
        $reqResourceAccess = $app.RequiredResourceAccess | Where-Object {$_.ResourceAppId -eq $UsageServiceFirstPartyAppId}

        if($reqResourceAccess -eq $Null)
        {
            $shouldAddRequiredPerms = $true
        }
        else
        {
            if ($reqResourceAccess.ResourceAccess -eq $Null)
            {
                $shouldAddRequiredPerms = $true
            }
            else
            {
                $spReqPermClusRW = $reqResourceAccess.ResourceAccess | Where-Object {$_.Id -eq $ClusterReadWritePermission}
                $spReqPermClusR = $reqResourceAccess.ResourceAccess | Where-Object {$_.Id -eq $ClusterReadPermission}
                $spReqPermClusNodeRW = $reqResourceAccess.ResourceAccess | Where-Object {$_.Id -eq $ClusterNodeReadWritePermission}
                $spReqPermClusNodeR = $reqResourceAccess.ResourceAccess | Where-Object {$_.Id -eq $ClusterNodeReadPermission}

                # If App has these permissions, we have already added the required permissions earlier. Not need to add again.
                if(($spReqPermClusRW -ne $Null) -and ($spReqPermClusR -ne $Null) -and ($spReqPermClusNodeRW -ne $Null) -and ($spReqPermClusNodeR -ne $Null))
                {
                    $shouldAddRequiredPerms = $false
                }
                else
                {
                    $shouldAddRequiredPerms = $true
                }
            }
        }
    }

    # Add the required permissions
    if($shouldAddRequiredPerms -eq $true)
    {
        $requiredResourcesAccess = Get-RequiredResourceAccess
        Retry-Command -ScriptBlock { Set-AzureADApplication -ObjectId $app.ObjectId -RequiredResourceAccess $requiredResourcesAccess | Out-Null} -RetryIfNullOutput $false
    }
}

function Check-UsageAppRoles{
param(
    [string] $AppId
    )

    Write-Verbose "Checking admin consent status for AAD Application $AppId"

    $appSP = Retry-Command -ScriptBlock { Get-AzureADServicePrincipal -Filter "AppId eq '$AppId'"}

    # Try Get-AzureADServiceAppRoleAssignment as well to get app role assignments. WAC token falls under this case.
    $assignedPerms = Retry-Command -ScriptBlock { @(Get-AzureADServiceAppRoleAssignedTo -ObjectId $appSP.ObjectId) + @(Get-AzureADServiceAppRoleAssignment -ObjectId $appSP.ObjectId)} -RetryIfNullOutput $false

    $clusterRead = $assignedPerms | where { ($_.Id -eq $ClusterReadPermission) }
    $clusterReadWrite = $assignedPerms | where { ($_.Id -eq $ClusterReadWritePermission) }
    $clusterNodeRead = $assignedPerms | where { ($_.Id -eq $ClusterNodeReadPermission) }
    $clusterNodeReadWrite = $assignedPerms | where { ($_.Id -eq $ClusterNodeReadWritePermission) }

    $assignedPermsList = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.DirectoryObject]
    $assignedPermsList.Add($clusterRead)
    $assignedPermsList.Add($clusterReadWrite)
    $assignedPermsList.Add($clusterNodeRead)
    $assignedPermsList.Add($clusterNodeReadWrite)

    Foreach ($perm in $assignedPermsList)
    {
        if($perm -eq $null)
        {
            return $false
        }

        if($perm.DeletionTimestamp -ne $Null -and ($perm.DeletionTimestamp -gt $perm.CreationTimestamp))
        {
            return $false
        }
    }

    return $true
}

function Create-Application{
param(
    [string] $AppName
    )

    Write-Verbose "Creating AAD Application $AppName"
    # If the subscription is just registered to have HCI resources, sometimes it may take a while for the billing service principal to propogate

    $usagesp = Retry-Command -ScriptBlock { Get-AzureADServicePrincipal -Filter "AppId eq '$UsageServiceFirstPartyAppId'"}

    $requiredResourcesAccess = Get-RequiredResourceAccess

    # Create application
    $app = Retry-Command -ScriptBlock { New-AzureADApplication -DisplayName $AppName -RequiredResourceAccess $requiredResourcesAccess }
    $sp = Retry-Command -ScriptBlock { New-AzureADServicePrincipal -AppId $app.AppId }

    Write-Verbose "Created new AAD Application $app.AppId"

    return $app.AppId
}

function Grant-AdminConsent{
param(
    [string] $AppId
    )

    Write-Verbose "Granting admin consent for AAD Application Id $AppId"
    $usagesp = Retry-Command -ScriptBlock { Get-AzureADServicePrincipal -Filter "AppId eq '$UsageServiceFirstPartyAppId'"}
    $appSP = Retry-Command -ScriptBlock { Get-AzureADServicePrincipal -Filter "AppId eq '$AppId'"}

    try 
    {
        Retry-Command -ScriptBlock { New-AzureADServiceAppRoleAssignment -ObjectId $appSP.ObjectId -PrincipalId $appSP.ObjectId -ResourceId $usagesp.ObjectId -Id $ClusterReadPermission} -RetryIfNullOutput $false
        Retry-Command -ScriptBlock { New-AzureADServiceAppRoleAssignment -ObjectId $appSP.ObjectId -PrincipalId $appSP.ObjectId -ResourceId $usagesp.ObjectId -Id $ClusterReadWritePermission} -RetryIfNullOutput $false
        Retry-Command -ScriptBlock { New-AzureADServiceAppRoleAssignment -ObjectId $appSP.ObjectId -PrincipalId $appSP.ObjectId -ResourceId $usagesp.ObjectId -Id $ClusterNodeReadPermission} -RetryIfNullOutput $false
        Retry-Command -ScriptBlock { New-AzureADServiceAppRoleAssignment -ObjectId $appSP.ObjectId -PrincipalId $appSP.ObjectId -ResourceId $usagesp.ObjectId -Id $ClusterNodeReadWritePermission} -RetryIfNullOutput $false
    }
    catch 
    {
        Write-Debug "Exception occurred when granting admin consent"
        $ErrorMessage = $_.Exception.Message
        Write-Debug $ErrorMessage
        return $False
    }

    return $True
}

function Remove-OldScopes{
param(
    [string] $AppId
    )

    Write-Verbose "Removing old scopes on AAD Application with Id $AppId"
    $appSP = Retry-Command -ScriptBlock { Get-AzureADServicePrincipal -Filter "AppId eq '$AppId'"}

    # Remove AzureStackHCI.Billing.Sync and AzureStackHCI.Census.Sync permissions if present as we dont need them
    $assignedPerms = Retry-Command -ScriptBlock { Get-AzureADServiceAppRoleAssignedTo -ObjectId $appSP.ObjectId} -RetryIfNullOutput $false

    $billingSync = $assignedPerms | where { ($_.Id -eq $BillingSyncPermission) }
    $censusSync = $assignedPerms | where { ($_.Id -eq $CensusSyncPermission) }

    if($billingSync -eq $Null -or $censusSync -eq $Null)
    {
        # Try Get-AzureADServiceAppRoleAssignment as well to get app role assignments. WAC token falls under this case.
        $assignedPerms = Retry-Command -ScriptBlock { Get-AzureADServiceAppRoleAssignment -ObjectId $appSP.ObjectId} -RetryIfNullOutput $false
    }

    if($billingSync -eq $Null)
    {
        $billingSync = $assignedPerms | where { ($_.Id -eq $BillingSyncPermission) }
    }

    if($censusSync -eq $Null)
    {
        $censusSync = $assignedPerms | where { ($_.Id -eq $CensusSyncPermission) }
    }

    if($billingSync -ne $Null)
    {
        Retry-Command -ScriptBlock { Remove-AzureADServiceAppRoleAssignment -ObjectId $appSP.ObjectId -AppRoleAssignmentId $billingSync.ObjectId | Out-Null} -RetryIfNullOutput $false
    }

    if($censusSync -ne $Null)
    {
        Retry-Command -ScriptBlock { Remove-AzureADServiceAppRoleAssignment -ObjectId $appSP.ObjectId -AppRoleAssignmentId $censusSync.ObjectId | Out-Null} -RetryIfNullOutput $false
    }
}

function Azure-Login{
param(
    [string] $SubscriptionId,
    [string] $TenantId,
    [string] $ArmAccessToken,
    [string] $GraphAccessToken,
    [string] $AccountId,
    [string] $EnvironmentName,
    [string] $ProgressActivityName,
    [string] $Region,
    [bool]   $UseDeviceAuthentication
    )

    Write-Progress -Id $MainProgressBarId -activity $ProgressActivityName -status $InstallAzResourcesMessage -percentcomplete 10

    try
    {
        Import-Module -Name Az.Resources -ErrorAction Stop
    }
    catch
    {
        try
        {
            Import-PackageProvider -Name Nuget -MinimumVersion "2.8.5.201" -ErrorAction Stop
        }
        catch
        {
            Install-PackageProvider NuGet -Force | Out-Null
        }

        Install-Module -Name Az.Resources -Force -AllowClobber
        Import-Module -Name Az.Resources
    }

    Write-Progress -Id $MainProgressBarId -activity $ProgressActivityName -status $InstallAzureADMessage -percentcomplete 20

    try
    {
        Import-Module -Name AzureAD -ErrorAction Stop
    }
    catch
    {
        Install-Module -Name AzureAD -Force -AllowClobber
        Import-Module -Name AzureAD
    }

    Write-Progress -Id $MainProgressBarId -activity $ProgressActivityName -status $LoggingInToAzureMessage -percentcomplete 30

    if($EnvironmentName -eq $AzurePPE)
    {
        Add-AzEnvironment -Name $AzurePPE -PublishSettingsFileUrl "https://windows.azure-test.net/publishsettings/index" -ServiceEndpoint "https://management-preview.core.windows-int.net/" -ManagementPortalUrl "https://windows.azure-test.net/" -ActiveDirectoryEndpoint "https://login.windows-ppe.net/" -ActiveDirectoryServiceEndpointResourceId "https://management.core.windows.net/" -ResourceManagerEndpoint "https://api-dogfood.resources.windows-int.net/" -GalleryEndpoint "https://df.gallery.azure-test.net/" -GraphEndpoint "https://graph.ppe.windows.net/" -GraphAudience "https://graph.ppe.windows.net/" | Out-Null
    }

    $ConnectAzureADEnvironmentName = $EnvironmentName
    $ConnectAzAccountEnvironmentName = $EnvironmentName

    if($EnvironmentName -eq $AzureCanary)
    {
        $ConnectAzureADEnvironmentName = $AzureCloud

        if([string]::IsNullOrEmpty($Region))
        {
            $Region = Get-DefaultRegion -EnvironmentName $EnvironmentName
        }

        # Normalize region name
        $Region = Normalize-RegionName -Region $Region

        $ConnectAzAccountEnvironmentName = ($AzureCanary + $Region)

        $azEnv = (Get-AzEnvironment -Name $AzureCloud)
        $azEnv.Name = $ConnectAzAccountEnvironmentName
        $azEnv.ResourceManagerUrl = ('https://{0}.management.azure.com/' -f $Region)
        $azEnv | Add-AzEnvironment | Out-Null
    }

    Disconnect-AzAccount -ErrorAction Ignore | Out-Null

    if([string]::IsNullOrEmpty($ArmAccessToken) -or [string]::IsNullOrEmpty($GraphAccessToken) -or [string]::IsNullOrEmpty($AccountId))
    {
        # Interactive login

        $IsIEPresent = Test-Path "$env:SystemRoot\System32\ieframe.dll"

        if([string]::IsNullOrEmpty($TenantId))
        {
            if(($UseDeviceAuthentication -eq $false) -and ($IsIEPresent))
            {
                Connect-AzAccount -Environment $ConnectAzAccountEnvironmentName -SubscriptionId $SubscriptionId -Scope Process | Out-Null
            }
            else # Use -UseDeviceAuthentication as IE Frame is not available to show Azure login popup
            {
                Write-Progress -Id $MainProgressBarId -activity $ProgressActivityName -Completed # Hide progress activity as it blocks the console output
                Connect-AzAccount -Environment $ConnectAzAccountEnvironmentName -SubscriptionId $SubscriptionId -UseDeviceAuthentication -Scope Process | Out-Null
            }
        }
        else
        {
            if(($UseDeviceAuthentication -eq $false) -and ($IsIEPresent))
            {
                Connect-AzAccount -Environment $ConnectAzAccountEnvironmentName -TenantId $TenantId -SubscriptionId $SubscriptionId -Scope Process | Out-Null
            }
            else # Use -UseDeviceAuthentication as IE Frame is not available to show Azure login popup
            {
                Write-Progress -Id $MainProgressBarId -activity $ProgressActivityName -Completed # Hide progress activity as it blocks the console output
                Connect-AzAccount -Environment $ConnectAzAccountEnvironmentName -TenantId $TenantId -SubscriptionId $SubscriptionId -UseDeviceAuthentication -Scope Process | Out-Null
            }
        }

        Write-Progress -Id $MainProgressBarId -activity $ProgressActivityName -status $ConnectingToAzureAD -percentcomplete 35

        $azContext = Get-AzContext
        $TenantId = $azContext.Tenant.Id
        $AccountId = $azContext.Account.Id
        $GraphAccessToken = Get-GraphAccessToken -TenantId $TenantId -EnvironmentName $EnvironmentName

        Connect-AzureAD -AzureEnvironmentName $ConnectAzureADEnvironmentName -TenantId $TenantId -AadAccessToken $GraphAccessToken -AccountId $AccountId | Out-Null
    }
    else
    {
        # Not an interactive login

        if([string]::IsNullOrEmpty($TenantId))
        {
            Connect-AzAccount -Environment $ConnectAzAccountEnvironmentName -SubscriptionId $SubscriptionId -AccessToken $ArmAccessToken -AccountId $AccountId -GraphAccessToken $GraphAccessToken -Scope Process | Out-Null
        }
        else
        {
            Connect-AzAccount -Environment $ConnectAzAccountEnvironmentName -TenantId $TenantId -SubscriptionId $SubscriptionId -AccessToken $ArmAccessToken -AccountId $AccountId -GraphAccessToken $GraphAccessToken -Scope Process | Out-Null
        }

        Write-Progress -Id $MainProgressBarId -activity $ProgressActivityName -status $ConnectingToAzureAD -percentcomplete 35

        $azContext = Get-AzContext
        $TenantId = $azContext.Tenant.Id
        Connect-AzureAD -AzureEnvironmentName $ConnectAzureADEnvironmentName -TenantId $TenantId -AadAccessToken $GraphAccessToken -AccountId $AccountId | Out-Null
    }

    return $TenantId
}

function Normalize-RegionName{
param(
    [string] $Region
    )
    $regionName = $Region -replace '\s',''
    $regionName = $regionName.ToLower()
    return $regionName
}

function Validate-RegionName{
param(
    [string] $Region,
    [ref] $SupportedRegions
    )
    $resources = Retry-Command -ScriptBlock { Get-AzResourceProvider -ProviderNamespace Microsoft.AzureStackHCI } -RetryIfNullOutput $true
    $locations = $resources.Where{($_.ResourceTypes.ResourceTypeName -eq 'clusters' -and $_.RegistrationState -eq 'Registered')}.Locations

    $locations | foreach {
        $regionName = Normalize-RegionName -Region $_
        if ($regionName -eq $Region)
        {
            # Supported region

            return $True
        }
    }

    $SupportedRegions.value = $locations -join ','
    return $False
}

function Get-ClusterDNSSuffix{
param(
    [System.Management.Automation.Runspaces.PSSession] $Session
    )

    $clusterNameResourceGUID = Invoke-Command -Session $Session -ScriptBlock { (Get-ItemProperty -Path HKLM:\Cluster -Name ClusterNameResource).ClusterNameResource }
    $clusterDNSSuffix = Invoke-Command -Session $Session -ScriptBlock { (Get-ClusterResource $using:clusterNameResourceGUID | Get-ClusterParameter DnsSuffix).Value }
    return $clusterDNSSuffix
}

function Get-ClusterDNSName{
param(
    [System.Management.Automation.Runspaces.PSSession] $Session
    )

    $clusterNameResourceGUID = Invoke-Command -Session $Session -ScriptBlock { (Get-ItemProperty -Path HKLM:\Cluster -Name ClusterNameResource).ClusterNameResource }
    $clusterDNSName = Invoke-Command -Session $Session -ScriptBlock { (Get-ClusterResource $using:clusterNameResourceGUID | Get-ClusterParameter DnsName).Value }
    return $clusterDNSName
}

function Check-ConnectionToCloudBillingService{
param(
    $ClusterNodes,
    [System.Management.Automation.PSCredential] $Credential,
    [string] $HealthEndpoint,
    [System.Collections.ArrayList] $HealthEndPointCheckFailedNodes,
    [string] $ClusterDNSSuffix
    )

    Foreach ($clusNode in $ClusterNodes)
    {
        $nodeSession = $null

        try
        {
            if($Credential -eq $Null)
            {
                $nodeSession = New-PSSession -ComputerName ($clusNode.Name + "." + $ClusterDNSSuffix)
            }
            else
            {
                $nodeSession = New-PSSession -ComputerName ($clusNode.Name + "." + $ClusterDNSSuffix) -Credential $Credential
            }

            # Check if node can reach cloud billing service
            $healthResponse = Invoke-Command -Session $nodeSession -ScriptBlock { Invoke-WebRequest $Using:HealthEndpoint -UseBasicParsing}

            if(($healthResponse -eq $Null) -or ($healthResponse.StatusCode -ne [int][system.net.httpstatuscode]::ok))
            {
                Write-Verbose ("StatusCode of invoking cloud billing service health endpoint on node " + $clusNode.Name + " : " + $healthResponse.StatusCode)
                $HealthEndPointCheckFailedNodes.Add($clusNode.Name) | Out-Null
                continue
            }
        }
        catch
        {
            Write-Verbose ("Exception occurred while testing health endpoint connectivity on Node: " + $clusNode.Name + " Exception: " + $_.Exception)
            $HealthEndPointCheckFailedNodes.Add($clusNode.Name) | Out-Null
            continue
        }
    }
}

function Setup-Certificates{
param(
    $ClusterNodes,
    [System.Management.Automation.PSCredential] $Credential,
    [string] $ResourceName,
    [string] $ObjectId,
    [string] $CertificateThumbprint,
    [string] $AppId,
    [string] $TenantId,
    [string] $CloudId,
    [string] $ServiceEndpoint,
    [string] $BillingServiceApiScope,
    [string] $GraphServiceApiScope,
    [string] $Authority,
    [System.Collections.ArrayList] $NewCertificateFailedNodes,
    [System.Collections.ArrayList] $SetCertificateFailedNodes,
    [System.Collections.ArrayList] $OSNotLatestOnNodes,
    [System.Collections.HashTable] $CertificatesToBeMaintained,
    [string] $ClusterDNSSuffix
    )

    $userProvidedCertAddedToAAD = $false
    Foreach ($clusNode in $ClusterNodes)
    {
        $nodeSession = $null

        try
        {
            if($Credential -eq $Null)
            {
                $nodeSession = New-PSSession -ComputerName ($clusNode.Name + "." + $ClusterDNSSuffix)
            }
            else
            {
                $nodeSession = New-PSSession -ComputerName ($clusNode.Name + "." + $ClusterDNSSuffix) -Credential $Credential
            }
        }
        catch
        {
            Write-Debug ("Exception occurred in establishing new PSSession. ErrorMessage : " + $_.Exception.Message)
            Write-Debug $_
            $NewCertificateFailedNodes.Add($clusNode.Name) | Out-Null
            $SetCertificateFailedNodes.Add($clusNode.Name) | Out-Null
            continue
        }

        # Check if all nodes have required OS version
        $nodeUBR = Invoke-Command -Session $nodeSession -ScriptBlock { (Get-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").UBR }
        $nodeBuildNumber = Invoke-Command -Session $nodeSession -ScriptBlock { (Get-CimInstance -ClassName CIM_OperatingSystem).BuildNumber }

        if(($nodeBuildNumber -lt $GAOSBuildNumber) -or (($nodeBuildNumber -eq $GAOSBuildNumber) -and ($nodeUBR -lt $GAOSUBR)))
        {
            $OSNotLatestOnNodes.Add($clusNode.Name) | Out-Null
            continue
        }

        if([string]::IsNullOrEmpty($CertificateThumbprint))
        {
            # User did not specify certificate, using self-signed certificate
            try
            {
                $certBase64 = Invoke-Command -Session $nodeSession -ScriptBlock { New-AzureStackHCIRegistrationCertificate }
            }
            catch
            {
                Write-Debug ("Exception occurred in New-AzureStackHCIRegistrationCertificate. ErrorMessage : " + $_.Exception.Message)
                Write-Debug $_
                $NewCertificateFailedNodes.Add($clusNode.Name) | Out-Null
                continue
            }
        }
        else
        {
            # Get certificate from cert store.
            $x509Cert = $Null;
            try
            {
                $x509Cert = Invoke-Command -Session $nodeSession -ScriptBlock { Get-ChildItem Cert:\LocalMachine -Recurse | Where { $_.Thumbprint -eq $Using:CertificateThumbprint} | Select-Object -First 1}
            }
            catch{}

            # Certificate not found on node
            if($x509Cert -eq $Null)
            {
                $CertificateNotFoundErrorMessage = $CertificateNotFoundOnNode -f $CertificateThumbprint,$clusNode.Name
                return $CertificateNotFoundErrorMessage
            }

            # Certificate should be valid for atleast 60 days from now
            $60days = New-TimeSpan -Days 60
            $expectedValidTo = (Get-Date) + $60days

            if($x509Cert.NotAfter -lt $expectedValidTo)
            {
                $UserCertificateValidationErrorMessage = ($UserCertValidationErrorMessage -f $CertificateThumbprint, $x509Cert.NotAfter)
                return $UserCertificateValidationErrorMessage
            }

            $certBase64 = [System.Convert]::ToBase64String($x509Cert.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert))
        }

        $Cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]([System.Convert]::FromBase64String($CertBase64))

        # If user provided cert is not already added to AAD app or if we are using one certificate per node
        if(($userProvidedCertAddedToAAD -eq $false) -or ([string]::IsNullOrEmpty($CertificateThumbprint)))
        {
            $AddAppCredentialMessageProgress = $AddAppCredentialMessage -f $ResourceName
            Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $AddAppCredentialMessageProgress -percentcomplete 80
            $now = [System.DateTime]::UtcNow
            $appCredential = Retry-Command -ScriptBlock { New-AzureADApplicationKeyCredential -ObjectId $ObjectId -Type AsymmetricX509Cert -Usage Verify -Value $CertBase64 -StartDate $now -EndDate $Cert.NotAfter}
            $CertificatesToBeMaintained.Add($appCredential.KeyId, $true)
            $userProvidedCertAddedToAAD = $true

            # Wait till the credential added is returned in Get to avoid the rush in using this new certificate. Gives more time for AAD replication before using certificate.
            $credReturned = Retry-Command -ScriptBlock { Get-AzureADApplicationKeyCredential -ObjectId $ObjectId | where {($_.KeyId -eq $appCredential.KeyId)} }
        }

        # Set the certificate - Certificate will be set after testing the certificate by calling cloud service API
        try
        {
            $SetCertParams = @{
                        ServiceEndpoint = $ServiceEndpoint
                        BillingServiceApiScope = $BillingServiceApiScope
                        GraphServiceApiScope = $GraphServiceApiScope
                        AADAuthority = $Authority
                        AppId = $AppId
                        TenantId = $TenantId
                        CloudId = $CloudId
                        CertificateThumbprint = $CertificateThumbprint
                    }

            Invoke-Command -Session $nodeSession -ScriptBlock { Set-AzureStackHCIRegistrationCertificate @Using:SetCertParams }
        }
        catch
        {
            Write-Debug ("Exception occurred in Set-AzureStackHCIRegistrationCertificate. ErrorMessage : " + $_.Exception.Message)
            Write-Debug $_
            $SetCertificateFailedNodes.Add($clusNode.Name) | Out-Null
            continue
        }
    }

    return $null
}

function Create-ArcApplication{
param(
    [string] $AppName,
    [ref] $AppId,
    [ref] $Secret,
    [Switch] $IsWAC
    )
    Write-Verbose "Creating AAD Arc Agent Application $AppName"

    # Create application and service principal for Arc agent app.
    $app = Retry-Command -ScriptBlock { New-AzureADApplication -DisplayName $AppName }
    $sp = Retry-Command -ScriptBlock { New-AzureADServicePrincipal -AppId $app.AppId }

    # Usually takes 10-20 seconds after app creation to be able to assign roles. Can take upto 140 seconds for AAD to replicate in rare cases.
    $stopLoop = $false
    [int]$retryCount = "0"
    [int]$maxRetryCount = "14"

    do {
        try 
        {
            New-AzRoleAssignment -ObjectId $sp.ObjectId -RoleDefinitionName $AzureConnectedMachineOnboardingRole | Out-Null
            New-AzRoleAssignment -ObjectId $sp.ObjectId -RoleDefinitionName $AzureConnectedMachineResourceAdministratorRole | Out-Null
            $stopLoop = $true
            break
        }
        catch
        {
            $positionMessage = $_.InvocationInfo.PositionMessage

            if ($retryCount -ge $maxRetryCount)
            {
                # Cleanup
                Remove-AzureADApplication -ObjectId $app.ObjectId

                # Timed out.
                Write-Debug ("Failed to assign roles to service principal with App Id $($app.AppId). ErrorMessage: " + $_.Exception.Message + " PositionalMessage: " + $positionMessage)
                throw $_
            }

            if ($_.Exception.Message.Contains("Microsoft.Authorization/roleAssignments/write"))
            {
                # Cleanup
                Remove-AzureADApplication -ObjectId $app.ObjectId

                Write-Verbose "Create-ArcApplication Missing Permissions. IsWAC: $IsWAC"

                if($IsWAC -eq $false)
                {
                    # Insufficient privilige error.
                    Write-Error -Message $ArcAgentRolesInsufficientPreviligeMessage -ErrorAction Continue
                }

                return $false
            }
            
            # Service principal creation hasn't propogated fully yet, usually takes 10-20 seconds.
            Write-Verbose "Could not assign roles to service principal with App Id $($app.AppId). Retrying in 10 seconds..."
            Start-Sleep -Seconds 10
            $retryCount = $retryCount + 1
        }
    }
    While (-Not $stopLoop)

    # Set password to never expire.
    $start = Get-Date
    $end = $start.AddYears($IndefinitelyYears)
    $pw = Retry-Command -ScriptBlock { New-AzureADServicePrincipalPasswordCredential -ObjectId $sp.ObjectId -StartDate $start -EndDate $end }

    Write-Verbose "Created new AAD Arc Agent Application $($app.AppId)"
    $AppId.Value = $app.AppId
    $Secret.Value = $pw.Value
    return $true
}

function Remove-ArcApplication{
    param(
        [string] $AppId
        )
    
        $app = Get-AzureADApplication -Filter "AppId eq '$AppId'"
        $sp = Get-AzureADServicePrincipal -Filter "AppId eq '$AppId'"
        
        # Remove assigned roles.
        try
        {
            $roles = Get-AzRoleAssignment -ObjectId $sp.ObjectId
            
            $roles | ForEach-Object { 
                $roleName = $_.RoleDefinitionName
                if (($roleName -eq $AzureConnectedMachineOnboardingRole) -or ($roleName -eq $AzureConnectedMachineResourceAdministratorRole))
                {
                    Remove-AzRoleAssignment -ObjectId $sp.ObjectId -RoleDefinitionName $roleName | Out-Null
                }
            }
        }
        catch
        {
            Write-Warning -Message $ArcRolesCleaningWarningMessage
            Write-Debug ("Exception occurred in clearing roles on service principal with App Id {0}. ErrorMessage : {1}" -f ($AppId), ($_.Exception.Message))
            Write-Debug $_
        }

        # Delete application.
        Remove-AzureADApplication -ObjectId $app.ObjectId
}

function Enable-ArcForServers{
param(
    [System.Management.Automation.Runspaces.PSSession] $Session,
    [System.Management.Automation.PSCredential] $Credential,
    [string] $ClusterDNSSuffix
    )
    # Create new sessions for all nodes in cluster.
    $clusterNodeNames = Invoke-Command -Session $Session -ScriptBlock { Get-ClusterNode } | ForEach-Object { ($_.Name + "." + $ClusterDNSSuffix) }
    if($Credential -eq $Null)
    {
        $clusterNodeSessions = New-PSSession -ComputerName $clusterNodeNames
    }
    else
    {
        $clusterNodeSessions = New-PSSession -ComputerName $clusterNodeNames -Credential $Credential
    }

    $retStatus = [ErrorDetail]::Success

    # Start running
    try
    {
        Invoke-Command -Session $clusterNodeSessions -ScriptBlock {
            # Cluster scheduled task is triggered asynchronously. Use Get-ScheduledTask to get the task state and wait for its completion.
            Get-ScheduledTask -TaskName $using:ArcRegistrationTaskName | Start-ScheduledTask

            Start-Sleep -Seconds $using:ClusterScheduledTaskSleepTimeSeconds
            $limit = (Get-Date).AddMinutes($using:ClusterScheduledTaskWaitTimeMinutes)

            while ((Get-ScheduledTask -TaskName $using:ArcRegistrationTaskName).State -eq $using:ClusterScheduledTaskRunningState -and (Get-Date) -lt $limit) {
                Start-Sleep -Seconds $using:ClusterScheduledTaskSleepTimeSeconds
            }

            if((Get-ScheduledTask -TaskName $using:ArcRegistrationTaskName).State -ne $using:ClusterScheduledTaskReadyState)
            {
                throw ("Cluster scheduled task runtime exceeded the max configured wait time of {0} minutes" -f ($using:ClusterScheduledTaskWaitTimeMinutes))
            }
        }

        # Show warning if any of the nodes failed to register with Arc
        $enabledArcStatus = [ArcStatus]::Enabled
        Invoke-Command -Session $Session -ScriptBlock {
            $nodeStatus = $(Get-AzureStackHCIArcIntegration).NodesArcStatus

            if ($nodeStatus -ne $null -and $nodeStatus.Count -ge $clusterNodeNames.Count)
            {
                Foreach ($node in $nodeStatus.Keys)
                {
                    if($nodeStatus[$node] -ne $using:enabledArcStatus)
                    {
                        Write-Warning -Message $using:RegisterArcFailedWarningMessage
                        $retStatus = [ErrorDetail]::ArcIntegrationFailedOnNodes
                        break
                    }
                }
            }
            else
            {
                Write-Warning -Message $using:RegisterArcFailedWarningMessage
                $retStatus = [ErrorDetail]::ArcIntegrationFailedOnNodes
            }
        }
    }
    catch
    {
        Write-Warning -Message $RegisterArcFailedWarningMessage
        $retStatus = [ErrorDetail]::ArcIntegrationFailedOnNodes
        Write-Debug ("Exception occurred in registering nodes to Arc For Servers. ErrorMessage : {0}" -f ($_.Exception.Message))
        Write-Debug $_
    }

    # Cleanup sessions.
    Remove-PSSession $clusterNodeSessions | Out-Null

    return $retStatus
}

function Disable-ArcForServers{
param(
    [System.Management.Automation.Runspaces.PSSession] $Session,
    [System.Management.Automation.PSCredential] $Credential,
    [string] $ClusterDNSSuffix
    )

    $res = $true
    $AgentUninstaller_LogFile = "ConnectedMachineAgentUninstallationLog.txt";
    $AgentInstaller_Name      = "AzureConnectedMachineAgent.msi";
    $AgentExecutable_Path     = $Env:Programfiles + '\AzureConnectedMachineAgent\azcmagent.exe'

    $clusterNodeNames = Invoke-Command -Session $Session -ScriptBlock { Get-ClusterNode } | ForEach-Object { ($_.Name + "." + $ClusterDNSSuffix) }
    if($Credential -eq $Null)
    {
        $clusterNodeSessions = New-PSSession -ComputerName $clusterNodeNames
    }
    else
    {
        $clusterNodeSessions = New-PSSession -ComputerName $clusterNodeNames -Credential $Credential
    }

    $nodeArcStatus = Invoke-Command -Session $Session -ScriptBlock { $(Get-AzureStackHCIArcIntegration)}
    if($nodeArcStatus.ClusterArcStatus -eq [ArcStatus]::Disabled)
    {
        return $res
    }

    $disableFailedOnANode = $false

    try
    {
        Invoke-Command -Session $clusterNodeSessions -ScriptBlock {
            Disable-AzureStackHCIArcIntegration -AgentUninstallerLogFile $using:AgentUninstaller_LogFile -AgentInstallerName $using:AgentInstaller_Name -AgentExecutablePath $using:AgentExecutable_Path
        }
    }
    catch
    {
        $positionMessage = $_.InvocationInfo.PositionMessage
        Write-Debug ("Exception occurred in un-registering nodes from Arc For Servers. ErrorMessage: " + $_.Exception.Message + " PositionalMessage: " + $positionMessage)
        Write-Debug $_
        $disableFailedOnANode = $true
    }

    if ($disableFailedOnANode -eq $true)
    {
        $nodeStatus = Invoke-Command -Session $Session -ScriptBlock { $(Get-AzureStackHCIArcIntegration).NodesArcStatus }
        foreach ($node in $nodeStatus.Keys)
        {
            if ($nodeStatus[$node] -ne [ArcStatus]::Disabled)
            {
                $res = $false
                $UnregisterArcFailedErrorMessage = $UnregisterArcFailedError -f $node
                Write-Error -Message $UnregisterArcFailedErrorMessage -ErrorAction Continue
            }
        }
    }

    # Cleanup sessions.
    Remove-PSSession $clusterNodeSessions | Out-Null
    return $res
}

function Register-ArcForServers{
param(
    [bool] $IsManagementNode,
    [string] $ComputerName,
    [System.Management.Automation.PSCredential] $Credential,
    [string] $TenantId,
    [string] $SubscriptionId,
    [string] $ResourceGroup,
    [string] $Region,
    [string] $AppName,
    [string] $ClusterDNSSuffix,
    [Switch] $IsWAC,
    [string] $Environment
    )

    if($IsManagementNode)
    {
        if($Credential -eq $Null)
        {
            $session = New-PSSession -ComputerName $ComputerName
        }
        else
        {
            $session = New-PSSession -ComputerName $ComputerName -Credential $Credential
        }
    }
    else
    {
        $session = New-PSSession -ComputerName localhost
    }

    $clusterName = Invoke-Command -Session $session -ScriptBlock { (Get-Cluster).Name }
    $clusterDNSName = Get-ClusterDNSName -Session $session

    Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $RegisterArcProgressActivityName -Status $FetchingRegistrationState -PercentComplete 1
    $arcStatus = Invoke-Command -Session $session -ScriptBlock { Get-AzureStackHCIArcIntegration }

    $AppId = [string]::Empty
    $Secret = [string]::Empty

    # Register resource providers
    Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $RegisterArcProgressActivityName -Status $RegisterArcRPMessage -PercentComplete 10
    Register-AzResourceProvider -ProviderNamespace Microsoft.HybridCompute | Out-Null
    Register-AzResourceProvider -ProviderNamespace Microsoft.GuestConfiguration | Out-Null

    $arcAppId = $arcStatus.ApplicationId
    $shouldCreateApp = $true

    if(-Not [string]::IsNullOrEmpty($arcAppId))
    {
        $existingArcApp = Retry-Command -ScriptBlock { Get-AzureADApplication -Filter "AppId eq '$arcAppId'" } -RetryIfNullOutput $false

         if(($existingArcApp -ne $Null) -and ($arcStatus.ClusterArcStatus -eq [ArcStatus]::Enabled))
         {
            $shouldCreateApp = $false
         }
    }

    # Setup Arc agent.
    if ($shouldCreateApp)
    {
        # Create application.
        $CreatingAADAppMessageProgress = $CreatingAADAppMessage -f $AppName, $TenantId
        Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $RegisterArcProgressActivityName -Status $CreatingAADAppMessageProgress -PercentComplete 30
        $appCreated = Create-ArcApplication -AppName $AppName -AppId ([ref]$appId) -Secret ([ref]$secret) -IsWAC:$IsWAC

        if($appCreated -eq $false)
        {
            return [ErrorDetail]::ArcPermissionsMissing
        }

        $arcCommand = Invoke-Command -Session $session -ScriptBlock { Get-Command -Name 'Initialize-AzureStackHCIArcIntegration' -ErrorAction SilentlyContinue }
        if ($arcCommand.Parameters.ContainsKey('Cloud')) 
        {
            $arcEnvironment = $Environment
            if( $Environment -eq $AzureCanary)
            {
                $arcEnvironment = $AzureCloud
            }
            Write-Debug ("invoking Initialize-AzureStackHCIArcIntegration with cloud switch")
            $ArcRegistrationParams = @{
                AppId = $AppId
                Secret = $Secret
                TenantId = $TenantId
                SubscriptionId = $SubscriptionId
                Region = $Region
                ResourceGroup = $ResourceGroup 
                cloud  = $arcEnvironment 
            }    
        }
        else
        {
            Write-Debug ("invoking Initialize-AzureStackHCIArcIntegration without cloud switch")
            $ArcRegistrationParams = @{
                AppId = $AppId
                Secret = $Secret
                TenantId = $TenantId
                SubscriptionId = $SubscriptionId
                Region = $Region
                ResourceGroup = $ResourceGroup
            }
        }
        
        # Save Arc context.
        Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $RegisterArcProgressActivityName -Status $SetupArcMessage -PercentComplete 40
        
        Invoke-Command -Session $session -ScriptBlock { Initialize-AzureStackHCIArcIntegration @Using:ArcRegistrationParams }
    }

    # Register clustered scheduled task
    try
    {
        # Connect to cluster and use that session for registering clustered scheduled task
        if($Credential -eq $Null)
        {
            $clusterNameSession = New-PSSession -ComputerName ($clusterDNSName + "." + $ClusterDNSSuffix)
        }
        else
        {
            $clusterNameSession = New-PSSession -ComputerName ($clusterDNSName + "." + $ClusterDNSSuffix) -Credential $Credential
        }

        Invoke-Command -Session $clusterNameSession -ScriptBlock { 
            $task =  Get-ScheduledTask -TaskName $using:ArcRegistrationTaskName -ErrorAction SilentlyContinue
            $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-Command $using:registerArcScript"
            
            # Repeat the script every hour of every day, starting from now.
            $date = Get-Date
            $dailyTrigger = New-ScheduledTaskTrigger -Daily -At $date
            $hourlyTrigger = New-ScheduledTaskTrigger -Once -At $date -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Hours 23 -Minutes 55)
            $dailyTrigger.Repetition = $hourlyTrigger.Repetition

            if (-Not $task)
            {
                Register-ClusteredScheduledTask -TaskName $using:ArcRegistrationTaskName -TaskType ClusterWide -Action $action -Trigger $dailyTrigger -Cluster $Using:clusterName
            }
            else
            {
                # Update cluster schedule task.
                Set-ClusteredScheduledTask -TaskName $using:ArcRegistrationTaskName -Action $action -Trigger $dailyTrigger -Cluster $Using:clusterName
            }
        } | Out-Null
    }
    catch
    {
        $positionMessage = $_.InvocationInfo.PositionMessage
        Write-Error ("Exception occurred in registering cluster scheduled task. ErrorMessage: " + $_.Exception.Message + " PositionalMessage: " + $positionMessage) -Category OperationStopped
        throw
    }
    finally
    {
        if($clusterNameSession -ne $null)
        {
            Remove-PSSession $clusterNameSession -ErrorAction Ignore | Out-Null
        }
    }

    # Run
    Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $RegisterArcProgressActivityName -Status $StartingArcAgentMessage -PercentComplete 50
    $arcResult = Enable-ArcForServers -Session $session -Credential $Credential -ClusterDNSSuffix $ClusterDNSSuffix

    Write-Progress -Id $ArcProgressBarId -activity $RegisterArcProgressActivityName -Completed

    Remove-PSSession $session | Out-Null

    Write-Verbose "Successfully registered cluster with Arc for Servers."

    return $arcResult
}

function Unregister-ArcForServers{
param(
    [bool] $IsManagementNode,
    [string] $ComputerName,
    [System.Management.Automation.PSCredential] $Credential,
    [string] $ResourceId,
    [Switch] $Force,
    [string] $ClusterDNSSuffix
    )

    if($IsManagementNode)
    {
        if($Credential -eq $Null)
        {
            $session = New-PSSession -ComputerName $ComputerName
        }
        else
        {
            $session = New-PSSession -ComputerName $ComputerName -Credential $Credential
        }
    }
    else
    {
        $session = New-PSSession -ComputerName localhost
    }

    $clusterName = Invoke-Command -Session $session -ScriptBlock { (Get-Cluster).Name }
    $clusterDNSName = Get-ClusterDNSName -Session $session

    $cmdlet = Invoke-Command -Session $session -ScriptBlock { Get-Command Get-AzureStackHCIArcIntegration -Type Cmdlet -ErrorAction Ignore }

    if($cmdlet -eq $null)
    {
        return $true
    }

    Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $UnregisterArcProgressActivityName -Status $FetchingRegistrationState -PercentComplete 1
    $arcStatus = Invoke-Command -Session $session -ScriptBlock { Get-AzureStackHCIArcIntegration }
    $hciStatus = Invoke-Command -Session $session -ScriptBlock { Get-AzureStackHCI }
    $arcResourceId = $ResourceId + $HCIArcInstanceName
    $arcResourceExtensions = $arcResourceId + $HCIArcExtensions

    if ($arcStatus.ClusterArcStatus -eq [ArcStatus]::Enabled)
    {
        Invoke-Command -Session $session -ScriptBlock { Clear-AzureStackHCIArcIntegration -SetDisableInProgress}
    }

    $arcres = Get-AzResource -ResourceId $arcResourceId -ApiVersion $HCIArcAPIVersion -ErrorAction Ignore

    # Set aggregateState on HCI RP ArcSettings to DisableInProgress
    if(($arcres -ne $null) -and ($arcres.Properties.aggregateState -ne $ArcSettingsDisableInProgressState))
    {
        $disableProperties = @{"aggregateState" = $ArcSettingsDisableInProgressState}
        $arcres = New-AzResource -ResourceId $arcResourceId -ApiVersion $HCIArcAPIVersion -PropertyObject $disableProperties -Force
    }

    if($arcres -ne $null)
    {
        $extensions = Get-AzResource -ResourceId $arcResourceExtensions -ApiVersion $HCIArcAPIVersion
    }

    $extensionsCleanupSucceeded = $true

    if($extensions -ne $null)
    {
        # Remove extensions one by one. If -Force is passed write warning and proceed, else write error and stop
        for($extIndex = 0; $extIndex -lt $extensions.Count; $extIndex++)
        {
            $extension = $extensions[$extIndex]

            try
            {
                $DeletingExtensionMessageProgress = $DeletingExtensionMessage -f $extension.Name, $clusterName
                $ProgressRange = 27 # Difference between previous and next progress number
                $PercentComplete = [Math]::Round( 2 + ((($extIndex+1)/($extensions.Count)) * $ProgressRange) )
                Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $UnregisterArcProgressActivityName -Status $DeletingExtensionMessageProgress -PercentComplete $PercentComplete
                Remove-AzResource -ResourceId $extension.ResourceId -ApiVersion $HCIArcAPIVersion -Force -ErrorAction Stop | Out-Null
            }
            catch
            {
                $extensionsCleanupSucceeded = $false
                $positionMessage = $_.InvocationInfo.PositionMessage
                Write-Debug ("Exception occurred in removing extension " + $extension.Name + ". ErrorMessage: " + $_.Exception.Message + " PositionalMessage: " + $positionMessage)

                if($Force -eq $true)
                {
                    $ArcExtensionCleanupFailedWarningMsg = $ArcExtensionCleanupFailedWarning -f $extension.Name
                    Write-Warning $ArcExtensionCleanupFailedWarningMsg
                }
                else
                {
                    $ArcExtensionCleanupFailedErrorMsg = $ArcExtensionCleanupFailedError -f $extension.Name
                    Write-Error -Message $ArcExtensionCleanupFailedErrorMsg -ErrorAction Continue
                }
            }
        }
    }

    if(($Force -eq $false) -and ($extensionsCleanupSucceeded -eq $false))
    {
        return $false
    }

    # Clean up clustered scheduled task, if it exists.
    try
    {
        # Connect to cluster and use that session for registering clustered scheduled task
        if($Credential -eq $Null)
        {
            $clusterNameSession = New-PSSession -ComputerName ($clusterDNSName + "." + $ClusterDNSSuffix)
        }
        else
        {
            $clusterNameSession = New-PSSession -ComputerName ($clusterDNSName + "." + $ClusterDNSSuffix) -Credential $Credential
        }

        Invoke-Command -Session $clusterNameSession -ScriptBlock {
            $task =  Get-ScheduledTask -TaskName $using:ArcRegistrationTaskName -ErrorAction SilentlyContinue
            if ($task)
            {
                Unregister-ClusteredScheduledTask -Cluster $Using:clusterName -TaskName $using:ArcRegistrationTaskName
            }
        } | Out-Null
    }
    catch
    {
        $positionMessage = $_.InvocationInfo.PositionMessage
        Write-Error ("Exception occurred in unregistering cluster scheduled task. ErrorMessage: " + $_.Exception.Message + " PositionalMessage: " + $positionMessage) -Category OperationStopped
        throw
    }
    finally
    {
        if($clusterNameSession -ne $null)
        {
            Remove-PSSession $clusterNameSession -ErrorAction Ignore | Out-Null
        }
    }

    # Unregister all nodes.
    Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $UnregisterArcProgressActivityName -Status $WaitingUnregisterMessage -PercentComplete 30
    $disabled = Disable-ArcForServers -Session $session -Credential $Credential -ClusterDNSSuffix $ClusterDNSSuffix

    if ($disabled)
    {
        # Call HCI RP to clean up all Arc proxy resources
        $arcResource = Get-AzResource -ResourceId $arcResourceId -ErrorAction Ignore

        if($arcResource -ne $Null)
        {
            $DeletingArcCloudResourceMessageProgress = $DeletingArcCloudResourceMessage -f $arcResourceId
            Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $UnregisterArcProgressActivityName -Status $DeletingArcCloudResourceMessageProgress -PercentComplete 40
            Remove-AzResource -ResourceId $arcResourceId -Force | Out-Null
        }

        # Clean up Arc registry.
        $appId = $arcStatus.ApplicationId
        if (-Not [string]::IsNullOrEmpty($appId))
        {
            $app = Retry-Command -ScriptBlock { Get-AzureADApplication -Filter "AppId eq '$appId'" } -RetryIfNullOutput $false

            if($app -ne $Null)
            {
                $DeletingAADApplicationMessageProgress = $DeletingAADApplicationMessage -f $app.DisplayName
                Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $UnregisterArcProgressActivityName -Status $DeletingAADApplicationMessageProgress -PercentComplete 50
                Remove-ArcApplication -AppId $appId
            }
        }

        if ($arcStatus.ClusterArcStatus -ne [ArcStatus]::Disabled)
        {
            Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -Activity $UnregisterArcProgressActivityName -Status $CleanArcMessage -PercentComplete 80
            Invoke-Command -Session $session -ScriptBlock { Clear-AzureStackHCIArcIntegration }

            Write-Verbose "Successfully unregistered cluster from Arc for Servers"
        }
    }

    Write-Progress -Id $ArcProgressBarId -ParentId $MainProgressBarId -activity $UnregisterArcProgressActivityName -Completed
    return $disabled
}

enum OperationStatus
{
    Unused;
    Failed;
    Success;
    PendingForAdminConsent;
    Cancelled;
    RegisterSucceededButArcFailed
}

enum ConnectionTestResult
{
    Unused;
    Succeeded;
    Failed
}

enum ErrorDetail
{
    Unused;
    ArcPermissionsMissing;
    ArcIntegrationFailedOnNodes;
    Success
}

<#
    .Description
    Register-AzStackHCI creates a Microsoft.AzureStackHCI cloud resource representing the on-premises cluster and registers the on-premises cluster with Azure.
 
    .PARAMETER SubscriptionId
    Specifies the Azure Subscription to create the resource. This is the only Mandatory parameter.

    .PARAMETER Region
    Specifies the Region to create the resource. Default is EastUS.

    .PARAMETER ResourceName
    Specifies the resource name of the resource created in Azure. If not specified, on-premises cluster name is used.

    .PARAMETER Tag
    Specifies the resource tags for the resource in Azure in the form of key-value pairs in a hash table. For example: @{key0="value0";key1=$null;key2="value2"}

    .PARAMETER TenantId
    Specifies the Azure TenantId.

    .PARAMETER ResourceGroupName
    Specifies the Azure Resource Group name. If not specified <LocalClusterName>-rg will be used as resource group name.

    .PARAMETER ArmAccessToken
    Specifies the ARM access token. Specifying this along with GraphAccessToken and AccountId will avoid Azure interactive logon.

    .PARAMETER GraphAccessToken
    Specifies the Graph access token. Specifying this along with ArmAccessToken and AccountId will avoid Azure interactive logon.

    .PARAMETER AccountId
    Specifies the ARM access token. Specifying this along with ArmAccessToken and GraphAccessToken will avoid Azure interactive logon.

    .PARAMETER EnvironmentName
    Specifies the Azure Environment. Default is AzureCloud. Valid values are AzureCloud, AzureChinaCloud, AzurePPE, AzureCanary, AzureUSGovernment

    .PARAMETER ComputerName
    Specifies the cluster name or one of the cluster node in on-premise cluster that is being registered to Azure.

    .PARAMETER CertificateThumbprint
    Specifies the thumbprint of the certificate available on all the nodes. User is responsible for managing the certificate.

    .PARAMETER RepairRegistration
    Repair the current Azure Stack HCI registration with the cloud. This cmdlet deletes the local certificates on the clustered nodes and the remote certificates in the Azure AD application in the cloud and generates new replacement certificates for both. The resource group, resource name, and other registration choices are preserved.

    .PARAMETER UseDeviceAuthentication
    Use device code authentication instead of an interactive browser prompt.
    
    .PARAMETER EnableAzureArcServer
    Specifying this parameter to $false will skip registering the cluster nodes with Arc for servers.

    .PARAMETER Credential
    Specifies the credential for the ComputerName. Default is the current user executing the Cmdlet.

    .PARAMETER IsWAC
    Registrations through Windows Admin Center specifies this parameter to true.

    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject
    Result: Success or Failed or PendingForAdminConsent or Cancelled.
    ResourceId: Resource ID of the resource created in Azure.
    PortalResourceURL: Azure Portal Resource URL.
    PortalAADAppPermissionsURL: Azure Portal URL for AAD App permissions page.

    .EXAMPLE
    Invoking on one of the cluster node.
    C:\PS>Register-AzStackHCI -SubscriptionId "12a0f531-56cb-4340-9501-257726d741fd"
    Result: Success
    ResourceId: /subscriptions/12a0f531-56cb-4340-9501-257726d741fd/resourceGroups/DemoHCICluster1-rg/providers/Microsoft.AzureStackHCI/clusters/DemoHCICluster1
    PortalResourceURL: https://portal.azure.com/#@c31c0dbb-ce27-4c78-ad26-a5f717c14557/resource/subscriptions/12a0f531-56cb-4340-9501-257726d741fd/resourceGroups/DemoHCICluster1-rg/providers/Microsoft.AzureStackHCI/clusters/DemoHCICluster1/overview
    PortalAADAppPermissionsURL: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/CallAnAPI/appId/980be58d-578c-4cff-b4cd-43e9c3a77826/isMSAApp/

    .EXAMPLE
    Invoking from the management node
    C:\PS>Register-AzStackHCI -SubscriptionId "12a0f531-56cb-4340-9501-257726d741fd" -ComputerName ClusterNode1
    Result: Success
    ResourceId: /subscriptions/12a0f531-56cb-4340-9501-257726d741fd/resourceGroups/DemoHCICluster2-rg/providers/Microsoft.AzureStackHCI/clusters/DemoHCICluster2
    PortalResourceURL: https://portal.azure.com/#@c31c0dbb-ce27-4c78-ad26-a5f717c14557/resource/subscriptions/12a0f531-56cb-4340-9501-257726d741fd/resourceGroups/DemoHCICluster2-rg/providers/Microsoft.AzureStackHCI/clusters/DemoHCICluster2/overview
    PortalAADAppPermissionsURL: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/CallAnAPI/appId/950be58d-578c-4cff-b4cd-43e9c3a77866/isMSAApp/

    .EXAMPLE
    Invoking from WAC
    C:\PS>Register-AzStackHCI -SubscriptionId "12a0f531-56cb-4340-9501-257726d741fd" -ArmAccessToken etyer..ere= -GraphAccessToken acyee..rerrer -AccountId user1@corp1.com -Region westus -ResourceName DemoHCICluster3 -ResourceGroupName DemoHCIRG
    Result: PendingForAdminConsent
    ResourceId: /subscriptions/12a0f531-56cb-4340-9501-257726d741fd/resourceGroups/DemoHCIRG/providers/Microsoft.AzureStackHCI/clusters/DemoHCICluster3
    PortalResourceURL: https://portal.azure.com/#@c31c0dbb-ce27-4c78-ad26-a5f717c14557/resource/subscriptions/12a0f531-56cb-4340-9501-257726d741fd/resourceGroups/DemoHCIRG/providers/Microsoft.AzureStackHCI/clusters/DemoHCICluster3/overview
    PortalAADAppPermissionsURL: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/CallAnAPI/appId/980be58d-578c-4cff-b4cd-43e9c3a77866/isMSAApp/

    .EXAMPLE
    Invoking with all the parameters
    C:\PS>Register-AzStackHCI -SubscriptionId "12a0f531-56cb-4340-9501-257726d741fd" -Region westus -ResourceName HciCluster1 -TenantId "c31c0dbb-ce27-4c78-ad26-a5f717c14557" -ResourceGroupName HciClusterRG -ArmAccessToken eerrer..ere= -GraphAccessToken acee..rerrer -AccountId user1@corp1.com -EnvironmentName AzureCloud -ComputerName node1hci -Credential Get-Credential
    Result: Success
    ResourceId: /subscriptions/12a0f531-56cb-4340-9501-257726d741fd/resourceGroups/HciClusterRG/providers/Microsoft.AzureStackHCI/clusters/HciCluster1
    PortalResourceURL: https://portal.azure.com/#@c31c0dbb-ce27-4c78-ad26-a5f717c14557/resource/subscriptions/12a0f531-56cb-4340-9501-257726d741fd/resourceGroups/HciClusterRG/providers/Microsoft.AzureStackHCI/clusters/HciCluster1/overview
    PortalAADAppPermissionsURL: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/CallAnAPI/appId/990be58d-578c-4cff-b4cd-43e9c3a77866/isMSAApp/
#>
function Register-AzStackHCI{
param(
    [Parameter(Mandatory = $true)]
    [string] $SubscriptionId,

    [Parameter(Mandatory = $false)]
    [string] $Region,

    [Parameter(Mandatory = $false)]
    [string] $ResourceName,

    [Parameter(Mandatory = $false)]
    [System.Collections.Hashtable] $Tag,

    [Parameter(Mandatory = $false)]
    [string] $TenantId,

    [Parameter(Mandatory = $false)]
    [string] $ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string] $ArmAccessToken,

    [Parameter(Mandatory = $false)]
    [string] $GraphAccessToken,

    [Parameter(Mandatory = $false)]
    [string] $AccountId,

    [Parameter(Mandatory = $false)]
    [string] $EnvironmentName = $AzureCloud,

    [Parameter(Mandatory = $false)]
    [string] $ComputerName,

    [Parameter(Mandatory = $false)]
    [string] $CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [Switch]$RepairRegistration,

    [Parameter(Mandatory = $false)]
    [Switch]$UseDeviceAuthentication,
    
    [Parameter(Mandatory = $false)]
    [Switch]$EnableAzureArcServer = $true,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential] $Credential,

    [Parameter(Mandatory = $false)]
    [Switch]$IsWAC
    )

    try
    {
        Setup-Logging -LogFilePrefix "RegisterHCI"

        $registrationOutput = New-Object -TypeName PSObject
        $operationStatus = [OperationStatus]::Unused

        try
        {
            Import-PackageProvider -Name Nuget -MinimumVersion "2.8.5.201" -ErrorAction Stop
        }
        catch
        {
            Install-PackageProvider NuGet -Force | Out-Null
        }

        Show-LatestModuleVersion

        if([string]::IsNullOrEmpty($ComputerName))
        {
            $ComputerName = [Environment]::MachineName
            $IsManagementNode = $False
        }
        else
        {
            $IsManagementNode = $True
        }

        Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $FetchingRegistrationState -percentcomplete 1

        if($IsManagementNode)
        {
            if($Credential -eq $Null)
            {
                $clusterNodeSession = New-PSSession -ComputerName $ComputerName
            }
            else
            {
                $clusterNodeSession = New-PSSession -ComputerName $ComputerName -Credential $Credential
            }
        }
        else
        {
            $clusterNodeSession = New-PSSession -ComputerName localhost
        }

        $RegContext = Invoke-Command -Session $clusterNodeSession -ScriptBlock { Get-AzureStackHCI }

        if($RepairRegistration -eq $true)
        {
            if(-Not ([string]::IsNullOrEmpty($RegContext.AzureResourceUri)))
            {
                if([string]::IsNullOrEmpty($ResourceName))
                {
                    $ResourceName = $RegContext.AzureResourceUri.Split('/')[8]
                }

                if([string]::IsNullOrEmpty($ResourceGroupName))
                {
                    $ResourceGroupName = $RegContext.AzureResourceUri.Split('/')[4]
                }
            }
            else
            {
                Write-Error -Message $NoExistingRegistrationExistsErrorMessage
                $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                Write-Output $registrationOutput
                return
            }
        }

        Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $InstallRSATClusteringMessage -percentcomplete 4

        $clusScript = {
                $clusterPowershell = Get-WindowsFeature -Name RSAT-Clustering-PowerShell;
                if ( $clusterPowershell.Installed -eq $false)
                {
                    Install-WindowsFeature RSAT-Clustering-PowerShell | Out-Null;
                }
        }

        Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $ValidatingParametersFetchClusterName -percentcomplete 8;
        Invoke-Command -Session $clusterNodeSession -ScriptBlock $clusScript
        $getCluster = Invoke-Command -Session $clusterNodeSession -ScriptBlock { Get-Cluster }
        $clusterNodes = Invoke-Command -Session $clusterNodeSession -ScriptBlock { Get-ClusterNode }
        $clusterDNSSuffix = Get-ClusterDNSSuffix -Session $clusterNodeSession
        $clusterDNSName = Get-ClusterDNSName -Session $clusterNodeSession

        if([string]::IsNullOrEmpty($ResourceName))
        {
            if($getCluster -eq $Null)
            {
                $NoClusterErrorMessage = $NoClusterError -f $ComputerName
                Write-Error -Message $NoClusterErrorMessage
                $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                Write-Output $registrationOutput
                return
            }
            else
            {
                $ResourceName = $getCluster.Name
            }
        }

        if([string]::IsNullOrEmpty($ResourceGroupName))
        {
            $ResourceGroupName = $ResourceName + "-rg"
        }

        Write-Verbose "Register-AzStackHCI triggered - Region: $Region ResourceName: $ResourceName `
            SubscriptionId: $SubscriptionId Tenant: $TenantId ResourceGroupName: $ResourceGroupName `
            AccountId: $AccountId EnvironmentName: $EnvironmentName CertificateThumbprint: $CertificateThumbprint `
            RepairRegistration: $RepairRegistration EnableAzureArcServer: $EnableAzureArcServer IsWAC: $IsWAC"

        if(($EnvironmentName -eq $AzureChinaCloud) -and ($EnableAzureArcServer -eq $true))
        {
            $ArcNotAvailableMessage = $ArcIntegrationNotAvailableForCloudError -f $EnvironmentName
            Write-Error -Message $ArcNotAvailableMessage 
            $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
            Write-Output $registrationOutput
            return
        }

        if(-Not ([string]::IsNullOrEmpty($Region)))
        {
            $Region = Normalize-RegionName -Region $Region
        }

        $TenantId = Azure-Login -SubscriptionId $SubscriptionId -TenantId $TenantId -ArmAccessToken $ArmAccessToken -GraphAccessToken $GraphAccessToken -AccountId $AccountId -EnvironmentName $EnvironmentName -ProgressActivityName $RegisterProgressActivityName -UseDeviceAuthentication $UseDeviceAuthentication -Region $Region

        $resourceId = Get-ResourceId -ResourceName $ResourceName -SubscriptionId $SubscriptionId -ResourceGroupName $ResourceGroupName
        Write-Verbose "ResourceId : $resourceId"
        $resource = Get-AzResource -ResourceId $resourceId -ErrorAction Ignore
        $resGroup = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction Ignore

        if($resource -ne $null)
        {
            $resourceLocation = Normalize-RegionName -Region $resource.Location

            if([string]::IsNullOrEmpty($Region))
            {
                $Region = $resourceLocation
            }
            elseif($Region -ne $resourceLocation)
            {
                $ResourceExistsInDifferentRegionErrorMessage = $ResourceExistsInDifferentRegionError -f $resourceLocation, $Region
                Write-Error -Message $ResourceExistsInDifferentRegionErrorMessage
                $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                Write-Output $registrationOutput
                return
            }
        }
        else
        {
            if($resGroup -eq $Null)
            {
                if([string]::IsNullOrEmpty($Region))
                {
                    $Region = Get-DefaultRegion -EnvironmentName $EnvironmentName
                }
            }
            else
            {
                if([string]::IsNullOrEmpty($Region))
                {
                    $Region = $resGroup.Location
                }
            }
        }

        # Normalize region name
        $Region = Normalize-RegionName -Region $Region

        $portalResourceUrl = Get-PortalHCIResourcePageUrl -TenantId $TenantId -EnvironmentName $EnvironmentName -SubscriptionId $SubscriptionId -ResourceGroupName $ResourceGroupName -ResourceName $ResourceName -Region $Region

        if(($RegContext.RegistrationStatus -eq [RegistrationStatus]::Registered) -and ($RepairRegistration -eq $false))
        {
            if(($RegContext.AzureResourceUri -eq $resourceId))
            {
                if($resource -ne $Null)
                {
                    # Already registered with same resource Id and resource exists
                    $appId = $resource.Properties.aadClientId
                    $appPermissionsPageUrl = Get-PortalAppPermissionsPageUrl -AppId $appId -TenantId $TenantId -EnvironmentName $EnvironmentName -Region $Region
                    $operationStatus = [OperationStatus]::Success
                }
                else
                {
                    # Already registered with same resource Id and resource does not exists
                    $AlreadyRegisteredErrorMessage = $CloudResourceDoesNotExist -f $resourceId
                    Write-Error -Message $AlreadyRegisteredErrorMessage
                    $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                    Write-Output $registrationOutput
                    return
                }
            }
            else # Already registered with different resource Id
            {
                $AlreadyRegisteredErrorMessage = $RegisteredWithDifferentResourceId -f $RegContext.AzureResourceUri
                Write-Error -Message $AlreadyRegisteredErrorMessage
                $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                Write-Output $registrationOutput
                return
            }
        }
        else
        {
            Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $RegisterAzureStackRPMessage -percentcomplete 40

            $regRP = Register-AzResourceProvider -ProviderNamespace Microsoft.AzureStackHCI

            # Validate that the input region is supported by the Stack HCI RP
            $supportedRegions = [string]::Empty
            $regionSupported = Validate-RegionName -Region $Region -SupportedRegions ([ref]$supportedRegions)

            if ($regionSupported -eq $False)
            {
                $RegionNotSupportedMessage = $RegionNotSupported -f $Region, $supportedRegions
                Write-Error -Message $RegionNotSupportedMessage
                $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                Write-Output $registrationOutput
                return
            }

            # Lookup cloud endpoint URL from region name

            if($Region -eq $Region_EASTUSEUAP)
            {
                $ServiceEndpointAzureCloud = $ServiceEndpointsAzureCloud[$Region]
            }
            else
            {
                $ServiceEndpointAzureCloud = $ServiceEndpointAzureCloudFrontDoor
            }

            if($resource -eq $Null)
            {
                # Create new application

                $CreatingAADAppMessageProgress = $CreatingAADAppMessage -f $ResourceName, $TenantId
                Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $CreatingAADAppMessageProgress -percentcomplete 50

                $appId = Create-Application -AppName $ResourceName

                $appPermissionsPageUrl = Get-PortalAppPermissionsPageUrl -AppId $appId -TenantId $TenantId -EnvironmentName $EnvironmentName -Region $Region

                Write-Verbose "Created AAD Application with Id $appId"

                # Create new resource by calling RP

                if($resGroup -eq $Null)
                {
                     $CreatingResourceGroupMessageProgress = $CreatingResourceGroupMessage -f $ResourceGroupName
                     Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $CreatingResourceGroupMessageProgress -percentcomplete 55
                     $resGroup = New-AzResourceGroup -Name $ResourceGroupName -Location $Region -Tag @{$ResourceGroupCreatedByName = $ResourceGroupCreatedByValue}
                }


                $CreatingCloudResourceMessageProgress = $CreatingCloudResourceMessage -f $ResourceName
                Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $CreatingCloudResourceMessageProgress -percentcomplete 60
                $properties = @{"aadClientId"="$appId";"aadTenantId"="$TenantId"}
                $resource = New-AzResource -ResourceId $resourceId -Location $Region -ApiVersion $RPAPIVersion -PropertyObject $properties -Tag $Tag -Force

                # Try Granting admin consent for requested permissions

                $GrantingAdminConsentMessageProgress = $GrantingAdminConsentMessage -f $ResourceName
                Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $GrantingAdminConsentMessageProgress -percentcomplete 65
                $adminConsented = Grant-AdminConsent -AppId $appId

                if($adminConsented -eq $False)
                {
                    $AdminConsentWarningMsg = $AdminConsentWarning -f $ResourceName, $appPermissionsPageUrl
                    Write-Warning $AdminConsentWarningMsg
                    $operationStatus = [OperationStatus]::PendingForAdminConsent
                }
            }
            else
            {
                # Resource and Application exists. Check admin consent status

                $appId = $resource.Properties.aadClientId

                $appPermissionsPageUrl = Get-PortalAppPermissionsPageUrl -AppId $appId -TenantId $TenantId -EnvironmentName $EnvironmentName -Region $Region

                # Existing AAD app might not have the newly added scopes, if so add them.
                AddRequiredPermissionsIfNotPresent -AppId $appId
                $rolesPresent = Check-UsageAppRoles -AppId $appId
        
                $GrantingAdminConsentMessageProgress = $GrantingAdminConsentMessage -f $ResourceName
                Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $GrantingAdminConsentMessageProgress -percentcomplete 65

                if($rolesPresent -eq $False)
                {
                    # Try Granting admin consent for requested permissions

                    $adminConsented = Grant-AdminConsent -AppId $appId

                    if($adminConsented -eq $False)
                    {
                        $AdminConsentWarningMsg = $AdminConsentWarning -f $ResourceName, $appPermissionsPageUrl
                        Write-Warning $AdminConsentWarningMsg
                        $operationStatus = [OperationStatus]::PendingForAdminConsent
                    }
                }
            }

            if($operationStatus -ne [OperationStatus]::PendingForAdminConsent)
            {
                # At this point Application should be created and consented by admin.

                $appId = $resource.Properties.aadClientId
                $cloudId = $resource.Properties.cloudId 
                $app = Retry-Command -ScriptBlock { Get-AzureADApplication -Filter "AppId eq '$appId'"}
                $objectId = $app.ObjectId
                $appSP = Retry-Command -ScriptBlock { Get-AzureADServicePrincipal -Filter "AppId eq '$appId'"}
                $spObjectId = $appSP.ObjectId

                # Add certificate

                Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $GettingCertificateMessage -percentcomplete 70

                $CertificatesToBeMaintained = @{}
                $NewCertificateFailedNodes = [System.Collections.ArrayList]::new()
                $SetCertificateFailedNodes = [System.Collections.ArrayList]::new()
                $OSNotLatestOnNodes = [System.Collections.ArrayList]::new()

                $ServiceEndpoint = ""
                $Authority = ""
                $BillingServiceApiScope = ""
                $GraphServiceApiScope = ""

                Get-EnvironmentEndpoints -EnvironmentName $EnvironmentName -ServiceEndpoint ([ref]$ServiceEndpoint) -Authority ([ref]$Authority) -BillingServiceApiScope ([ref]$BillingServiceApiScope) -GraphServiceApiScope ([ref]$GraphServiceApiScope)

                $setupCertsError = Setup-Certificates -ClusterNodes $clusterNodes -Credential $Credential -ResourceName $ResourceName -ObjectId $objectId -CertificateThumbprint $CertificateThumbprint -AppId $appId -TenantId $TenantId -CloudId $cloudId `
                                    -ServiceEndpoint $ServiceEndpoint -BillingServiceApiScope $BillingServiceApiScope -GraphServiceApiScope $GraphServiceApiScope -Authority $Authority -NewCertificateFailedNodes $NewCertificateFailedNodes `
                                    -SetCertificateFailedNodes $SetCertificateFailedNodes -OSNotLatestOnNodes $OSNotLatestOnNodes -CertificatesToBeMaintained $CertificatesToBeMaintained -ClusterDNSSuffix $clusterDNSSuffix

                if($setupCertsError -ne $null)
                {
                    Write-Error -Message $setupCertsError
                    $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                    Write-Output $registrationOutput
                    return
                }

                if(($SetCertificateFailedNodes.Count -ge 1) -or ($NewCertificateFailedNodes.Count -ge 1))
                {
                    # Failed on atleast 1 node
                    $SettingCertificateFailedMessage = $SettingCertificateFailed -f ($NewCertificateFailedNodes -join ","),($SetCertificateFailedNodes -join ",")
                    Write-Error -Message $SettingCertificateFailedMessage
                    $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                    Write-Output $registrationOutput
                    return
                }

                if($OSNotLatestOnNodes.Count -ge 1)
                {
                    $NotAllTheNodesInClusterAreGAError = $NotAllTheNodesInClusterAreGA -f ($OSNotLatestOnNodes -join ",")
                    Write-Error -Message $NotAllTheNodesInClusterAreGAError
                    $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                    Write-Output $registrationOutput
                    return
                }

                Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $RegisterAndSyncMetadataMessage -percentcomplete 90

                # Register by calling on-prem usage service Cmdlet

                $RegistrationParams = @{
                                            ServiceEndpoint = $ServiceEndpoint
                                            BillingServiceApiScope = $BillingServiceApiScope
                                            GraphServiceApiScope = $GraphServiceApiScope
                                            AADAuthority = $Authority
                                            AppId = $appId
                                            TenantId = $TenantId
                                            CloudId = $cloudId
                                            SubscriptionId = $SubscriptionId
                                            ObjectId = $objectId
                                            ResourceName = $ResourceName
                                            ProviderNamespace = "Microsoft.AzureStackHCI"
                                            ResourceArmId = $resourceId
                                            ServicePrincipalClientId = $spObjectId
                                            CertificateThumbprint = $CertificateThumbprint
                                        }

                Invoke-Command -Session $clusterNodeSession -ScriptBlock { Set-AzureStackHCIRegistration @Using:RegistrationParams }

                # Delete all certificates except certificates which we created in this current registration flow.
                if(($RepairRegistration -eq $true) -or (-Not ([string]::IsNullOrEmpty($CertificateThumbprint))) )
                {
                    $aadAppKeyCreds = Retry-Command -ScriptBlock {Get-AzureADApplicationKeyCredential -ObjectId $objectId}
                    Foreach ($keyCred in $aadAppKeyCreds)
                    {
                        if($CertificatesToBeMaintained[$keyCred.KeyId] -eq $true)
                        {
                            Write-Verbose ($SkippingDeleteCertificateFromAADApp -f $keyCred.KeyId)
                            continue
                        }
                        else
                        {
                            Write-Verbose ($DeletingCertificateFromAADApp -f $keyCred.KeyId)
                            Retry-Command -ScriptBlock { Remove-AzureADApplicationKeyCredential -ObjectId $objectId -KeyId $keyCred.KeyId} -RetryIfNullOutput $false
                        }
                    }
                }

                # Delete old unused scopes if present
                Remove-OldScopes -AppId $appId

                $operationStatus = [OperationStatus]::Success
            }
        }

        if (($operationStatus -ne [OperationStatus]::PendingForAdminConsent) -and ($EnableAzureArcServer -eq $true))
        {
            Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -status $RegisterArcMessage -percentcomplete 91

            $ArcCmdletsAbsentOnNodes = [System.Collections.ArrayList]::new()

            Foreach ($clusNode in $clusterNodes)
            {
                $nodeSession = $null

                try
                {
                    if($Credential -eq $Null)
                    {
                        $nodeSession = New-PSSession -ComputerName ($clusNode.Name + "." + $clusterDNSSuffix)
                    }
                    else
                    {
                        $nodeSession = New-PSSession -ComputerName ($clusNode.Name + "." + $clusterDNSSuffix) -Credential $Credential
                    }
                }
                catch
                {
                    Write-Debug ("Exception occurred in establishing new PSSession. ErrorMessage : " + $_.Exception.Message)
                    Write-Debug $_
                    $ArcCmdletsAbsentOnNodes.Add($clusNode.Name) | Out-Null
                    continue
                }

                # Check if node has Arc registration Cmdlets
                $cmdlet = Invoke-Command -Session $nodeSession -ScriptBlock { Get-Command Get-AzureStackHCIArcIntegration -Type Cmdlet -ErrorAction Ignore }

                if($cmdlet -eq $null)
                {
                    $ArcCmdletsAbsentOnNodes.Add($clusNode.Name) | Out-Null
                }

                if($nodeSession -ne $null)
                {
                    Remove-PSSession $nodeSession -ErrorAction Ignore | Out-Null
                }
            }

            if($ArcCmdletsAbsentOnNodes.Count -ge 1)
            {
                # Show Arc error on 20h2 only if -EnableAzureArcServer:$true is explicity passed by user
                if($PSBoundParameters.ContainsKey('EnableAzureArcServer'))
                {
                    $ArcCmdletsNotAvailableErrorMsg = $ArcCmdletsNotAvailableError -f ($ArcCmdletsAbsentOnNodes -join ",")
                    Write-Error -Message $ArcCmdletsNotAvailableErrorMsg
                    $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                    Write-Output $registrationOutput
                    return
                }
            }
            else
            {
                $arcResourceId = $resourceId + $HCIArcInstanceName
                $arcResourceGroupName = $ResourceGroupName

                $arcres = Get-AzResource -ResourceId $arcResourceId -ApiVersion $HCIArcAPIVersion -ErrorAction Ignore

                if($arcres -eq $null)
                {
                    $arcres = New-AzResource -ResourceId $arcResourceId -ApiVersion $HCIArcAPIVersion -Force
                }
                else
                {
                    if ($arcres.Properties.aggregateState -eq $ArcSettingsDisableInProgressState)
                    {
                        Write-Error -Message $ArcRegistrationDisableInProgressError
                        $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                        Write-Output $registrationOutput
                        return
                    }
                }

                $arcResourceGroupName = $arcres.Properties.arcInstanceResourceGroup
                $arcAppName = $ResourceName + ".arc"

                Write-Verbose "Register-AzStackHCI: Arc registration triggered. ArcResourceGroupName: $arcResourceGroupName"
                $arcResult = Register-ArcForServers -IsManagementNode $IsManagementNode -ComputerName $ComputerName -Credential $Credential -TenantId $TenantId -SubscriptionId $SubscriptionId -ResourceGroup $arcResourceGroupName -Region $Region -AppName $arcAppName -ClusterDNSSuffix $clusterDNSSuffix -IsWAC:$IsWAC -Environment:$EnvironmentName

                if($arcResult -ne [ErrorDetail]::Success)
                {
                    $operationStatus = [OperationStatus]::RegisterSucceededButArcFailed
                    $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyErrorDetail -Value $arcResult
                }
            }
        }

        Write-Progress -Id $MainProgressBarId -activity $RegisterProgressActivityName -Completed

        $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value $operationStatus
        $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyPortalResourceURL -Value $portalResourceUrl
        $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResourceId -Value $resourceId
        $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyPortalAADAppPermissionsURL -Value $appPermissionsPageUrl

        if($operationStatus -eq [OperationStatus]::PendingForAdminConsent)
        {
            $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyDetails -Value $AdminConsentWarningMsg
        }
        else
        {
            $registrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyDetails -Value $RegistrationSuccessDetailsMessage
        }

        Write-Output $registrationOutput
    }
    catch
    {
        Write-Error -Exception $_.Exception -Category OperationStopped -ErrorAction Continue
        # Get script line number, offset and Command that resulted in exception. Write-Error with the exception above does not write this info.
        $positionMessage = $_.InvocationInfo.PositionMessage
        Write-Error ("Exception occurred in Register-AzStackHCI : " + $positionMessage) -Category OperationStopped
        throw
    }
    finally
    {
        try{ Disconnect-AzAccount | Out-Null } catch{}
        try{ Disconnect-AzureAD | Out-Null } catch{}
        Stop-Transcript | out-null
    }
}

<#
    .Description
    Unregister-AzStackHCI deletes the Microsoft.AzureStackHCI cloud resource representing the on-premises cluster and unregisters the on-premises cluster with Azure.
    The registered information available on the cluster is used to unregister the cluster if no parameters are passed.

    .PARAMETER SubscriptionId
    Specifies the Azure Subscription to create the resource

    .PARAMETER Region
    Specifies the Region the resource is created in Azure.

    .PARAMETER ResourceName
    Specifies the resource name of the resource created in Azure. If not specified, on-premises cluster name is used.

    .PARAMETER TenantId
    Specifies the Azure TenantId.

    .PARAMETER ResourceGroupName
    Specifies the Azure Resource Group name. If not specified <LocalClusterName>-rg will be used as resource group name.

    .PARAMETER ArmAccessToken
    Specifies the ARM access token. Specifying this along with GraphAccessToken and AccountId will avoid Azure interactive logon.

    .PARAMETER GraphAccessToken
    Specifies the Graph access token. Specifying this along with ArmAccessToken and AccountId will avoid Azure interactive logon.

    .PARAMETER AccountId
    Specifies the ARM access token. Specifying this along with ArmAccessToken and GraphAccessToken will avoid Azure interactive logon.

    .PARAMETER EnvironmentName
    Specifies the Azure Environment. Default is AzureCloud. Valid values are AzureCloud, AzureChinaCloud, AzurePPE, AzureCanary, AzureUSGovernment

    .PARAMETER UseDeviceAuthentication
    Use device code authentication instead of an interactive browser prompt.

    .PARAMETER ComputerName
    Specifies one of the cluster node in on-premise cluster that is being registered to Azure.

    .PARAMETER DisableOnlyAzureArcServer
    Specifying this parameter to $true will only unregister the cluster nodes with Arc for servers and Azure Stack HCI registration will not be altered.

    .PARAMETER Credential
    Specifies the credential for the ComputerName. Default is the current user executing the Cmdlet.

    .PARAMETER Force
    Specifies that unregistration should continue even if we could not delete the Arc extensions on the nodes.

    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject
    Result: Success or Failed or Cancelled.

    .EXAMPLE
    Invoking on one of the cluster node
    C:\PS>Unregister-AzStackHCI
    Result: Success

    .EXAMPLE
    Invoking from the management node
    C:\PS>Unregister-AzStackHCI -ComputerName ClusterNode1
    Result: Success

    .EXAMPLE
    Invoking from WAC
    C:\PS>Unregister-AzStackHCI -SubscriptionId "12a0f531-56cb-4340-9501-257726d741fd" -ArmAccessToken etyer..ere= -GraphAccessToken acyee..rerrer -AccountId user1@corp1.com -ResourceName DemoHCICluster3 -ResourceGroupName DemoHCIRG -Confirm:$False
    Result: Success

    .EXAMPLE
    Invoking with all the parameters
    C:\PS>Unregister-AzStackHCI -SubscriptionId "12a0f531-56cb-4340-9501-257726d741fd" -ResourceName HciCluster1 -TenantId "c31c0dbb-ce27-4c78-ad26-a5f717c14557" -ResourceGroupName HciClusterRG -ArmAccessToken eerrer..ere= -GraphAccessToken acee..rerrer -AccountId user1@corp1.com -EnvironmentName AzureCloud -ComputerName node1hci -Credential Get-Credential
    Result: Success
#>
function Unregister-AzStackHCI{
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [string] $SubscriptionId,

    [Parameter(Mandatory = $false)]
    [string] $ResourceName,

    [Parameter(Mandatory = $false)]
    [string] $TenantId,

    [Parameter(Mandatory = $false)]
    [string] $ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string] $ArmAccessToken,

    [Parameter(Mandatory = $false)]
    [string] $GraphAccessToken,

    [Parameter(Mandatory = $false)]
    [string] $AccountId,

    [Parameter(Mandatory = $false)]
    [string] $EnvironmentName = $AzureCloud,

    [Parameter(Mandatory = $false)]
    [string] $Region,

    [Parameter(Mandatory = $false)]
    [string] $ComputerName,

    [Parameter(Mandatory = $false)]
    [Switch]$UseDeviceAuthentication,

    [Parameter(Mandatory = $false)]
    [Switch]$DisableOnlyAzureArcServer = $false,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential] $Credential,

    [Parameter(Mandatory = $false)]
    [Switch] $Force
    )

    try
    {
        Setup-Logging -LogFilePrefix "UnregisterHCI"

        $unregistrationOutput = New-Object -TypeName PSObject
        $operationStatus = [OperationStatus]::Unused

        if([string]::IsNullOrEmpty($ComputerName))
        {
            $ComputerName = [Environment]::MachineName
            $IsManagementNode = $False
        }
        else
        {
            $IsManagementNode = $True
        }

        Write-Progress -Id $MainProgressBarId -activity $UnregisterProgressActivityName -status $FetchingRegistrationState -percentcomplete 1

        if($IsManagementNode)
        {
            if($Credential -eq $Null)
            {
                $clusterNodeSession = New-PSSession -ComputerName $ComputerName
            }
            else
            {
                $clusterNodeSession = New-PSSession -ComputerName $ComputerName -Credential $Credential
            }

            $RegContext = Invoke-Command -Session $clusterNodeSession -ScriptBlock { Get-AzureStackHCI }
        }
        else
        {
            $RegContext = Get-AzureStackHCI
            $clusterNodeSession = New-PSSession -ComputerName localhost
        }

        $clusScript = {
                $clusterPowershell = Get-WindowsFeature -Name RSAT-Clustering-PowerShell;
                if ( $clusterPowershell.Installed -eq $false)
                {
                    Install-WindowsFeature RSAT-Clustering-PowerShell | Out-Null;
                }
            }

        Invoke-Command -Session $clusterNodeSession -ScriptBlock $clusScript
        $clusterDNSSuffix = Get-ClusterDNSSuffix -Session $clusterNodeSession
        $clusterDNSName = Get-ClusterDNSName -Session $clusterNodeSession

        Write-Progress -Id $MainProgressBarId -activity $UnregisterProgressActivityName -status $ValidatingParametersRegisteredInfo -percentcomplete 5

        if([string]::IsNullOrEmpty($ResourceName) -or [string]::IsNullOrEmpty($SubscriptionId))
        {
            if($RegContext.RegistrationStatus -ne [RegistrationStatus]::Registered)
            {
                Write-Error -Message $RegistrationInfoNotFound
                $unregistrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                Write-Output $unregistrationOutput
                return
            }
        }

        if([string]::IsNullOrEmpty($SubscriptionId))
        {
            $SubscriptionId = $RegContext.AzureResourceUri.Split('/')[2]
        }

        if([string]::IsNullOrEmpty($ResourceGroupName))
        {
            $ResourceGroupName = If ($RegContext.RegistrationStatus -ne [RegistrationStatus]::Registered) { $ResourceName + "-rg" } Else { $RegContext.AzureResourceUri.Split('/')[4] }
        }

        if([string]::IsNullOrEmpty($ResourceName))
        {
            $ResourceName = $RegContext.AzureResourceUri.Split('/')[8]
        }

        $resourceId = Get-ResourceId -ResourceName $ResourceName -SubscriptionId $SubscriptionId -ResourceGroupName $ResourceGroupName

        if ($PSCmdlet.ShouldProcess($resourceId))
        {
            Write-Verbose "Unregister-AzStackHCI triggered - ResourceName: $ResourceName Region: $Region `
                   SubscriptionId: $SubscriptionId Tenant: $TenantId ResourceGroupName: $ResourceGroupName `
                   AccountId: $AccountId EnvironmentName: $EnvironmentName DisableOnlyAzureArcServer: $DisableOnlyAzureArcServer Force:$Force"

            if(-Not ([string]::IsNullOrEmpty($Region)))
            {
                $Region = Normalize-RegionName -Region $Region
            }

            $TenantId = Azure-Login -SubscriptionId $SubscriptionId -TenantId $TenantId -ArmAccessToken $ArmAccessToken -GraphAccessToken $GraphAccessToken -AccountId $AccountId -EnvironmentName $EnvironmentName -ProgressActivityName $UnregisterProgressActivityName -UseDeviceAuthentication $UseDeviceAuthentication -Region $Region

            Write-Progress -Id $MainProgressBarId -activity $UnregisterProgressActivityName -status $UnregisterArcMessage -percentcomplete 40

            $arcUnregisterRes = Unregister-ArcForServers -IsManagementNode $IsManagementNode -ComputerName $ComputerName -Credential $Credential -ResourceId $resourceId -Force:$Force -ClusterDNSSuffix $clusterDNSSuffix

            if($arcUnregisterRes -eq $false)
            {
                $unregistrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Failed
                Write-Output $unregistrationOutput
                return
            }
            else
            {
                if ($DisableOnlyAzureArcServer -eq $true)
                {
                    $unregistrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value [OperationStatus]::Success
                    Write-Output $unregistrationOutput
                    return
                }
            }

            Write-Progress -Id $MainProgressBarId -activity $UnregisterProgressActivityName -status $UnregisterHCIUsageMessage -percentcomplete 45
        
            if($RegContext.RegistrationStatus -eq [RegistrationStatus]::Registered)
            {

                Invoke-Command -Session $clusterNodeSession -ScriptBlock { Remove-AzureStackHCIRegistration }
                $clusterNodes = Invoke-Command -Session $clusterNodeSession -ScriptBlock { Get-ClusterNode }

                Foreach ($clusNode in $clusterNodes)
                {
                    $nodeSession = $null

                    try
                    {
                        if($Credential -eq $Null)
                        {
                            $nodeSession = New-PSSession -ComputerName ($clusNode.Name + "." + $clusterDNSSuffix)
                        }
                        else
                        {
                            $nodeSession = New-PSSession -ComputerName ($clusNode.Name + "." + $clusterDNSSuffix) -Credential $Credential
                        }

                        if([Environment]::MachineName -eq $clusNode.Name)
                        {
                            Remove-AzureStackHCIRegistrationCertificate
                        }
                        else
                        {
                            Invoke-Command -Session $nodeSession -ScriptBlock { Remove-AzureStackHCIRegistrationCertificate }
                        }
                    }
                    catch
                    {
                        Write-Warning ($FailedToRemoveRegistrationCertWarning -f $clusNode.Name)
                        Write-Debug ("Exception occurred in clearing certificate on {0}. ErrorMessage : {1}" -f ($clusNode.Name), ($_.Exception.Message))
                        Write-Debug $_
                        continue
                    }
                }
            }

            $resource = Get-AzResource -ResourceId $resourceId -ErrorAction Ignore

            if($resource -ne $Null)
            {
                $appId = $resource.Properties.aadClientId
                $app = Retry-Command -ScriptBlock { Get-AzureADApplication -Filter "AppId eq '$appId'"} -RetryIfNullOutput $false
                
                if($app -ne $Null)
                {
                    $DeletingAADApplicationMessageProgress = $DeletingAADApplicationMessage -f $ResourceName
                    Write-Progress -Id $MainProgressBarId -activity $UnregisterProgressActivityName -status $DeletingAADApplicationMessageProgress -percentcomplete 60
                    Retry-Command -ScriptBlock { Remove-AzureADApplication -ObjectId $app.ObjectId} -RetryIfNullOutput $false
                }

                $DeletingCloudResourceMessageProgress = $DeletingCloudResourceMessage -f $ResourceName
                Write-Progress -Id $MainProgressBarId -activity $UnregisterProgressActivityName -status $DeletingCloudResourceMessageProgress -percentcomplete 80

                $remResource = Remove-AzResource -ResourceId $resourceId -Force
            }

            $resGroup = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction Ignore
            if($resGroup -ne $Null)
            {
                $resGroupTags = $resGroup.Tags

                if($resGroupTags -ne $null)
                {
                    $resGroupTagsCreatedBy = $resGroupTags[$ResourceGroupCreatedByName]

                    # If resource is created by us during registration and if there are no resources in resource group, then delete it.
                    if($resGroupTagsCreatedBy -eq $ResourceGroupCreatedByValue)
                    {
                        $resourcesInRG = Get-AzResource -ResourceGroupName $ResourceGroupName

                        if($resourcesInRG -eq $null) # Resource group is empty
                        {
                            Remove-AzResourceGroup -Name $ResourceGroupName -Force | Out-Null
                        }
                    }
                }
            }

            $operationStatus = [OperationStatus]::Success
        }
        else
        {
            $operationStatus = [OperationStatus]::Cancelled
        }

        Write-Progress -Id $MainProgressBarId -activity $UnregisterProgressActivityName -Completed

        $unregistrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value $operationStatus

        if ($operationStatus -eq [OperationStatus]::Success)
        {
            $unregistrationOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyDetails -Value $UnregistrationSuccessDetailsMessage
        }

        Write-Output $unregistrationOutput
    }
    catch
    {
        Write-Error -Exception $_.Exception -Category OperationStopped -ErrorAction Continue
        # Get script line number, offset and Command that resulted in exception. Write-Error with the exception above does not write this info.
        $positionMessage = $_.InvocationInfo.PositionMessage
        Write-Error ("Exception occurred in Unregister-AzStackHCI : " + $positionMessage) -Category OperationStopped
        throw
    }
    finally
    {
        try{ Disconnect-AzAccount | Out-Null } catch{}
        try{ Disconnect-AzureAD | Out-Null } catch{}
        Stop-Transcript | out-null
    }
}

<#
    .Description
    Test-AzStackHCIConnection verifies connectivity from on-premises clustered nodes to the Azure services required by Azure Stack HCI.

    .PARAMETER EnvironmentName
    Specifies the Azure Environment. Default is AzureCloud. Valid values are AzureCloud, AzureChinaCloud, AzurePPE, AzureCanary, AzureUSGovernment

    .PARAMETER Region
    Specifies the Region to connect to. Not used unless it is Canary region.

    .PARAMETER ComputerName
    Specifies one of the cluster node in on-premise cluster that is being registered to Azure.

    .PARAMETER Credential
    Specifies the credential for the ComputerName. Default is the current user executing the Cmdlet.

    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject
    Test: Name of the test performed.
    EndpointTested: Endpoint used in the test.
    IsRequired: True or False
    Result: Succeeded or Failed
    FailedNodes: List of nodes on which the test failed.

    .EXAMPLE
    Invoking on one of the cluster node. Success case.
    C:\PS>Test-AzStackHCIConnection
    Test: Connect to Azure Stack HCI Service
    EndpointTested: https://azurestackhci-df.azurefd.net/health
    IsRequired: True
    Result: Succeeded

    .EXAMPLE
    Invoking on one of the cluster node. Failed case.
    C:\PS>Test-AzStackHCIConnection
    Test: Connect to Azure Stack HCI Service
    EndpointTested: https://azurestackhci-df.azurefd.net/health
    IsRequired: True
    Result: Failed
    FailedNodes: Node1inClus2, Node2inClus3
#>
function Test-AzStackHCIConnection{
param(
    [Parameter(Mandatory = $false)]
    [string] $EnvironmentName = $AzureCloud,

    [Parameter(Mandatory = $false)]
    [string] $Region,

    [Parameter(Mandatory = $false)]
    [string] $ComputerName,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential] $Credential
    )

    try
    {
        Setup-Logging -LogFilePrefix "TestAzStackHCIConnection"

        $testConnectionnOutput = New-Object -TypeName PSObject
        $connectionTestResult = [ConnectionTestResult]::Unused

        if([string]::IsNullOrEmpty($ComputerName))
        {
            $ComputerName = [Environment]::MachineName
            $IsManagementNode = $False
        }
        else
        {
            $IsManagementNode = $True
        }

        if($IsManagementNode)
        {
            if($Credential -eq $Null)
            {
                $clusterNodeSession = New-PSSession -ComputerName $ComputerName
            }
            else
            {
                $clusterNodeSession = New-PSSession -ComputerName $ComputerName -Credential $Credential
            }
        }
        else
        {
            $clusterNodeSession = New-PSSession -ComputerName localhost
        }

        if(-not([string]::IsNullOrEmpty($Region)))
        {
            $Region = Normalize-RegionName -Region $Region

            if($Region -eq $Region_EASTUSEUAP)
            {
                $ServiceEndpointAzureCloud = $ServiceEndpointsAzureCloud[$Region]
            }
            else
            {
                $ServiceEndpointAzureCloud = $ServiceEndpointAzureCloudFrontDoor
            }
        }

        $clusScript = {
                $clusterPowershell = Get-WindowsFeature -Name RSAT-Clustering-PowerShell;
                if ( $clusterPowershell.Installed -eq $false)
                {
                    Install-WindowsFeature RSAT-Clustering-PowerShell | Out-Null;
                }
            }

        Invoke-Command -Session $clusterNodeSession -ScriptBlock $clusScript
        $getCluster = Invoke-Command -Session $clusterNodeSession -ScriptBlock { Get-Cluster }
        $clusterDNSSuffix = Get-ClusterDNSSuffix -Session $clusterNodeSession
        $clusterDNSName = Get-ClusterDNSName -Session $clusterNodeSession

        if($getCluster -eq $Null)
        {
            $NoClusterErrorMessage = $NoClusterError -f $ComputerName
            Write-Error -Message $NoClusterErrorMessage
            return
        }
        else
        {
            $ServiceEndpoint = ""
            $Authority = ""
            $BillingServiceApiScope = ""
            $GraphServiceApiScope = ""

            Get-EnvironmentEndpoints -EnvironmentName $EnvironmentName -ServiceEndpoint ([ref]$ServiceEndpoint) -Authority ([ref]$Authority) -BillingServiceApiScope ([ref]$BillingServiceApiScope) -GraphServiceApiScope ([ref]$GraphServiceApiScope)
            $EndPointToInvoke = $ServiceEndpoint + $HealthEndpointPath

            $clusterNodes = Invoke-Command -Session $clusterNodeSession -ScriptBlock { Get-ClusterNode }
            $HealthEndPointCheckFailedNodes = [System.Collections.ArrayList]::new()

            $testConnectionnOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyTest -Value $ConnectionTestToAzureHCIServiceName
            $testConnectionnOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyEndpointTested -Value $EndPointToInvoke
            $testConnectionnOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyIsRequired -Value $True

            Check-ConnectionToCloudBillingService -ClusterNodes $clusterNodes -Credential $Credential -HealthEndpoint $EndPointToInvoke -HealthEndPointCheckFailedNodes $HealthEndPointCheckFailedNodes -ClusterDNSSuffix $clusterDNSSuffix

            if($HealthEndPointCheckFailedNodes.Count -ge 1)
            {
                # Failed on atleast 1 node
                $connectionTestResult = [ConnectionTestResult]::Failed
                $testConnectionnOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyFailedNodes -Value $HealthEndPointCheckFailedNodes
            }
            else
            {
                $connectionTestResult = [ConnectionTestResult]::Succeeded
            }

            $testConnectionnOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value $connectionTestResult
            Write-Output $testConnectionnOutput
            return
        }
    }
    catch
    {
        Write-Error -Exception $_.Exception -Category OperationStopped -ErrorAction Continue
        # Get script line number, offset and Command that resulted in exception. Write-Error with the exception above does not write this info.
        $positionMessage = $_.InvocationInfo.PositionMessage
        Write-Error ("Exception occurred in Test-AzStackHCIConnection : " + $positionMessage) -Category OperationStopped
        throw
    }
    finally
    {
        Stop-Transcript | out-null
    }
}

<#
    .Description
    Set-AzStackHCI modifies resource properties of the Microsoft.AzureStackHCI cloud resource representing the on-premises cluster to enable or disable features.
    .PARAMETER ComputerName
    Specifies one of the cluster node in on-premise cluster that is registered to Azure.
    .PARAMETER Credential
    Specifies the credential for the ComputerName. Default is the current user executing the Cmdlet.
    .PARAMETER ResourceId
    Specifies the fully qualified resource ID, including the subscription, as in the following example: `/Subscriptions/`subscription ID`/providers/Microsoft.AzureStackHCI/clusters/MyCluster`
    .PARAMETER EnableWSSubscription
    Specifies if Windows Server Subscription should be enabled or disabled. Enabling this feature starts billing through your Azure subscription for Windows Server guest licenses.
    .PARAMETER DiagnosticLevel
    Specifies the diagnostic level for the cluster.
    .PARAMETER TenantId
    Specifies the Azure TenantId.
    .PARAMETER ArmAccessToken
    Specifies the ARM access token. Specifying this along with GraphAccessToken and AccountId will avoid Azure interactive logon.
    .PARAMETER GraphAccessToken
    Specifies the Graph access token. Specifying this along with ArmAccessToken and AccountId will avoid Azure interactive logon.
    .PARAMETER AccountId
    Specifies the ARM access token. Specifying this along with ArmAccessToken and GraphAccessToken will avoid Azure interactive logon.
    .PARAMETER EnvironmentName
    Specifies the Azure Environment. Default is AzureCloud. Valid values are AzureCloud, AzureChinaCloud, AzurePPE, AzureCanary, AzureUSGovernment
    .PARAMETER UseDeviceAuthentication
    Use device code authentication instead of an interactive browser prompt.
    .PARAMETER Force
    Forces the command to run without asking for user confirmation.
    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject
    Result: Success or Failed or Cancelled.
    .EXAMPLE
    Invoking on one of the cluster node to enable Windows Server Subscription feature
    PS C:\> Set-AzStackHCI -EnableWSSubscription $true
    Result: Success
    .EXAMPLE
    Invoking from the management node to set the diagnostic level to Basic
    PS C:\> Set-AzStackHCI -ComputerName ClusterNode1 -DiagnosticLevel Basic
    Result: Success
#>
function Set-AzStackHCI{
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
[OutputType([PSCustomObject])]
param(
    [Parameter(Position = 0, Mandatory = $false)]
    [string] $ComputerName,
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential] $Credential,
    [Parameter(Mandatory = $false)]
    [string] $ResourceId,
    [Parameter(Mandatory = $false)]
    [Bool] $EnableWSSubscription,
    [Parameter(Mandatory = $false)]
    [DiagnosticLevel] $DiagnosticLevel,
    [Parameter(Mandatory = $false)]
    [string] $TenantId,

    [Parameter(Mandatory = $false)]
    [string] $ArmAccessToken,

    [Parameter(Mandatory = $false)]
    [string] $GraphAccessToken,

    [Parameter(Mandatory = $false)]
    [string] $AccountId,

    [Parameter(Mandatory = $false)]
    [string] $EnvironmentName = $AzureCloud,

    [Parameter(Mandatory = $false)]
    [Switch]$UseDeviceAuthentication,

    [Parameter(Mandatory = $false)]
    [Switch] $Force
    )

    $setOutput          = New-Object -TypeName PSObject
    $doSetResource      = $false
    $needShouldContinue = $false
    $doAzAuth           = $false
    $isManagementNode   = $false
    $nodeSessionParams  = @{}
    $subscriptionId     = [string]::Empty
    $armResourceId      = [string]::Empty
    $armResource        = $null

    $successMessage     = New-Object -TypeName System.Text.StringBuilder

    try
    {
        Setup-Logging -LogFilePrefix "SetAzStackHCI"

        Show-LatestModuleVersion

        if([string]::IsNullOrEmpty($ComputerName))
        {
            $ComputerName = [Environment]::MachineName
            $isManagementNode = $false
        }
        else
        {
            $isManagementNode = $true
        }

        Write-Progress -Id $MainProgressBarId -Activity $SetProgressActivityName -Status $SetProgressStatusGathering -PercentComplete 5

        if($PSBoundParameters.ContainsKey('ResourceId') -eq $false)
        {
            $regContext = $null

            if($isManagementNode)
            {
                $nodeSessionParams.Add('ComputerName', $ComputerName)

                if($Credential -ne $null)
                {
                    $nodeSessionParams.Add('Credential', $Credential)
                }

                $regContext = Invoke-Command @nodeSessionParams -ScriptBlock { Get-AzureStackHCI }
            }
            else
            {
                $regContext = Get-AzureStackHCI
            }

            if ($regContext.RegistrationStatus -ne [RegistrationStatus]::Registered)
            {
                Write-Error -Category InvalidOperation -Message $SetAzResourceClusterNotRegistered

                $setOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value ([OperationStatus]::Failed)
                $setOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyErrorDetail -Value $SetAzResourceClusterNotRegistered

                Write-Output $setOutput

                return
            }

            $clusScript = {
                    $clusterPowershell = Get-WindowsFeature -Name RSAT-Clustering-PowerShell;
                    if ( $clusterPowershell.Installed -eq $false)
                    {
                        Install-WindowsFeature RSAT-Clustering-PowerShell | Out-Null;
                    }
                }

            Invoke-Command @nodeSessionParams -ScriptBlock $clusScript

            $clusterNodes = Invoke-Command @nodeSessionParams -ScriptBlock { Get-ClusterNode }

            $nodeDown = $false
            $nodeDown = ($clusterNodes | % { if ($_.State -ne 'Up') { return $true } })

            if ($nodeDown -eq $true)
            {
                Write-Error -Category ConnectionError -Message $SetAzResourceClusterNodesDown

                $setOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value ([OperationStatus]::Failed)
                $setOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyErrorDetail -Value $SetAzResourceClusterNodesDown

                Write-Output $setOutput

                return
            }

            $subscriptionId    = $regContext.AzureResourceUri.Split('/')[2]
            $resourceGroupName = $regContext.AzureResourceUri.Split('/')[4]
            $resourceName      = $regContext.AzureResourceUri.Split('/')[8]

            $armResourceId = Get-ResourceId -SubscriptionId $subscriptionId -ResourceGroupName $resourceGroupName -ResourceName $resourceName
        }
        else
        {
            $armResourceId  = $ResourceId
            $subscriptionId = $ResourceId.Split('/')[2]
        }

        Write-Progress -Id $MainProgressBarId -Activity $SetProgressActivityName -Status $SetProgressStatusGetAzureResource -PercentComplete 20

        if($PSBoundParameters.ContainsKey('ArmAccessToken') -eq $true)
        {
            $doAzAuth = $true
        }
        else
        {
            $azContext = Get-AzContext -ErrorAction SilentlyContinue

            if ($azContext -eq $null)
            {
                $doAzAuth = $true
            }
            else
            {
                if ($azContext.Subscription.Id -ne $subscriptionId)
                {
                    $currentOperation = ($SetProgressStatusOpSwitching -f $subscriptionId)
                    Write-Progress -Id $MainProgressBarId -Activity $SetProgressActivityName -Status $SetProgressStatusGetAzureResource -CurrentOperation $currentOperation -PercentComplete 35

                    $azContext = Set-AzContext -SubscriptionId $subscriptionId -ErrorAction Stop
                }
            }
        }

        if ($doAzAuth -eq $true)
        {
            $azureLoginParameters = @{
                                        'SubscriptionId'          = $subscriptionId;
                                        'TenantId'                = $TenantId;
                                        'ArmAccessToken'          = $ArmAccessToken;
                                        'GraphAccessToken'        = $GraphAccessToken;
                                        'AccountId'               = $AccountId;
                                        'EnvironmentName'         = $EnvironmentName;
                                        'UseDeviceAuthentication' = $UseDeviceAuthentication;
                                        'ProgressActivityName'    = $SetProgressActivityName
                                     }

            $TenantId = Azure-Login @azureLoginParameters
        }
        else 
        {
            try
            {
                Import-Module -Name Az.Resources -ErrorAction Stop
            }
            catch
            {
                try
                {
                    Import-PackageProvider -Name Nuget -MinimumVersion "2.8.5.201" -ErrorAction Stop
                }
                catch
                {
                    Install-PackageProvider NuGet -Force | Out-Null
                }

                Install-Module -Name Az.Resources -Force -AllowClobber
                Import-Module -Name Az.Resources
            }    
        }

        $armResource = Get-AzResource -ResourceId $armResourceId -ExpandProperties -ErrorAction Stop

        $properties  = $armResource.Properties

        if ($properties.desiredProperties -eq $null)
        {
            #
            # Create desiredProperties object with default values
            #
            $desiredProperties = New-Object -TypeName PSObject
            $desiredProperties | Add-Member -MemberType NoteProperty -Name 'windowsServerSubscription' -Value 'Disabled'
            $desiredProperties | Add-Member -MemberType NoteProperty -Name 'diagnosticLevel' -Value 'Basic'

            $properties | Add-Member -MemberType NoteProperty -Name 'desiredProperties' -Value $desiredProperties
        }

        if ($PSBoundParameters.ContainsKey('EnableWSSubscription'))
        {
            if ($EnableWSSubscription -eq $true)
            {
                $properties.desiredProperties.windowsServerSubscription = 'Enabled';

                $successMessage.Append($SetAzResourceSuccessWSSE) | Out-Null;
            }
            else
            {
                $properties.desiredProperties.windowsServerSubscription = 'Disabled';

                $successMessage.Append($SetAzResourceSuccessWSSD) | Out-Null;
            }

            $doSetResource      = $true
            $needShouldContinue = $true
        }

        if ($PSBoundParameters.ContainsKey('DiagnosticLevel'))
        {
            $properties.desiredProperties.diagnosticLevel = $DiagnosticLevel.ToString()

            if ($successMessage.Length -gt 0)
            {
                $successMessage.AppendFormat(" {0}", ($SetAzResourceSuccessDiagLevel -f $DiagnosticLevel.ToString())) | Out-Null
            }
            else
            {
                $successMessage.AppendFormat("{0}", ($SetAzResourceSuccessDiagLevel -f $DiagnosticLevel.ToString())) | Out-Null
            }

            $doSetResource = $true
        }

        if ($doSetResource -eq $true)
        {
            if ($PSCmdlet.ShouldProcess($armResourceId, $SetProgressShouldProcess))
            {
                if ($needShouldContinue -eq $true)
                {
                    if (($Force -or $PSCmdlet.ShouldContinue($SetProgressShouldContinue, $SetProgressShouldContinueCaption)) -eq $false)
                    {
                        return;
                    }
                }

                Write-Progress -Id $MainProgressBarId -Activity $SetProgressActivityName -Status $SetProgressStatusUpdatingProps -PercentComplete 60

                $setAzResourceParameters = @{
                                            'ResourceId'  = $armResource.Id;
                                            'Properties'  = $properties;
                                            'ApiVersion'  = $RPAPIVersion
                                            }

                $localResult = Set-AzResource @setAzResourceParameters -Confirm:$false -Force -ErrorAction Stop

                if ($PSBoundParameters.ContainsKey('EnableWSSubscription') -and ($EnableWSSubscription -eq $false))
                {
                    Write-Warning -Message $SetProgressWarningWSSD
                }

                if ($PSBoundParameters.ContainsKey('DiagnosticLevel') -and ($DiagnosticLevel -eq [DiagnosticLevel]::Off))
                {
                    Write-Warning -Message $SetProgressWarningDiagnosticOff
                }
            }
            else
            {
                return;
            }
        }

        #
        # Schedule a sync on the cluster
        #
        if($PSBoundParameters.ContainsKey('ResourceId') -eq $false)
        {
            if ($doSetResource -eq $true)
            {
                Write-Progress -Id $MainProgressBarId -Activity $SetProgressActivityName -Status $SetProgressStatusSyncCluster -PercentComplete 90

                Invoke-Command @nodeSessionParams -ScriptBlock { Sync-AzureStackHCI }
            }
        }

        Write-Progress -Id $MainProgressBarId -activity $SetProgressActivityName -Completed

        $setOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyResult -Value ([OperationStatus]::Success)
        $setOutput | Add-Member -MemberType NoteProperty -Name $OutputPropertyDetails -Value ($successMessage.ToString())

        Write-Output $setOutput
    }
    catch
    {
        Write-Error -Exception $_.Exception -Category OperationStopped -ErrorAction Continue

        # Get script line number, offset and Command that resulted in exception. Write-Error with the exception above does not write this info.
        $positionMessage = $_.InvocationInfo.PositionMessage
        Write-Error ("Exception occurred in {0} : {1}" -f $PSCmdlet.MyInvocation.InvocationName, $positionMessage) -Category OperationStopped

        throw
    }
    finally
    {
        if ($doAzAuth -eq $true)
        {
            try { Disconnect-AzAccount | Out-Null } catch{}
            try { Disconnect-AzureAD | Out-Null } catch{}
        }

        Stop-Transcript | out-null
    }
}

#
# IMDS Attestation Section
#
function Add-VMDevicesForImds{
param(
    [hashtable] $VmAdapterParams,
    [hashtable] $VmAdapterAdditionalParams,
    [hashtable] $VmAdapterVlanParams,
    [hashtable] $SessionParams
)
    $ret = @{ 
            Return    = $null
            Exception = $null
    }
    $sc = {
        param([hashtable]$VmAdapterParams, [hashtable]$VmAdapterAdditionalParams, [hashtable]$VmAdapterVlanParams)

        try
        {
            $hostVmSwitch   = $VmAdapterParams.VMSwitch
            $adapterParams  = @{
                    VM      = $VmAdapterParams.VM
                    Name    = $VmAdapterParams.Name
            }

            Write-Information "Checking for previously configured adapter"
            $foundAdapter       = Get-VMNetworkAdapter @adapterParams -ErrorAction SilentlyContinue
            $adapterCount       = ($foundAdapter | Measure-Object).Count

            if ($adapterCount -eq 0)
            {
                Write-Information "Creating IMDS network adapter on guest $($VM.Name)"
                $vmAdapter = Add-VMNetworkAdapter @adapterParams -Confirm: $false -Passthru
            }
            elseif ($adapterCount -eq 1)
            {
                Write-Information "Found existing adapter on guest $($VM.Name)"
                $vmAdapter = $foundAdapter
            }
            else 
            {
                Write-Warning "Found additional IMDS configuration on guest $($VM.Name) adapter count=$($adapterCount)"
                $vmAdapter = $foundAdapter[0]    
            }

            $vmAdapter      = $vmAdapter | Set-VMNetworkAdapter @VmAdapterAdditionalParams -Confirm: $false -Passthru
        
            Connect-VMNetworkAdapter -VMNetworkAdapter $vmAdapter -VMSwitch $hostVmSwitch -Confirm: $false

            $vmAdapter      = Set-VMNetworkAdapterVlan -VMNetworkAdapter $vmAdapter @VmAdapterVlanParams -Confirm: $false -Passthru
        
            $ret.Return = $vmAdapter
            return $ret
        }
        catch
        {
            $ret.Exception = $_
            return $ret
        }
        finally
        {
            if ($ret.Exception) { try{ Remove-VMNetworkAdapter -VMNetworkAdapter $vmAdapter -Force }catch{}}
        }
    }

    $ret = Invoke-Command @SessionParams -ScriptBlock $sc -ArgumentList $VmAdapterParams,$VmAdapterAdditionalParams,$VmAdapterVlanParams
    
    if ($ret.Exception)
    {
        Write-Error "Unable to configure IMDS Service on VM. $($ret.Exception)"
        throw
    }

    return $ret.Return
}

function Add-HostDevicesForImds{
param(
    [hashtable] $VmSwitchParams,
    [hashtable] $HostAdapterVlanParams,
    [hashtable] $NetAdapterIpParams,
    [hashtable] $SessionParams
)
    $sc = {
        param([hashtable]$VmSwitchParams, [hashtable]$HostAdapterVlanParams, [hashtable]$NetAdapterIpParams)

        $ret = @{ 
            Return    = $null
            Exception = $null
        }
        try
        {
            Write-Information "Searching for previous IMDS switch"
            if ($VmSwitchParams.SwitchId)
            {
                $findSwitch         = Get-VMSwitch -Id $VmSwitchParams.SwitchId -ErrorAction SilentlyContinue
            }
            

            $switchCount = ($findSwitch | Measure-Object).Count

            if ($switchCount -eq 0)
            {
                Write-Information "Creating IMDS switch"
                $VmSwitchParams.Remove("SwitchId")
                $hostSwitch     = New-VMSwitch @VmSwitchParams
            }
            elseif ($switchCount -eq 1)
            {
                Write-Information "Found existing IMDS Service Switch."
                $hostSwitch = $findSwitch
            }
        
            $hostVMNetAdapter   = Get-VMNetworkAdapter -ManagementOS -SwitchName $hostSwitch.Name | Where-Object { $_.SwitchId -eq $hostSwitch.Id }

            if (!$hostVMNetAdapter)
            {
                throw("Missing host adapter.")
            }

            $hostNetAdapter     = Get-NetAdapter | Where-Object { ($_.MacAddress -replace "[^a-zA-Z0-9]","") -eq ($hostVMNetAdapter.MacAddress -replace "[^a-zA-Z0-9]","") }

            $nooutput           = $hostNetAdapter | Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue

            $hostNetAdapterIP   = $hostNetAdapter | New-NetIPAddress @NetAdapterIpParams

            $hostNetAdapter     = $hostNetAdapter | Rename-NetAdapter -NewName $hostSwitch.Name -PassThru -ErrorAction SilentlyContinue

            $HostAdapterVlanCommonParams = @{
                VMNetworkAdapter    = $hostVMNetAdapter
            }

            Set-VMNetworkAdapterVlan @HostAdapterVlanCommonParams @HostAdapterVlanParams -Confirm: $false| Out-Null
            
            $ret.Return = $hostSwitch.Id
            return $ret
        }
        catch
        {
            $ret.Exception = $_
            return $ret
        }
        finally
        {
            if ($ret.Exception) { try{ Remove-VMSwitch -VMSwitch $hostSwitch -Force }catch{}}
        }
    }

    $ret = Invoke-Command @SessionParams -ScriptBlock $sc -ArgumentList $VMSwitchParams,$HostAdapterVlanParams,$NetAdapterIpParams

    if ($ret.Exception)
    {
        Write-Error "Unable to configure IMDS Service on host. $($ret.Exception)"
        throw
    }

    return $ret.Return
}


$TemplateHostImdsParams = @{
    Name                    = "AZSHCI_HOST-IMDS_DO_NOT_MODIFY"
    SwitchType              = "Internal"
    Notes                   = "Managed by Azure Stack HCI IMDS Attestation Service"
    Promiscuous             = $true
    PrimaryVlanId           = 10
    SecondaryVlanIdList     = 200
    IPAddress               = "169.254.169.253"
    PrefixLength            = 16
    NetFirewallRuleName     = "AzsHci-ImdsAttestation-Allow-In"
}
$TemplateVmImdsParams = @{
    Name                    = "AZSHCI_GUEST-IMDS_DO_NOT_MODIFY"
    MacAddressSpoofing      = "Off"
    DhcpGuard               = "On"
    RouterGuard             = "On"
    NotMonitoredInCluster   = $true
    Isolated                = $true
    PrimaryVlanId           = 10
    SecondaryVlanId         = 200
}
<#
    .Description
    Enable-AzStackHCIAttestation configures the host and enables specified guests for IMDS attestation.
    
    .PARAMETER ComputerName
    Specifies the AzureStack HCI host to perform the operation on. Note: this host should match the host of VMName.

    .PARAMETER Credential
    Specifies the credential for the ComputerName. Default is the current user executing the Cmdlet.

    .PARAMETER AddVM
    After enabling each cluster node for Attestation, add all guests on each node.

    .PARAMETER Force
    No confirmations.

    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject
    Cluster:     Name of cluster
    Node:        Name of the host.
    Attestation: IMDS Attestation status.

    .EXAMPLE
    Invoking on one of the cluster node.
    C:\PS>Enable-AzStackHCIAttestation -AddVM

    .EXAMPLE
    Invoking from WAC/Management node and adding all existing VMs cluster-wide
    C:\PS>Enable-AzStackHCIAttestation -ComputerName "host1" -AddVM
#>
function Enable-AzStackHCIAttestation{
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([PSCustomObject])]
param(
    [Parameter(Position = 0, Mandatory = $false)]
    [string] $ComputerName,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $false)]
    [switch] $AddVM,

    [Parameter(Mandatory = $false)]
    [switch] $Force
    )

    begin
    {   
        if ($Force)
        {
            $ConfirmPreference = 'None'
        }

        try
        {
            $logPath = "EnableAzureStackHCIImds"
            Setup-Logging -LogFilePrefix $logPath
            #Show-LatestModuleVersion

            $enableImdsOutputList = [System.Collections.ArrayList]::new()
            $HyperVInstallConfirmed = $false

            if([string]::IsNullOrEmpty($ComputerName))
            {
                $ComputerName = [Environment]::MachineName
                $IsManagementNode = $False
            }
            else
            {
                $IsManagementNode = $True
            }

            $percentComplete = 1
            Write-Progress -Id $MainProgressBarId -activity $EnableAzsHciImdsActivity -status $FetchingRegistrationState -percentcomplete $percentComplete
            
            $SessionParams = @{
                    ErrorAction = "Stop"
            }

            if($IsManagementNode)
            {
                $SessionParams.Add("ComputerName", $ComputerName)
                
                if($Null -eq $Credential)
                {
                    $SessionParams.Add("Credential", $Credential)
                }
            }
            else
            {
                # An empty SessionParams will ensure commands run locally without issue
                #$SessionParams.add("ComputerName", "localhost")
            }

            # Validate cluster is registered
            $RegContext = Invoke-Command @SessionParams -ScriptBlock { Get-AzureStackHCI }

            if($RegContext.RegistrationStatus -ne [RegistrationStatus]::Registered)
            {
                throw $ImdsClusterNotRegistered
            }

            $percentComplete = 5
            Write-Progress -Id $MainProgressBarId -activity $EnableAzsHciImdsActivity -status $DiscoveringClusterNodes -percentcomplete $percentComplete

            $ClusterName  = Invoke-Command @SessionParams -ScriptBlock { (Get-Cluster).Name }
            $ClusterNodes = Invoke-Command @SessionParams -ScriptBlock { Get-ClusterNode }

            # Validate Cluster nodes are online
            if (($ClusterNodes | Where {$_.State -ne [Microsoft.FailoverClusters.PowerShell.ClusterNodeState]::Up} | Measure-Object).Count -ne 0)
            {
                throw $AllClusterNodesAreNotOnline
            }

            $percentComplete = 10
            Write-Progress -Id $MainProgressBarId -activity $EnableAzsHciImdsActivity -status $DiscoveringClusterNodes -percentcomplete $percentComplete

            $nodePercentChunk = (100 - ($percentComplete + 5)) / $ClusterNodes.Count / 2

        }
        catch
        {
            Write-Error -Exception $_.Exception -Category OperationStopped
            $positionMessage = $_.InvocationInfo.PositionMessage
            Write-Error ("Exception occurred in Enable-AzueStackHCIImdsAttestation : " + $positionMessage) -Category OperationStopped
            Stop-Transcript | out-null
            throw $_
        }
    }

    Process
    {
        foreach ($node in $ClusterNodes)
        {
            $NodeName = $node.Name
            
            try 
            {
                Write-Information "Enabling IMDS Attestation on $NodeName"
                
                $percentComplete = $percentComplete + ($nodePercentChunk / 2)
                $ConfiguringClusterNode -f $NodeName | % { Write-Progress -Id $MainProgressBarId -activity $EnableAzsHciImdsActivity -status $_ -percentcomplete $percentComplete }

                $SessionParams["ComputerName"] = $NodeName
            
                if ($NodeName -ieq [Environment]::MachineName)
                {
                    $SessionParams.Remove("ComputerName")
                }

                $needHyperV = Invoke-Command @SessionParams -ScriptBlock { (Get-WindowsFeature -Name RSAT-Hyper-V-Tools).Installed -eq $false }   
                if ($needHyperV)
                {
                    if ($Force -or $HyperVInstallConfirmed -or $PSCmdlet.ShouldContinue($ShouldContinueHyperVInstall -f $NodeName, "Install Management Tools"))
                    {
                        if ($HyperVInstallConfirmed -or $PSCmdlet.ShouldProcess("Windows Feature RSAT-Hyper-V-Tools is installed on $($NodeName).", "Install RSAT-Hyper-V-Tools?", ""))
                        {
                            $HyperVInstallConfirmed = $true
                            Invoke-Command @SessionParams -ScriptBlock { Install-WindowsFeature RSAT-Hyper-V-Tools | Out-Null }
                        }
                    }
                    else
                    {
                        throw "Hyper-V RSAT tools required to continue"
                    }
                }
            
                $attestationSwitchId = Invoke-Command @SessionParams -ScriptBlock { (Get-AzureStackHCIAttestation).AttestationSwitchId }

                $HostVmSwitchParams = @{
                                Name                = $TemplateHostImdsParams["Name"]
                                SwitchType          = $TemplateHostImdsParams["SwitchType"]
                                Notes               = $TemplateHostImdsParams["Notes"]
                                SwitchId            = $attestationSwitchId
                }
                $HostAdapterVlanParams = @{
                                Promiscuous         = $TemplateHostImdsParams["Promiscuous"]
                                PrimaryVlanId       = $TemplateHostImdsParams["PrimaryVlanId"]
                                SecondaryVlanIdList = $TemplateHostImdsParams["SecondaryVlanIdList"]
                }
                $NetAdapterIpParams = @{
                                IPAddress           = $TemplateHostImdsParams["IPAddress"]
                                PrefixLength        = $TemplateHostImdsParams["PrefixLength"]
                }

                # Validate or Configure a new switch on host
                if($attestationSwitchId -or $Force -or $PSCmdlet.ShouldContinue($ConfirmEnableImds, "Enable Cluster $($ClusterName)?"))
                {
                    $Force = $true
                    if ($PSCmdlet.ShouldProcess("IMDS Service will be configured/validated on the host $($NodeName).", "A switch managed by the IMDS Service must be configured/validated on the host $($NodeName). Process host?", ""))
                    {
                        $percentComplete = $percentComplete + ($nodePercentChunk / 2)
                        $ConfiguringClusterNode -f $NodeName | % { Write-Progress -Id $MainProgressBarId -activity $EnableAzsHciImdsActivity -status $_ -percentcomplete $percentComplete }
                        
                        $NotifyServiceNewSwitch = !$attestationSwitchId
                        $attestationSwitchId = Add-HostDevicesForImds -VmSwitchParams $HostVmSwitchParams -HostAdapterVlanParams $HostAdapterVlanParams -NetAdapterIpParams $NetAdapterIpParams -SessionParams $SessionParams
                        
                        # Wait for networking stack to stabalize
                        $percentComplete = $percentComplete + ($nodePercentChunk / 2)
                        Start-Sleep 10

                        if ($NotifyServiceNewSwitch)
                        {
                            Invoke-Command @SessionParams -ScriptBlock { param($switchId); Set-AzureStackHCIAttestation -SwitchId $switchId } -ArgumentList $attestationSwitchId | Out-Null
                        }

                        $firewallRule = Invoke-Command @SessionParams -ScriptBlock { param($ruleName) Enable-NetFirewallRule -Name $ruleName } -ArgumentList $TemplateHostImdsParams["NetFirewallRuleName"] 

                        $nodeAttestation = (Invoke-Command @SessionParams -ScriptBlock { Get-AzureStackHCIAttestation })

                        $enableImdsOutput = New-Object -TypeName PSObject
                        $enableImdsOutput | Add-Member -MemberType NoteProperty -Name ComputerName -Value ($nodeAttestation.ComputerName)
                        $enableImdsOutput | Add-Member -MemberType NoteProperty -Name Status -Value ([ImdsAttestationNodeStatus]($nodeAttestation.Status))
                        $enableImdsOutput | Add-Member -MemberType NoteProperty -Name Expiration -Value ($nodeAttestation.Expiration)
                        $enableImdsOutputList.Add($enableImdsOutput) | Out-Null
                    }
                    elseif ($WhatIfPreference.IsPresent)
                    {
                        $attestationSwitchId = "Whatif:$(New-Guid)"
                    }
                }
                else 
                {
                    return
                }          
            }
            catch 
            {
                Write-Error -Exception $_.Exception -Category OperationStopped
                $positionMessage = $_.InvocationInfo.PositionMessage
                Write-Error ("Exception occurred in Enable-AzStackHCIAttestation : " + $positionMessage) -Category OperationStopped
                Stop-Transcript | out-null
                throw $_
            }
        }

        if ($AddVM)
        {
            foreach ($node in $ClusterNodes)
            {
                $NodeName = $node.Name
                
                $SessionParams["ComputerName"] = $NodeName
            
                if ($NodeName -ieq [Environment]::MachineName)
                {
                    $SessionParams.Remove("ComputerName")
                }
                try 
                {
                    Write-Information "Adding VMs to IMDS Attestation on $NodeName"
                    $ConfiguringClusterNode -f $NodeName | % { Write-Progress -Id $MainProgressBarId -activity $EnableAzsHciImdsActivity -status $_ -percentcomplete $percentComplete }

                    Invoke-Command @SessionParams -ScriptBlock { Add-AzStackHCIVMAttestation -AddAll } | Out-Null
                }
                catch 
                {
                    Write-Error -Category OperationStopped $ErrorAddingAllVMs
                }
            }
        }

        Invoke-Command @SessionParams -ScriptBlock { Sync-AzureStackHCI }

        Write-Progress -Id $MainProgressBarId -activity $EnableAzsHciImdsActivity -status "Complete" -percentcomplete 100
    }
    End
    {
        $enableImdsOutputList | Write-Output
        Stop-Transcript | out-null
    }
}

<#
    .Description
    Disable-AzStackHCIAttestation disables IMDS Attestation on the host

    .PARAMETER RemoveVM
    Specifies the guests on each node should be removed from IMDS Attestation before disabling on cluster. Disable cannot continue before guests are removed.
    
    .PARAMETER ComputerName
    Specifies the AzureStack HCI host to perform the operation on.

    .PARAMETER Credential
    Specifies the credential for the ComputerName. Default is the current user executing the Cmdlet.

    .PARAMETER Force
    No confirmation.

    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject
    Cluster:     Name of cluster
    Node:        Name of the host.
    Attestation: IMDS Attestation status.

    .EXAMPLE
    Remove all guests from IMDS Attestation before disabling on cluster nodes.
    C:\PS>Disable-AzStackHCIAttestation -RemoveVM

    .EXAMPLE
    Invoking from the management node/WAC
    C:\PS>Disable-AzStackHCIAttestation -ComputerName "host1"
#>
function Disable-AzStackHCIAttestation{
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([PSCustomObject])]
param(
    [Parameter(Position = 0, Mandatory = $false)]
    [string] $ComputerName,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $false)]
    [switch] $RemoveVM,

    [Parameter(Mandatory = $false)]
    [switch] $Force
    )

    begin
    {   
        try
        {
            $logPath = "DisableAzureStackHCIImds"
            Setup-Logging -LogFilePrefix $logPath
            #Show-LatestModuleVersion

            $disableImdsOutputList = [System.Collections.ArrayList]::new()

            if([string]::IsNullOrEmpty($ComputerName))
            {
                $ComputerName = [Environment]::MachineName
                $IsManagementNode = $False
            }
            else
            {
                $IsManagementNode = $True
            }

            $percentComplete = 1
            Write-Progress -Id $MainProgressBarId -activity $DisableAzsHciImdsActivity -status $FetchingRegistrationState -percentcomplete $percentComplete
            
            $SessionParams = @{
                    ErrorAction = "Stop"
            }

            if($IsManagementNode)
            {
                $SessionParams.Add("ComputerName", $ComputerName)
                
                if($Null -eq $Credential)
                {
                    $SessionParams.Add("Credential", $Credential)
                }
            }
            else
            {
                # An empty SessionParams will ensure commands run locally without issue
                #$SessionParams.add("ComputerName", "localhost")
            }

            $percentComplete = 5
            Write-Progress -Id $MainProgressBarId -activity $DisableAzsHciImdsActivity -status $DiscoveringClusterNodes -percentcomplete $percentComplete

            $ClusterName  = Invoke-Command @SessionParams -ScriptBlock { (Get-Cluster).Name }            
            $ClusterNodes = Invoke-Command @SessionParams -ScriptBlock { Get-ClusterNode }

            foreach ($node in $ClusterNodes)
            {
                $percentComplete += 1
                $CheckingClusterNode -f $node.name | % {Write-Progress -Id $MainProgressBarId -activity $DisableAzsHciImdsActivity -status $_ -percentcomplete $percentComplete}
                $NodeName = $node.Name
                $SessionParams["ComputerName"] = $NodeName
            
                if (!$IsManagementNode -and ($NodeName -ieq $ComputerName))
                {
                    $SessionParams.Remove("ComputerName")
                }

                if (!$RemoveVM)
                {
                    $guests = Invoke-Command @SessionParams -ScriptBlock { Get-AzStackHCIVMAttestation -Local }
                    if (($guests | Measure-Object).Count -ne 0)
                    {
                        throw ("There are still guests connected to IMDS Attestation. Use switch -RemoveVM or Remove-AzStackHCIVMAttestation cmdlet.")
                    }
                }
                else 
                {
                    $RemovingVmImdsFromNode -f $node.name | % {Write-Progress -Id $MainProgressBarId -activity $DisableAzsHciImdsActivity -status $_ -percentcomplete $percentComplete}
                    $removedGuests = Invoke-Command @SessionParams -ScriptBlock { Remove-AzStackHCIVMAttestation -RemoveAll }
                }
            }

            $percentComplete = 10
            Write-Progress -Id $MainProgressBarId -activity $DisableAzsHciImdsActivity -status $DiscoveringClusterNodes -percentcomplete $percentComplete

            $nodePercentChunk = (100 - ($percentComplete + 5)) / $ClusterNodes.Count
        }
        catch
        {
            Write-Error -Exception $_.Exception -Category OperationStopped
            $positionMessage = $_.InvocationInfo.PositionMessage
            Write-Error ("Exception occurred in Enable-AzueStackHCIImdsAttestation : " + $positionMessage) -Category OperationStopped
            Stop-Transcript | out-null
            throw $_
        }
    }

    Process
    {
        if($Force -or $PSCmdlet.ShouldContinue($ConfirmDisableImds, "Disable Cluster $($ClusterName)?"))
        {
            foreach ($node in $ClusterNodes)
            {
                $NodeName = $node.Name
                
                try 
                {
                    Write-Information "Disabling IMDS Attestation on $NodeName"
                    
                    $percentComplete = $percentComplete + ($nodePercentChunk / 2)
                    $DisablingIMDSOnNode -f $NodeName | % {Write-Progress -Id $MainProgressBarId -activity $DisableAzsHciImdsActivity -status $_ -percentcomplete $percentComplete;}
    
                    $SessionParams["ComputerName"] = $NodeName
                
                    if ($NodeName -ieq [Environment]::MachineName)
                    {
                        $SessionParams.Remove("ComputerName")
                    }
                
                    $attestationSwitchId = Invoke-Command @SessionParams -ScriptBlock { (Get-AzureStackHCIAttestation).AttestationSwitchId }
                    if ($attestationSwitchId -ne [Guid]::Empty -and $attestationSwitchId)
                    {
                        Invoke-Command @SessionParams -ScriptBlock { param($switchId); Get-VMSwitch -SwitchId $switchId -ErrorAction SilentlyContinue | Remove-VMSwitch -Force -ErrorAction SilentlyContinue } -ArgumentList $attestationSwitchId
                    }
    
    
                    $percentComplete = $percentComplete + ($nodePercentChunk / 2)
                    $DisablingIMDSOnNode -f $NodeName | % {Write-Progress -Id $MainProgressBarId -activity $DisableAzsHciImdsActivity -status $_ -percentcomplete $percentComplete; }
                    
                    Invoke-Command @SessionParams -ScriptBlock { param($switchId); Set-AzureStackHCIAttestation -SwitchId $switchId } -ArgumentList ([Guid]::Empty) | Out-Null
    
                    $nodeAttestation = (Invoke-Command @SessionParams -ScriptBlock { Get-AzureStackHCIAttestation })
                    $disableImdsOutput = New-Object -TypeName PSObject
                    $disableImdsOutput | Add-Member -MemberType NoteProperty -Name ComputerName -Value ($nodeAttestation.ComputerName)
                    $disableImdsOutput | Add-Member -MemberType NoteProperty -Name Status -Value ([ImdsAttestationNodeStatus]($nodeAttestation.Status))
                    $disableImdsOutput | Add-Member -MemberType NoteProperty -Name Expiration -Value ($nodeAttestation.Expiration)
                    $disableImdsOutputList.Add($disableImdsOutput) | Out-Null
    
                }
                catch 
                {
                    Write-Error -Exception $_.Exception -Category OperationStopped
                    $positionMessage = $_.InvocationInfo.PositionMessage
                    Write-Error ("Exception occurred in Enable-AzueStackHCIImdsAttestation : " + $positionMessage) -Category OperationStopped
                    Stop-Transcript | out-null
                    throw $_
                }
            }
        }

        Invoke-Command @SessionParams -ScriptBlock { Sync-AzureStackHCI }

        Write-Progress -Id $MainProgressBarId -activity $DisableAzsHciImdsActivity -status "Complete" -percentcomplete 100
    }
    End
    {
        $disableImdsOutputList | Write-Output
        Stop-Transcript | out-null
    }
}

<#
    .Description
    Add-AzStackHCIVMAttestation configures guests for AzureStack HCI IMDS Attestation.
    
    .PARAMETER VMName
    Specifies an array of guest VMs to enable.

    .PARAMETER VM
    Specifies an array of VM objects from Get-VM.

    .PARAMETER AddAll
    Specifies a switch that will add all current guest VMs on host to IMDS Attestation on the current node.

    .Parameter Force
    No confirmations.

    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject
    Name:            Name of the VM.
    AttestationHost: Host that VM is currently connected.
    Status:          Connection status.

    .EXAMPLE
    Adding all guests on current node
    C:\PS>Add-AzStackHCIVMAttestation -AddAll

    .EXAMPLE
    Invoking from the management node/WAC
    C:\PS>Invoke-Command -ScriptBlock {Add-AzStackHCIVMAttestation -VMName "guest1", "guest2"} -ComputerName "node1"
#>
function Add-AzStackHCIVMAttestation{
    [CmdletBinding(DefaultParameterSetName="VMName", SupportsShouldProcess)]
    [OutputType([PSCustomObject])]
param(
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "VMName")]
    [string[]] $VMName,

    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "VMObject")]
    [Object[]] $VM,

    [Parameter(Mandatory = $true, ParameterSetName = "AddAll")]
    [Switch]$AddAll,

    [Parameter(Mandatory = $false)]
    [switch] $Force
    )

    begin
    {   
        if ($Force)
        {
            $ConfirmPreference = 'None'
        }

        try
        {
            $logPath = "AddAzureStackHCIImds"
            Setup-Logging -LogFilePrefix $logPath

            $enableImdsOutputList = [System.Collections.ArrayList]::new()
            $ComputerName = [Environment]::MachineName

            $percentcomplete = 1
            Write-Progress -Id $SecondaryProgressBarId -activity $AddAzsHciImdsActivity -status $FetchingRegistrationState -percentcomplete $percentcomplete
            
            $SessionParams = @{
                    ErrorAction = "Stop"
            }

            # Validate cluster is registered
            $RegContext = Invoke-Command @SessionParams -ScriptBlock { Get-AzureStackHCI }

            if($RegContext.RegistrationStatus -ne [RegistrationStatus]::Registered)
            {
                throw $ImdsClusterNotRegistered
            }

            $percentcomplete = 2
            Write-Progress -Id $SecondaryProgressBarId -activity $AddAzsHciImdsActivity -status "Verifying attestation" -percentcomplete $percentComplete

            
            $attestationSwitchId = Invoke-Command @SessionParams -ScriptBlock { (Get-AzureStackHCIAttestation).AttestationSwitchId }

            # Validate or Configure a new switch on host
            if(!$attestationSwitchId)
            {
                $message = $AttestationNotEnabled -f $ComputerName
                throw $message
            }          

            if ($WhatIfPreference.IsPresent)
            {
                $attestationSwitchId = "Whatif:$(New-Guid)"
            }
            
            if ($PSCmdlet.ShouldProcess("Will use IMDS switch $($attestationSwitchId) on $($ComputerName).", "The IMDS switch $($attestationSwitchId) was validated on $($ComputerName). Select and Continue?", ""))
            {
                $attestationSwitch = Invoke-Command @SessionParams -ScriptBlock {param($attestationSwitchId) Get-VMSwitch -Id $attestationSwitchId} -ArgumentList $attestationSwitchId
            }
            else
            {
                return
            }
            

            if ($PSCmdlet.ParameterSetName -eq "AddAll")
            {
                $VirtualMachines = Invoke-Command @SessionParams -ScriptBlock { Get-VM }
                Write-Debug "EnableAll specified. Found ($(($VirtualMachines | Measure-Object).Count) guests VMs."
            }
        }
        catch
        {
            Write-Error -Exception $_.Exception -Category OperationStopped
            $positionMessage = $_.InvocationInfo.PositionMessage
            Write-Error ("Exception occurred in Add-AzStackHCIVMAttestation : " + $positionMessage) -Category OperationStopped
            Stop-Transcript | out-null
            throw $_
        }
    }

    Process
    {
        try 
        {
            if (!$attestationSwitch)
            {
                throw ("Did not validate host configuration")
            }
            Write-Information "Enabling IMDS Attestation on guest virtual machines"
            if ($VMName) 
            {
                $VirtualMachines = Invoke-Command @SessionParams -ScriptBlock {param($vms) Get-VM $vms} -ArgumentList (,$VMName)
            }
            elseif ($VM) 
            {
                $VirtualMachines = $VM
            }
            
            $VmNetAdapterParams = @{
                    Name                    = $TemplateVmImdsParams["Name"]
                    VmSwitch                = $attestationSwitch
            }
            $VmAdapterAdditionalParams = @{
                    MacAddressSpoofing      = $TemplateVmImdsParams["MacAddressSpoofing"]
                    DhcpGuard               = $TemplateVmImdsParams["DhcpGuard"]
                    RouterGuard             = $TemplateVmImdsParams["RouterGuard"]
                    NotMonitoredInCluster   = $TemplateVmImdsParams["NotMonitoredInCluster"]
            }
            $VmAdapterVlanParams = @{
                    Isolated                = $TemplateVmImdsParams["Isolated"]
                    PrimaryVlanId           = $TemplateVmImdsParams["PrimaryVlanId"]
                    SecondaryVlanId         = $TemplateVmImdsParams["SecondaryVlanId"]
            }

            foreach ($vm in $VirtualMachines)
            {
                if ($PSCmdlet.ShouldProcess("Added/Validated $($vm.Name) on host $($attestationSwitch.ComputerName)", "Add/Validate $($vm.Name) to IMDS Attestation on $($attestationSwitch.ComputerName)?", ""))
                {
                    $VmNetAdapterParams["VM"] = $vm
                    $vmAdapter = Add-VMDevicesForImds $VmNetAdapterParams $VmAdapterAdditionalParams $VmAdapterVlanParams $SessionParams
                    
                    $enableImdsOutput = New-Object -TypeName PSObject
                    $enableImdsOutput | Add-Member -MemberType NoteProperty -Name Name -Value $vm.Name
                    $enableImdsOutput | Add-Member -MemberType NoteProperty -Name AttestationHost -Value $ComputerName
                    $enableImdsOutput | Add-Member -MemberType NoteProperty -Name Status -Value ([VMAttestationStatus]::Connected)
                    $enableImdsOutputList.Add($enableImdsOutput) | Out-Null
                }
            } 
            
        }
        catch 
        {
            Write-Error -Exception $_.Exception -Category OperationStopped
            $positionMessage = $_.InvocationInfo.PositionMessage
            Write-Error ("Exception occurred in Add-AzStackHCIVMAttestation : " + $positionMessage) -Category OperationStopped
            Stop-Transcript | out-null
            throw $_
        }
    }
    End
    {
        $enableImdsOutputList | Write-Output
        Stop-Transcript | out-null
    }
}

<#
    .Description
    Remove-AzStackHCIVMAttestation removes guests from AzureStack HCI IMDS Attestation.
    
    .PARAMETER VMName
    Specifies an array of guest VMs to enable.

    .PARAMETER VM
    Specifies an array of VM objects from Get-VM.

    .PARAMETER RemoveAll
    Specifies a switch that will remove all guest VMs from Attestation on the current node

    .PARAMETER Force
    No confirmations.

    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject
    Name:            Name of the VM.
    AttestationHost: Host that VM is currently connected.
    Status:          Connection status.

    .EXAMPLE
    Removing all guests on current node
    C:\PS>Remove-AzStackHCIVMAttestation -RemoveVM

    .EXAMPLE
    Invoking from the management node/WAC
    C:\PS>Invoke-Command -ScriptBlock {Remove-AzStackHCIVMAttestation -VMName "guest1", "guest2"} -ComputerName "node1"
#>
function Remove-AzStackHCIVMAttestation{
    [CmdletBinding(DefaultParameterSetName="VMName", SupportsShouldProcess)]
    [OutputType([PSCustomObject])]
param(
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "VMName")]
    [string[]] $VMName,

    [parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "VMObject")]
    [Object[]] $VM,

    [Parameter(Mandatory = $true, ParameterSetName = "RemoveAll")]
    [Switch]$RemoveAll,

    [Parameter(Mandatory = $false)]
    [switch] $Force
    )

    begin
    {   
        if ($Force)
        {
            $ConfirmPreference = 'None'
        }

        try
        {
            $logPath = "RemoveAzureStackHCIImds"
            Setup-Logging -LogFilePrefix $logPath
            #Show-LatestModuleVersion

            $removeImdsOutputList = [System.Collections.ArrayList]::new()
            $ComputerName = [Environment]::MachineName

            $percentcomplete = 1
            Write-Progress -Id $SecondaryProgressBarId -activity $RemoveAzsHciImdsActivity -status $FetchingRegistrationState -percentcomplete $percentcomplete
            
            $SessionParams = @{
                    ErrorAction = "Stop"
            }

            $percentcomplete = 2
            Write-Progress -Id $SecondaryProgressBarId -activity $RemoveAzsHciImdsActivity -status "Removing guest attestation" -percentcomplete $percentComplete

            if ($PSCmdlet.ParameterSetName -eq "RemoveAll")
            {
                $VirtualMachines = Invoke-Command @SessionParams -ScriptBlock { param($adapterName); Get-VMNetworkAdapter -All -Name $adapterName -ErrorAction SilentlyContinue | % {Get-VM $_.VMId -ErrorAction SilentlyContinue} } -ArgumentList $TemplateVmImdsParams["Name"]
                Write-Debug "RemoveAll specified. Found ($(($VirtualMachines | Measure-Object).Count) guests VMs to remove IMDS Attestation from."
            }
        }
        catch
        {
            Write-Error -Exception $_.Exception -Category OperationStopped
            $positionMessage = $_.InvocationInfo.PositionMessage
            Write-Error ("Exception occurred in Remove-AzStackHCIVMAttestation : " + $positionMessage) -Category OperationStopped
            Stop-Transcript | out-null
            throw $_
        }
    }

    Process
    {
        try 
        {
            Write-Information "Removing IMDS Attestation on guest virtual machines"
            if ($VMName) 
            {
                $VirtualMachines = Invoke-Command @SessionParams -ScriptBlock {param($vms) Get-VM $vms} -ArgumentList (,$VMName)
            }
            elseif ($VM) 
            {
                $VirtualMachines = $VM
            }

            foreach ($vm in $VirtualMachines)
            {
                if ($PSCmdlet.ShouldProcess("Remove IMDS Attestation from $($vm.Name) on host $ComputerName", "Remove $($vm.Name) from IMDS Attestation on $ComputerName?", ""))
                {
                    Invoke-Command @SessionParams -ScriptBlock { param($adapterName); Remove-VMNetworkAdapter -VM $vm -Name $adapterName -ErrorAction Stop } -ArgumentList $TemplateVmImdsParams["Name"]
                    
                    $removeImdsOutput = New-Object -TypeName PSObject
                    $removeImdsOutput | Add-Member -MemberType NoteProperty -Name Name -Value $vm.Name
                    $removeImdsOutput | Add-Member -MemberType NoteProperty -Name AttestationHost -Value $ComputerName
                    $removeImdsOutput | Add-Member -MemberType NoteProperty -Name Status -Value ([VMAttestationStatus]::Disconnected)
                    $removeImdsOutputList.Add($removeImdsOutput) | Out-Null
                }
            }
            
        }
        catch 
        {
            Write-Error -Exception $_.Exception -Category OperationStopped
            $positionMessage = $_.InvocationInfo.PositionMessage
            Write-Error ("Exception occurred in Remove-AzStackHCIVMAttestation : " + $positionMessage) -Category OperationStopped
            Stop-Transcript | out-null
            throw $_
        }
    }
    End
    {
        $removeImdsOutputList | Write-Output
        Stop-Transcript | out-null
    }
}

<#
    .Description
    Get-AzStackHCIVMAttestation shows a list of guests added to IMDS Attestation on a node.

    .PARAMETER Local
    Only retrieve guests with Attestation from the node executing the cmdlet.

    .OUTPUTS
    PSCustomObject. Returns following Properties in PSCustomObject.
    Name:            Name of the VM.
    AttestationHost: Host that VM is currently connected.
    Status:          Connection status.

    .EXAMPLE
    Get all guests on cluster.
    C:\PS>Get-AzStackHCIVMAttestation

    .EXAMPLE
    Get all guests on current node.
    C:\PS>Get-AzStackHCIVMAttestation -Local
#>
function Get-AzStackHCIVMAttestation {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
param(
    [Parameter(Mandatory = $false)]
    [switch] $Local
)

    begin
    {   
        try
        {
            $getImdsOutputList = [System.Collections.ArrayList]::new()
            
            $SessionParams = @{
                    ErrorAction = "Stop"
            }
        }
        catch
        {
            Write-Error -Exception $_.Exception -Category OperationStopped
            $positionMessage = $_.InvocationInfo.PositionMessage
            Write-Error ("Exception occurred in Get-AzStackHCIVMAttestation : " + $positionMessage) -Category OperationStopped
            throw $_
        }
    }

    Process
    {
        try 
        {   
            $nodes = [Environment]::MachineName

            if (!$Local)
            {
                $nodes = (Get-ClusterNode | Select-Object Name).Name
            }

            foreach ($node in $nodes)
            {
                $SessionParams["ComputerName"] = $node
            
                if ($node -ieq [Environment]::MachineName)
                {
                    $SessionParams.Remove("ComputerName")
                }

                try 
                {
                    $VirtualMachinesAdapters = $null
                    $VirtualMachinesAdapters = Invoke-Command @SessionParams -ScriptBlock {param($adapterName); Get-VMNetworkAdapter -All -Name $adapterName -ErrorAction SilentlyContinue} -ArgumentList $TemplateVmImdsParams["Name"]
                }
                catch 
                {
                    Write-Error ("Exception occurred when querying cluster node $NodeName") -Category OperationStopped
                }
                
                foreach ($adapter in $VirtualMachinesAdapters)
                {
                    $getImdsOutput = New-Object -TypeName PSObject
                    $getImdsOutput | Add-Member -MemberType NoteProperty -Name Name -Value $adapter.VMName
                    $getImdsOutput | Add-Member -MemberType NoteProperty -Name AttestationHost -Value $node
                    $getImdsOutput | Add-Member -MemberType NoteProperty -Name Status -Value ([VMAttestationStatus]::Connected)
                    $getImdsOutputList.Add($getImdsOutput) | Out-Null
                }
            }   
        }
        catch 
        {
            Write-Error -Exception $_.Exception -Category OperationStopped
            $positionMessage = $_.InvocationInfo.PositionMessage
            Write-Error ("Exception occurred in Get-AzStackHCIVMAttestation : " + $positionMessage) -Category OperationStopped
            throw $_
        }
    }
    End
    {
        $getImdsOutputList | Write-Output
    }
}

Export-ModuleMember -Function Register-AzStackHCI
Export-ModuleMember -Function Unregister-AzStackHCI
Export-ModuleMember -Function Test-AzStackHCIConnection
Export-ModuleMember -Function Set-AzStackHCI
Export-ModuleMember -Function Enable-AzStackHCIAttestation
Export-ModuleMember -Function Disable-AzStackHCIAttestation
Export-ModuleMember -Function Add-AzStackHCIVMAttestation
Export-ModuleMember -Function Remove-AzStackHCIVMAttestation
Export-ModuleMember -Function Get-AzStackHCIVMAttestation
# SIG # Begin signature block
# MIIjhgYJKoZIhvcNAQcCoIIjdzCCI3MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAYaE0oqULFuxpu
# hrbwD7PzCcDHNipUobULc+R3HHBvGqCCDYEwggX/MIID56ADAgECAhMzAAACUosz
# qviV8znbAAAAAAJSMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjEwOTAyMTgzMjU5WhcNMjIwOTAxMTgzMjU5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDQ5M+Ps/X7BNuv5B/0I6uoDwj0NJOo1KrVQqO7ggRXccklyTrWL4xMShjIou2I
# sbYnF67wXzVAq5Om4oe+LfzSDOzjcb6ms00gBo0OQaqwQ1BijyJ7NvDf80I1fW9O
# L76Kt0Wpc2zrGhzcHdb7upPrvxvSNNUvxK3sgw7YTt31410vpEp8yfBEl/hd8ZzA
# v47DCgJ5j1zm295s1RVZHNp6MoiQFVOECm4AwK2l28i+YER1JO4IplTH44uvzX9o
# RnJHaMvWzZEpozPy4jNO2DDqbcNs4zh7AWMhE1PWFVA+CHI/En5nASvCvLmuR/t8
# q4bc8XR8QIZJQSp+2U6m2ldNAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUNZJaEUGL2Guwt7ZOAu4efEYXedEw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDY3NTk3MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAFkk3
# uSxkTEBh1NtAl7BivIEsAWdgX1qZ+EdZMYbQKasY6IhSLXRMxF1B3OKdR9K/kccp
# kvNcGl8D7YyYS4mhCUMBR+VLrg3f8PUj38A9V5aiY2/Jok7WZFOAmjPRNNGnyeg7
# l0lTiThFqE+2aOs6+heegqAdelGgNJKRHLWRuhGKuLIw5lkgx9Ky+QvZrn/Ddi8u
# TIgWKp+MGG8xY6PBvvjgt9jQShlnPrZ3UY8Bvwy6rynhXBaV0V0TTL0gEx7eh/K1
# o8Miaru6s/7FyqOLeUS4vTHh9TgBL5DtxCYurXbSBVtL1Fj44+Od/6cmC9mmvrti
# yG709Y3Rd3YdJj2f3GJq7Y7KdWq0QYhatKhBeg4fxjhg0yut2g6aM1mxjNPrE48z
# 6HWCNGu9gMK5ZudldRw4a45Z06Aoktof0CqOyTErvq0YjoE4Xpa0+87T/PVUXNqf
# 7Y+qSU7+9LtLQuMYR4w3cSPjuNusvLf9gBnch5RqM7kaDtYWDgLyB42EfsxeMqwK
# WwA+TVi0HrWRqfSx2olbE56hJcEkMjOSKz3sRuupFCX3UroyYf52L+2iVTrda8XW
# esPG62Mnn3T8AuLfzeJFuAbfOSERx7IFZO92UPoXE1uEjL5skl1yTZB3MubgOA4F
# 8KoRNhviFAEST+nG8c8uIsbZeb08SeYQMqjVEmkwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVWzCCFVcCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAlKLM6r4lfM52wAAAAACUjAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg45enVygm
# 81Ei6IkT4v6DiqRHsf0B3rWEEBMclTiojhYwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQAmyHcyp129Te2gIrRzh0CFhE1vYRl5W/1JK5kK4FlE
# 1ajAWRdMs2CwxwC/Q+ng4ILGSaA6LUFq35Y04YiJA3UzcNFZi2vPxiOedo7Yc1u7
# Yoaszub2WfWrpWQp/9krWl/ApZBbJ4WfUhtPE3ALelfxNOmesA1Xg5l85Yuytp7V
# paliXcc6IPHqT1MLas6ZIko55o7OOmgfUy+sFVYj8ZJqSOReskxCxQkqiaD7DcVC
# v2jJqy0EWFn7Mz3nDPbNK/lUSIcH5Ubo96r2pmrSgcw7VL09npBvf0ZHffr3fQHV
# 4uKCJfa7E3J6l49bdeeqdLDFomtv2z3+WiSY0dmeVev8oYIS5TCCEuEGCisGAQQB
# gjcDAwExghLRMIISzQYJKoZIhvcNAQcCoIISvjCCEroCAQMxDzANBglghkgBZQME
# AgEFADCCAVEGCyqGSIb3DQEJEAEEoIIBQASCATwwggE4AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIMXdRcrnwHsUx0AgPP2/LA1lSGD7Yneds42YDiWi
# GPM6AgZhkuD7WXMYEzIwMjExMjAyMTAwNDE3LjQwOFowBIACAfSggdCkgc0wgcox
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1p
# Y3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOkREOEMtRTMzNy0yRkFFMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIOPDCCBPEwggPZoAMCAQICEzMAAAFOjLHr7dey4wAAAAAAAU4w
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MjAxMTEyMTgyNjAxWhcNMjIwMjExMTgyNjAxWjCByjELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2Eg
# T3BlcmF0aW9uczEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046REQ4Qy1FMzM3LTJG
# QUUxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCG+5vo9Ur9k7nCE6alU9k1Av/D5G0b
# RLOSQRLfl76/siJwDVvlJs9rsnxmXoq2Vu/5BCVnAi69b0nUIrJNXQRPrxBby1kL
# 2WWjgAy4OWNlhTzYWN8SYLA1OqwjvBNncr1VejeHI018G1e5w8YwqwBhuK/IahIC
# M/t8UoTBIhKPsbG3NCInczU5GgHerG0Myp7ug9+8Es6joAl2pu88GefHg48ROnCG
# Avmb3xPppdhUHzpSwPhjLvMHPnilQAN2IjQcnArxdBQ3I6llOEIWwJdoin2GG/Fi
# VMyvK92bWOCwZSj42pcBXNNsob0So9yxRJXfHSuyU/fMgfrXTOq0ho2pAgMBAAGj
# ggEbMIIBFzAdBgNVHQ4EFgQUojVREyZC4/ay6+fmwmlq2qZgGeEwHwYDVR0jBBgw
# FoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
# L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljVGltU3RhUENB
# XzIwMTAtMDctMDEuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAx
# MC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDAN
# BgkqhkiG9w0BAQsFAAOCAQEAo4mcyK2Sr4FlF5VgTkRd1POeVebEWCvJjhs1IqbV
# fSJefNWXL5iYLxc2fJscNe7i86yrbBfsThj8uvQV7lx0JEGt/NT6nlUnYxyJB2ZK
# N1pPACcKMmHLeXUL6BMrgaE9Vl5zJQyr5hGfa6GLQeXert/8WxK45fusANXFqzEO
# B8pgwydlhxaFr+R7YH8ec++EJm+yIIF6tC1n5YvWy4mQNKBkFuk52FxDKoISQ02u
# txzLVmK3wRE3SVbaGQ0OixF65cymVOWmLIEFmyi0mGkI5kvKQBpbgl8foOKNrw0F
# 8+Q5Us6AfoJ11rbK5HUm3Utq975SKwcAVzAJCeM6YZW5lzCCBnEwggRZoAMCAQIC
# CmEJgSoAAAAAAAIwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIx
# NDY1NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF
# ++18aEssX8XD5WHCdrc+Zitb8BVTJwQxH0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRD
# DNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVHgc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSx
# z5NMksHEpl3RYRNuKMYa+YaAu99h/EbBJx0kZxJyGiGKr0tkiVBisV39dx898Fd1
# rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL/W7lmsqxqPJ6Kgox8NpOBpG2iAg16Hgc
# sOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB
# 4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqF
# bVUwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYD
# VR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwv
# cHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEB
# BE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwgaAGA1UdIAEB/wSBlTCB
# kjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsGAQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vUEtJL2RvY3MvQ1BTL2RlZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQe
# MiAdAEwAZQBnAGEAbABfAFAAbwBsAGkAYwB5AF8AUwB0AGEAdABlAG0AZQBuAHQA
# LiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUx
# vs8F4qn++ldtGTCzwsVmyWrf9efweL3HqJ4l4/m87WtUVwgrUYJEEvu5U4zM9GAS
# inbMQEBBm9xcF/9c+V4XNZgkVkt070IQyK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1
# L3mBZdmptWvkx872ynoAb0swRCQiPM/tA6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWO
# M7tiX5rbV0Dp8c6ZZpCM/2pif93FSguRJuI57BlKcWOdeyFtw5yjojz6f32WapB4
# pm3S4Zz5Hfw42JT0xqUKloakvZ4argRCg7i1gJsiOCC1JeVk7Pf0v35jWSUPei45
# V3aicaoGig+JFrphpxHLmtgOR5qAxdDNp9DvfYPw4TtxCd9ddJgiCGHasFAeb73x
# 4QDf5zEHpJM692VHeOj4qEir995yfmFrb3epgcunCaw5u+zGy9iCtHLNHfS4hQEe
# gPsbiSpUObJb2sgNVZl6h3M7COaYLeqN4DMuEin1wC9UJyH3yKxO2ii4sanblrKn
# QqLJzxlBTeCG+SqaoxFmMNO7dDJL32N79ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp
# 3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zPWAUu7w2gUDXa7wknHNWzfjUeCLraNtvT
# X4/edIhJEqGCAs4wggI3AgEBMIH4oYHQpIHNMIHKMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBP
# cGVyYXRpb25zMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpERDhDLUUzMzctMkZB
# RTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcG
# BSsOAwIaAxUAg8uPxL0/+sO+NO9xWDx5US8QfgKggYMwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIFAOVSdzswIhgPMjAy
# MTEyMDIwNjM1MzlaGA8yMDIxMTIwMzA2MzUzOVowdzA9BgorBgEEAYRZCgQBMS8w
# LTAKAgUA5VJ3OwIBADAKAgEAAgIZ0gIB/zAHAgEAAgIRKTAKAgUA5VPIuwIBADA2
# BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIB
# AAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBAHhaJBku3NjFxnnm6jY6ZT1WU7xhXsDF
# YIzl730xzRjX1Z6EkDgt4BtBdC/Hn9C/YJvuvgJks1AAZMP1DCaU0XtcYLiHeoii
# VEgOrZAn7RhoNGNyjdWd1rtIW9vw5QXxfvtSGivLL7EmKC+Ooq2A0Nunv5NUmqXO
# iP/G0uvHqNRsMYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# IDIwMTACEzMAAAFOjLHr7dey4wAAAAAAAU4wDQYJYIZIAWUDBAIBBQCgggFKMBoG
# CSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQggkVBZpQS
# 6jNAMfGXOhs8NcIE8+VuOVAzkpAjpydKrFowgfoGCyqGSIb3DQEJEAIvMYHqMIHn
# MIHkMIG9BCAI/g3imEuLgecw/rodQgpE3e8yMSuIAo7+6n3jyiUvkjCBmDCBgKR+
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAABToyx6+3XsuMAAAAA
# AAFOMCIEII+NnTJBAVpsxmqbPB1JTNO539grDK87NQHEmuslToGQMA0GCSqGSIb3
# DQEBCwUABIIBAIB2cPn6dVmb3yLPwgQ05rIptskoUA7WrbjpigLhULmmIQT6+B6Z
# Ntp0MOIpC0pSHTHp7l21GC1xsxd/M1FduQ3ccqVpxNoViSZZD9OODzXORmCf5qAj
# 13w+HoDI/N85P9i7IG6wGihisL4cm+aY/062ZS/C+W5/2kGV757yMgEbbBs0qjXy
# Uc4xadfi4wUcI2x5a8cJcPF53axIIZARdiolfGhE0CshdW44fNsbPnHVNTU1BQxc
# Cp4Wbdq3rY6BuyX2mrwe4IiTSd/8zyjpwdCbHhDPzusM/ho49cQcj8zm0hqY2/Wv
# U1Mjk7W0+PrrNxa5F1Bl4S0Z6dKvooMPjEM=
# SIG # End signature block
