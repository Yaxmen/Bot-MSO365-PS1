
# ----------------------------------------------------------------------------------
#
# Copyright Microsoft Corporation
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# http://www.apache.org/licenses/LICENSE-2.0
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
# ----------------------------------------------------------------------------------

<#
.Synopsis
Starts replication for the specified server.
.Description
The New-AzMigrateServerReplication cmdlet starts the replication for a particular discovered server in the Azure Migrate project.
.Link
https://docs.microsoft.com/powershell/module/az.migrate/new-azmigrateserverreplication
#>
function New-AzMigrateServerReplication {
    [OutputType([Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.IJob])]
    [CmdletBinding(DefaultParameterSetName = 'ByIdDefaultUser', PositionalBinding = $false)]
    param(
        [Parameter(ParameterSetName = 'ByIdDefaultUser', Mandatory)]
        [Parameter(ParameterSetName = 'ByIdPowerUser', Mandatory)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the machine ID of the discovered server to be migrated.
        ${MachineId},

        [Parameter(ParameterSetName = 'ByInputObjectDefaultUser', Mandatory)]
        [Parameter(ParameterSetName = 'ByInputObjectPowerUser', Mandatory)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api202001.IVMwareMachine]
        # Specifies the discovered server to be migrated. The server object can be retrieved using the Get-AzMigrateServer cmdlet.
        ${InputObject},

        [Parameter(ParameterSetName = 'ByIdPowerUser', Mandatory)]
        [Parameter(ParameterSetName = 'ByInputObjectPowerUser', Mandatory)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.IVMwareCbtDiskInput[]]
        # Specifies the disks on the source server to be included for replication.
        ${DiskToInclude},

        [Parameter(Mandatory)]
        [ValidateSet("NoLicenseType" , "WindowsServer")]
        [ArgumentCompleter( { "NoLicenseType" , "WindowsServer" })]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies if Azure Hybrid benefit is applicable for the source server to be migrated.
        ${LicenseType},

        [Parameter()]
        [ValidateSet("NoLicenseType" , "PAYG" , "AHUB")]
        [ArgumentCompleter( { "NoLicenseType" , "PAYG" , "AHUB" })]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies if Azure Hybrid benefit for SQL Server is applicable for the server to be migrated.
        ${SqlServerLicenseType},

        [Parameter(Mandatory)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the Resource Group id within the destination Azure subscription to which the server needs to be migrated.
        ${TargetResourceGroupId},

        [Parameter(Mandatory)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the Virtual Network id within the destination Azure subscription to which the server needs to be migrated.
        ${TargetNetworkId},

        [Parameter(Mandatory)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the Subnet name within the destination Virtual Netowk to which the server needs to be migrated.
        ${TargetSubnetName},

        [Parameter(DontShow)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Mapping.
        ${ReplicationContainerMapping},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Account id.
        ${VMWarerunasaccountID},

        [Parameter(Mandatory)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the name of the Azure VM to be created.
        ${TargetVMName},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the SKU of the Azure VM to be created.
        ${TargetVMSize},

        [Parameter(ParameterSetName = 'ByIdDefaultUser')]
        [Parameter(ParameterSetName = 'ByInputObjectDefaultUser')]
        [Parameter(ParameterSetName = 'ByIdPowerUser')]
        [Parameter(ParameterSetName = 'ByInputObjectPowerUser')]
        [ValidateSet("true" , "false")]
        [ArgumentCompleter( { "true" , "false" })]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies if replication be auto-repaired in case change tracking is lost for the source server under replication.
        ${PerformAutoResync},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the Availability Set to be used for VM creationSpecifies the Availability Set to be used for VM creation.
        ${TargetAvailabilitySet},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the Availability Zone to be used for VM creation.
        ${TargetAvailabilityZone},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.IVMwareCbtEnableMigrationInputTargetVmtags]
        # Specifies the tag to be used for VM creation.
        ${VMTag},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.IVMwareCbtEnableMigrationInputTargetNicTags]
        # Specifies the tag to be used for NIC creation.
        ${NicTag},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.IVMwareCbtEnableMigrationInputTargetDiskTags]
        # Specifies the tag to be used for disk creation.
        ${DiskTag},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.Collections.Hashtable]
        # Specifies the tag to be used for Resource creation.
        ${Tag},

        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the storage account to be used for boot diagnostics.
        ${TargetBootDiagnosticsStorageAccount},

        [Parameter(ParameterSetName = 'ByIdDefaultUser', Mandatory)]
        [Parameter(ParameterSetName = 'ByInputObjectDefaultUser', Mandatory)]
        [ValidateSet("Standard_LRS" , "Premium_LRS", "StandardSSD_LRS")]
        [ArgumentCompleter( { "Standard_LRS" , "Premium_LRS", "StandardSSD_LRS" })]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the type of disks to be used for the Azure VM.
        ${DiskType},
        
        [Parameter(ParameterSetName = 'ByIdDefaultUser', Mandatory)]
        [Parameter(ParameterSetName = 'ByInputObjectDefaultUser', Mandatory)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the Operating System disk for the source server to be migrated.
        ${OSDiskID},

        [Parameter(ParameterSetName = 'ByIdDefaultUser')]
        [Parameter(ParameterSetName = 'ByInputObjectDefaultUser')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [System.String]
        # Specifies the disk encyption set to be used.
        ${DiskEncryptionSetID},
    
        [Parameter()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Path')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Runtime.DefaultInfo(Script = '(Get-AzContext).Subscription.Id')]
        [System.String]
        # Azure Subscription ID.
        ${SubscriptionId},

        [Parameter()]
        [Alias('AzureRMContext', 'AzureCredential')]
        [ValidateNotNull()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Azure')]
        [System.Management.Automation.PSObject]
        # The credentials, account, tenant, and subscription used for communication with Azure.
        ${DefaultProfile},
    
        [Parameter(DontShow)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Runtime')]
        [System.Management.Automation.SwitchParameter]
        # Wait for .NET debugger to attach
        ${Break},
    
        [Parameter(DontShow)]
        [ValidateNotNull()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Runtime')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Runtime.SendAsyncStep[]]
        # SendAsync Pipeline Steps to be appended to the front of the pipeline
        ${HttpPipelineAppend},
    
        [Parameter(DontShow)]
        [ValidateNotNull()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Runtime')]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Runtime.SendAsyncStep[]]
        # SendAsync Pipeline Steps to be prepended to the front of the pipeline
        ${HttpPipelinePrepend},
    
        [Parameter(DontShow)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Runtime')]
        [System.Uri]
        # The URI for the proxy server to use
        ${Proxy},
    
        [Parameter(DontShow)]
        [ValidateNotNull()]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Runtime')]
        [System.Management.Automation.PSCredential]
        # Credentials for a proxy server to use for the remote call
        ${ProxyCredential},
    
        [Parameter(DontShow)]
        [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Category('Runtime')]
        [System.Management.Automation.SwitchParameter]
        # Use the default credentials for the proxy
        ${ProxyUseDefaultCredentials}
    )
    
    process {
        $parameterSet = $PSCmdlet.ParameterSetName
        $HasRunAsAccountId = $PSBoundParameters.ContainsKey('VMWarerunasaccountID')
        $HasTargetAVSet = $PSBoundParameters.ContainsKey('TargetAvailabilitySet')
        $HasTargetAVZone = $PSBoundParameters.ContainsKey('TargetAvailabilityZone')
        $HasVMTag = $PSBoundParameters.ContainsKey('VMTag')
        $HasNicTag = $PSBoundParameters.ContainsKey('NicTag')
        $HasDiskTag = $PSBoundParameters.ContainsKey('DiskTag')
        $HasTag = $PSBoundParameters.ContainsKey('Tag')
        $HasSqlServerLicenseType = $PSBoundParameters.ContainsKey('SqlServerLicenseType')
        $HasTargetBDStorage = $PSBoundParameters.ContainsKey('TargetBootDiagnosticsStorageAccount')
        $HasResync = $PSBoundParameters.ContainsKey('PerformAutoResync')
        $HasDiskEncryptionSetID = $PSBoundParameters.ContainsKey('DiskEncryptionSetID')
        $HasTargetVMSize = $PSBoundParameters.ContainsKey('TargetVMSize')

        $null = $PSBoundParameters.Remove('ReplicationContainerMapping')
        $null = $PSBoundParameters.Remove('VMWarerunasaccountID')
        $null = $PSBoundParameters.Remove('TargetAvailabilitySet')
        $null = $PSBoundParameters.Remove('TargetAvailabilityZone')
        $null = $PSBoundParameters.Remove('VMTag')
        $null = $PSBoundParameters.Remove('NicTag')
        $null = $PSBoundParameters.Remove('DiskTag')
        $null = $PSBoundParameters.Remove('Tag')
        $null = $PSBoundParameters.Remove('TargetBootDiagnosticsStorageAccount')
        $null = $PSBoundParameters.Remove('MachineId')
        $null = $PSBoundParameters.Remove('DiskToInclude')
        $null = $PSBoundParameters.Remove('TargetResourceGroupId')
        $null = $PSBoundParameters.Remove('TargetNetworkId')
        $null = $PSBoundParameters.Remove('TargetSubnetName')
        $null = $PSBoundParameters.Remove('TargetVMName')
        $null = $PSBoundParameters.Remove('TargetVMSize')
        $null = $PSBoundParameters.Remove('PerformAutoResync')
        $null = $PSBoundParameters.Remove('DiskType')
        $null = $PSBoundParameters.Remove('OSDiskID')
        $null = $PSBoundParameters.Remove('SqlServerLicenseType')
        $null = $PSBoundParameters.Remove('LicenseType')
        $null = $PSBoundParameters.Remove('DiskEncryptionSetID')

        $null = $PSBoundParameters.Remove('MachineId')
        $null = $PSBoundParameters.Remove('InputObject')

        $validLicenseSpellings = @{ 
            NoLicenseType = "NoLicenseType";
            WindowsServer = "WindowsServer"
        }
        $LicenseType = $validLicenseSpellings[$LicenseType]

        if ($parameterSet -match "DefaultUser") {
            $validDiskTypeSpellings = @{ 
                Standard_LRS    = "Standard_LRS";
                Premium_LRS     = "Premium_LRS";
                StandardSSD_LRS = "StandardSSD_LRS"
            }
            $DiskType = $validDiskTypeSpellings[$DiskType]
            
            if ($parameterSet -eq "ByInputObjectDefaultUser") {
                foreach ($onPremDisk in $InputObject.Disk) {
                    if ($onPremDisk.Uuid -eq $OSDiskID) {
                        $OSDiskID = $onPremDisk.Uuid
                        break
                    }
                }
            }
        }

        # Get the discovered machine Id.
        if (($parameterSet -match 'InputObject')) {
            $MachineId = $InputObject.Id
        }

        # Get the discovered machine object.
        $MachineIdArray = $MachineId.Split("/")
        $SiteType = $MachineIdArray[7]
        $SiteName = $MachineIdArray[8]
        $ResourceGroupName = $MachineIdArray[4]
        $MachineName = $MachineIdArray[10]

        $null = $PSBoundParameters.Add("Name", $MachineName)
        $null = $PSBoundParameters.Add("ResourceGroupName", $ResourceGroupName)
        $null = $PSBoundParameters.Add("SiteName", $SiteName)
        $InputObject = Get-AzMigrateMachine @PSBoundParameters

        $null = $PSBoundParameters.Remove('Name')
        $null = $PSBoundParameters.Remove('ResourceGroupName')
        $null = $PSBoundParameters.Remove('SiteName')
        
        # Get the site to get project name.
        $null = $PSBoundParameters.Add('ResourceGroupName', $ResourceGroupName)
        $null = $PSBoundParameters.Add('SiteName', $SiteName)
        $siteObject = Az.Migrate\Get-AzMigrateSite @PSBoundParameters
        if ($siteObject -and ($siteObject.Count -ge 1)) {
            $ProjectName = $siteObject.DiscoverySolutionId.Split("/")[8]
        }
        else {
            throw "Site not found"
        }
            
        $null = $PSBoundParameters.Remove('ResourceGroupName')
        $null = $PSBoundParameters.Remove('SiteName')

        # Get the solution to get vault name.
        $null = $PSBoundParameters.Add("ResourceGroupName", $ResourceGroupName)
        $null = $PSBoundParameters.Add("Name", "Servers-Migration-ServerMigration")
        $null = $PSBoundParameters.Add("MigrateProjectName", $ProjectName)
            
        $solution = Az.Migrate\Get-AzMigrateSolution @PSBoundParameters
        $VaultName = $solution.DetailExtendedDetail.AdditionalProperties.vaultId.Split("/")[8]
            
        $null = $PSBoundParameters.Remove('ResourceGroupName')
        $null = $PSBoundParameters.Remove("Name")
        $null = $PSBoundParameters.Remove("MigrateProjectName")

        if ($SiteType -ne "VMwareSites") {
            throw "Provider not supported"
        }
           
        # This supports Multi-Vcenter feature.
        if (!$HasRunAsAccountId) {

            # Get the VCenter object.
            $vcenterId = $InputObject.VCenterId
            if ($null -eq $vcenterId){
                throw "Cannot find Vcenter ID in discovered machine."
            }

            $vCenterIdArray = $vcenterId.Split("/")
            $vCenterName = $vCenterIdArray[10] 
            $vCenterSite = $vCenterIdArray[8]
            $vCenterResourceGroupName = $vCenterIdArray[4]

            $null = $PSBoundParameters.Add("Name", $vCenterName)
            $null = $PSBoundParameters.Add("ResourceGroupName", $vCenterResourceGroupName)
            $null = $PSBoundParameters.Add("SiteName", $vCenterSite)

            $vCenter = Get-AzMigrateVCenter @PSBoundParameters

            $null = $PSBoundParameters.Remove('Name')
            $null = $PSBoundParameters.Remove('ResourceGroupName')
            $null = $PSBoundParameters.Remove('SiteName')

            # Get the run as account Id.
            $VMWarerunasaccountID = $vCenter.RunAsAccountId
            if ($VMWarerunasaccountID -eq "") {
                throw "Run As Account missing."
            } 
        }

        $policyName = "migrate" + $SiteName + "policy"
        $null = $PSBoundParameters.Add('ResourceGroupName', $ResourceGroupName)
        $null = $PSBoundParameters.Add('ResourceName', $VaultName)
        $null = $PSBoundParameters.Add('PolicyName', $policyName)
        $policyObj = Az.Migrate\Get-AzMigrateReplicationPolicy @PSBoundParameters -ErrorVariable notPresent -ErrorAction SilentlyContinue
        if ($policyObj -and ($policyObj.Count -ge 1)) {
            $PolicyId = $policyObj.Id
        }
        else {
            throw "The replication infrastructure is not initialized. Run the initialize-azmigratereplicationinfrastructure script again."
        }
        $null = $PSBoundParameters.Remove('ResourceGroupName')
        $null = $PSBoundParameters.Remove('ResourceName')
        $null = $PSBoundParameters.Remove('PolicyName')

        $null = $PSBoundParameters.Add('ResourceGroupName', $ResourceGroupName)
        $null = $PSBoundParameters.Add('ResourceName', $VaultName)
        $allFabrics = Az.Migrate\Get-AzMigrateReplicationFabric @PSBoundParameters
        $FabricName = ""
        if ($allFabrics -and ($allFabrics.length -gt 0)) {
            foreach ($fabric in $allFabrics) {
                if (($fabric.Property.CustomDetail.InstanceType -ceq "VMwareV2") -and ($fabric.Property.CustomDetail.VmwareSiteId.Split("/")[8] -ceq $SiteName)) {
                    $FabricName = $fabric.Name
                    break
                }
            }
        }
        if ($FabricName -eq "") {
            throw "Fabric not found for given resource group."
        }
                
        $null = $PSBoundParameters.Add('FabricName', $FabricName)
        $peContainers = Az.Migrate\Get-AzMigrateReplicationProtectionContainer @PSBoundParameters
        $ProtectionContainerName = ""
        if ($peContainers -and ($peContainers.length -gt 0)) {
            $ProtectionContainerName = $peContainers[0].Name
        }

        if ($ProtectionContainerName -eq "") {
            throw "Container not found for given resource group."
        }

        $mappingName = "containermapping"
        $null = $PSBoundParameters.Add('MappingName', $mappingName)
        $null = $PSBoundParameters.Add("ProtectionContainerName", $ProtectionContainerName)

        $mappingObject = Az.Migrate\Get-AzMigrateReplicationProtectionContainerMapping @PSBoundParameters -ErrorVariable notPresent -ErrorAction SilentlyContinue
        if ($mappingObject -and ($mappingObject.Count -ge 1)) {
            $TargetRegion = $mappingObject.ProviderSpecificDetail.TargetLocation
        }
        else {
            throw "The replication infrastructure is not initialized. Run the initialize-azmigratereplicationinfrastructure script again."
        }
        $null = $PSBoundParameters.Remove('MappingName')

        # Validate sku size
        $hasAzComputeModule = $true
        try { Import-Module Az.Compute }catch { $hasAzComputeModule = $false }
        if ($hasAzComputeModule -and $HasTargetVMSize) {
            $null = $PSBoundParameters.Remove("ProtectionContainerName")
            $null = $PSBoundParameters.Remove("FabricName")
            $null = $PSBoundParameters.Remove('ResourceGroupName')
            $null = $PSBoundParameters.Remove('ResourceName')
            $null = $PSBoundParameters.Remove('SubscriptionId')
            $null = $PSBoundParameters.Add('Location', $TargetRegion)
            $allAvailableSkus = Get-AzVMSize @PSBoundParameters -ErrorVariable notPresent -ErrorAction SilentlyContinue
            if ($null -ne $allAvailableSkus) {
                $matchingComputeSku = $allAvailableSkus | Where-Object { $_.Name -eq $TargetVMSize }
                if ($null -ne $matchingComputeSku) {
                    $TargetVMSize = $matchingComputeSku.Name
                }
            }

            $null = $PSBoundParameters.Remove('Location')
            $null = $PSBoundParameters.Add("ProtectionContainerName", $ProtectionContainerName)
            $null = $PSBoundParameters.Add('FabricName', $FabricName)
            $null = $PSBoundParameters.Add("ResourceGroupName", $ResourceGroupName)
            $null = $PSBoundParameters.Add('ResourceName', $VaultName)
            $null = $PSBoundParameters.Add('SubscriptionId', $SubscriptionId)
        }

        $HashCodeInput = $SiteName + $TargetRegion
        $Source = @"
using System;
public class HashFunctions
{
public static int hashForArtifact(String artifact)
{
        int hash = 0;
        int al = artifact.Length;
        int tl = 0;
        char[] ac = artifact.ToCharArray();
        while (tl < al)
        {
            hash = ((hash << 5) - hash) + ac[tl++] | 0;
        }
        return Math.Abs(hash);
}
}
"@
        Add-Type -TypeDefinition $Source -Language CSharp
        $hash = [HashFunctions]::hashForArtifact($HashCodeInput)

        $LogStorageAccountID = "/subscriptions/" + $SubscriptionId + "/resourceGroups/" +
        $ResourceGroupName + "/providers/Microsoft.Storage/storageAccounts/migratelsa" + $hash
        $LogStorageAccountSas = "migratelsa" + $hash + '-cacheSas'
        if (!$HasTargetBDStorage) {
            $TargetBootDiagnosticsStorageAccount = $LogStorageAccountID
        }

        # Storage accounts need to be in the same subscription as that of the VM.
        if (($null -ne $TargetBootDiagnosticsStorageAccount) -and ($TargetBootDiagnosticsStorageAccount.length -gt 1)) {
            $TargetBDSASubscriptionId = $TargetBootDiagnosticsStorageAccount.Split('/')[2]
            $TargetSubscriptionId = $TargetResourceGroupId.Split('/')[2]
            if ($TargetBDSASubscriptionId -ne $TargetSubscriptionId) {
                $TargetBootDiagnosticsStorageAccount = $null
            }
        }
            
        if (!$HasResync) {
            $PerformAutoResync = "true"
        }
        $validBooleanSpellings = @{ 
            true  = "true";
            false = "false"
        }
        $PerformAutoResync = $validBooleanSpellings[$PerformAutoResync]

        $null = $PSBoundParameters.Add("MigrationItemName", $MachineName)
        $null = $PSBoundParameters.Add("PolicyId", $PolicyId)

        $ProviderSpecificDetails = [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.VMwareCbtEnableMigrationInput]::new()
        $ProviderSpecificDetails.DataMoverRunAsAccountId = $VMWarerunasaccountID
        $ProviderSpecificDetails.SnapshotRunAsAccountId = $VMWarerunasaccountID
        $ProviderSpecificDetails.InstanceType = 'VMwareCbt'
        $ProviderSpecificDetails.LicenseType = $LicenseType
        $ProviderSpecificDetails.PerformAutoResync = $PerformAutoResync
        if ($HasTargetAVSet) {
            $ProviderSpecificDetails.TargetAvailabilitySetId = $TargetAvailabilitySet
        }
        if ($HasTargetAVZone) {
            $ProviderSpecificDetails.TargetAvailabilityZone = $TargetAvailabilityZone
        }

        if ($HasSqlServerLicenseType) {
            $validSqlLicenseSpellings = @{ 
                NoLicenseType = "NoLicenseType";
                PAYG          = "PAYG";
                AHUB          = "AHUB"
            }
            $SqlServerLicenseType = $validSqlLicenseSpellings[$SqlServerLicenseType]
            $ProviderSpecificDetails.SqlServerLicenseType = $SqlServerLicenseType
        }

        $UserProvidedTags = $null
        if ($HasTag -And $Tag) {
            $UserProvidedTags += @{"Tag" = $Tag }
        }

        if ($HasVMTag -And $VMTag) {
            $UserProvidedTags += @{"VMTag" = $VMTag }
        }

        if ($HasNicTag -And $NicTag) {
            $UserProvidedTags += @{"NicTag" = $NicTag }
        }

        if ($HasDiskTag -And $DiskTag) {
            $UserProvidedTags += @{"DiskTag" = $DiskTag }
        }

        foreach ($tagtype in $UserProvidedTags.Keys) {
            $IllegalCharKey = New-Object Collections.Generic.List[String]
            $ExceededLengthKey = New-Object Collections.Generic.List[String]
            $ExceededLengthValue = New-Object Collections.Generic.List[String]
            $ResourceTag = $($UserProvidedTags.Item($tagtype))

            if ($ResourceTag.Count -gt 50) {
                throw "InvalidTags : Too many tags specified. Requested tag count - '$($ResourceTag.Count)'. Maximum number of tags allowed - '50'."
            }

            foreach ($key in $ResourceTag.Keys) {
                if ($key.length -eq 0) {
                    throw "InvalidTagName : The tag name must be non-null, non-empty and non-whitespace only. Please provide an actual value."
                }

                if ($key.length -gt 512) {
                    $ExceededLengthKey.add($key)
                }

                if ($key -match "[<>%&\?/.]") {
                    $IllegalCharKey.add($key)
                }

                if ($($ResourceTag.Item($key)).length -gt 256) {
                    $ExceededLengthValue.add($($ResourceTag.Item($key)))
                }
            }

            if ($IllegalCharKey.Count -gt 0) {
                throw "InvalidTagNameCharacters : The tag names '$($IllegalCharKey -join ', ')' have reserved characters '<,>,%,&,\,?,/' or control characters."
            }

            if ($ExceededLengthKey.Count -gt 0) {
                throw "InvalidTagName : Tag key too large. Following tag name '$($ExceededLengthKey -join ', ')' exceeded the maximum length. Maximum allowed length for tag name - '512' characters."
            }

            if ($ExceededLengthValue.Count -gt 0) {
                throw "InvalidTagValueLength : Tag value too large. Following tag value '$($ExceededLengthValue -join ', ')' exceeded the maximum length. Maximum allowed length for tag value - '256' characters."
            }

            if ($tagtype -eq "Tag" -or $tagtype -eq "DiskTag") {
                $ProviderSpecificDetails.SeedDiskTag = $ResourceTag
                $ProviderSpecificDetails.TargetDiskTag = $ResourceTag
            }

            if ($tagtype -eq "Tag" -or $tagtype -eq "NicTag") {
                $ProviderSpecificDetails.TargetNicTag = $ResourceTag
            }

            if ($tagtype -eq "Tag" -or $tagtype -eq "VMTag") {
                $ProviderSpecificDetails.TargetVmTag = $ResourceTag
            }
        }

        $ProviderSpecificDetails.TargetBootDiagnosticsStorageAccountId = $TargetBootDiagnosticsStorageAccount
        $ProviderSpecificDetails.TargetNetworkId = $TargetNetworkId
        $ProviderSpecificDetails.TargetResourceGroupId = $TargetResourceGroupId
        $ProviderSpecificDetails.TargetSubnetName = $TargetSubnetName

        if ($TargetVMName.length -gt 64 -or $TargetVMName.length -eq 0) {
            throw "The target virtual machine name must be between 1 and 64 characters long."
        }

        Import-Module Az.Resources
        $TargetResourceGroupName = $ProviderSpecificDetails.TargetResourceGroupId.Split('/')[4]
        $VMNamePresentinRg = Get-AzResource -ResourceGroupName $TargetResourceGroupName -Name $TargetVMName -ResourceType "Microsoft.Compute/virtualMachines" -ErrorVariable notPresent -ErrorAction SilentlyContinue
        if ($VMNamePresentinRg) {
            throw "The target virtual machine name must be unique in the target resource group."
        }

        if ($TargetVMName -notmatch "^[^_\W][a-zA-Z0-9\-]{0,63}(?<![-._])$") {
            throw "The target virtual machine name must begin with a letter or number, and can contain only letters, numbers, or hyphens(-). The names cannot contain special characters \/""[]:|<>+=;,?*@&, whitespace, or begin with '_' or end with '.' or '-'."
        }

        $ProviderSpecificDetails.TargetVMName = $TargetVMName
        if ($HasTargetVMSize) { $ProviderSpecificDetails.TargetVMSize = $TargetVMSize }
        $ProviderSpecificDetails.VmwareMachineId = $MachineId
        $uniqueDiskUuids = [System.Collections.Generic.HashSet[String]]::new([StringComparer]::InvariantCultureIgnoreCase)

        if ($parameterSet -match 'DefaultUser') {
            [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.IVMwareCbtDiskInput[]]$DiskToInclude = @()
            foreach ($onPremDisk in $InputObject.Disk) {
                if ($onPremDisk.Uuid -ne $OSDiskID) {
                    $DiskObject = [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.VMwareCbtDiskInput]::new()
                    $DiskObject.DiskId = $onPremDisk.Uuid
                    $DiskObject.DiskType = "Standard_LRS"
                    $DiskObject.IsOSDisk = "false"
                    $DiskObject.LogStorageAccountSasSecretName = $LogStorageAccountSas
                    $DiskObject.LogStorageAccountId = $LogStorageAccountID
                    if ($HasDiskEncryptionSetID) {
                        $DiskObject.DiskEncryptionSetId = $DiskEncryptionSetID
                    }
                    $DiskToInclude += $DiskObject
                }
            }
            $DiskObject = [Microsoft.Azure.PowerShell.Cmdlets.Migrate.Models.Api20210210.VMwareCbtDiskInput]::new()
            $DiskObject.DiskId = $OSDiskID
            $DiskObject.DiskType = $DiskType
            $DiskObject.IsOSDisk = "true"
            $DiskObject.LogStorageAccountSasSecretName = $LogStorageAccountSas
            $DiskObject.LogStorageAccountId = $LogStorageAccountID
            if ($HasDiskEncryptionSetID) {
                $DiskObject.DiskEncryptionSetId = $DiskEncryptionSetID
            }

            $DiskToInclude += $DiskObject
            $ProviderSpecificDetails.DisksToInclude = $DiskToInclude
        }
        else {
            foreach ($DiskObject in $DiskToInclude) {
                $DiskObject.LogStorageAccountSasSecretName = $LogStorageAccountSas
                $DiskObject.LogStorageAccountId = $LogStorageAccountID
            }
            $ProviderSpecificDetails.DisksToInclude = $DiskToInclude
        }


        # Check for duplicate disk UUID in user input/discovered VM.
        foreach ($disk in $ProviderSpecificDetails.DisksToInclude)
        {
            if ($uniqueDiskUuids.Contains($disk.DiskId)) {
                throw "The disk uuid '$($disk.DiskId)' is already taken."
            }
            $res = $uniqueDiskUuids.Add($disk.DiskId)
        }

        $null = $PSBoundParameters.add('ProviderSpecificDetail', $ProviderSpecificDetails)
        $null = $PSBoundParameters.Add('NoWait', $true)
        $output = Az.Migrate.internal\New-AzMigrateReplicationMigrationItem @PSBoundParameters
        $JobName = $output.Target.Split("/")[12].Split("?")[0]
        $null = $PSBoundParameters.Remove('NoWait')
        $null = $PSBoundParameters.Remove('ProviderSpecificDetail')
        $null = $PSBoundParameters.Remove("ResourceGroupName")
        $null = $PSBoundParameters.Remove("ResourceName")
        $null = $PSBoundParameters.Remove("FabricName")
        $null = $PSBoundParameters.Remove("MigrationItemName")
        $null = $PSBoundParameters.Remove("ProtectionContainerName")
        $null = $PSBoundParameters.Remove("PolicyId")

        $null = $PSBoundParameters.Add('JobName', $JobName)
        $null = $PSBoundParameters.Add('ResourceName', $VaultName)
        $null = $PSBoundParameters.Add('ResourceGroupName', $ResourceGroupName)
        
        return Az.Migrate.internal\Get-AzMigrateReplicationJob @PSBoundParameters
        
    }

}   
# SIG # Begin signature block
# MIIjkgYJKoZIhvcNAQcCoIIjgzCCI38CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDk4Tb4IygFV8U/
# yxQINYeciAuJWiXf4O4tqWagGjuj06CCDYEwggX/MIID56ADAgECAhMzAAAB32vw
# LpKnSrTQAAAAAAHfMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAxMjE1MjEzMTQ1WhcNMjExMjAyMjEzMTQ1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC2uxlZEACjqfHkuFyoCwfL25ofI9DZWKt4wEj3JBQ48GPt1UsDv834CcoUUPMn
# s/6CtPoaQ4Thy/kbOOg/zJAnrJeiMQqRe2Lsdb/NSI2gXXX9lad1/yPUDOXo4GNw
# PjXq1JZi+HZV91bUr6ZjzePj1g+bepsqd/HC1XScj0fT3aAxLRykJSzExEBmU9eS
# yuOwUuq+CriudQtWGMdJU650v/KmzfM46Y6lo/MCnnpvz3zEL7PMdUdwqj/nYhGG
# 3UVILxX7tAdMbz7LN+6WOIpT1A41rwaoOVnv+8Ua94HwhjZmu1S73yeV7RZZNxoh
# EegJi9YYssXa7UZUUkCCA+KnAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUOPbML8IdkNGtCfMmVPtvI6VZ8+Mw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDYzMDA5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAnnqH
# tDyYUFaVAkvAK0eqq6nhoL95SZQu3RnpZ7tdQ89QR3++7A+4hrr7V4xxmkB5BObS
# 0YK+MALE02atjwWgPdpYQ68WdLGroJZHkbZdgERG+7tETFl3aKF4KpoSaGOskZXp
# TPnCaMo2PXoAMVMGpsQEQswimZq3IQ3nRQfBlJ0PoMMcN/+Pks8ZTL1BoPYsJpok
# t6cql59q6CypZYIwgyJ892HpttybHKg1ZtQLUlSXccRMlugPgEcNZJagPEgPYni4
# b11snjRAgf0dyQ0zI9aLXqTxWUU5pCIFiPT0b2wsxzRqCtyGqpkGM8P9GazO8eao
# mVItCYBcJSByBx/pS0cSYwBBHAZxJODUqxSXoSGDvmTfqUJXntnWkL4okok1FiCD
# Z4jpyXOQunb6egIXvkgQ7jb2uO26Ow0m8RwleDvhOMrnHsupiOPbozKroSa6paFt
# VSh89abUSooR8QdZciemmoFhcWkEwFg4spzvYNP4nIs193261WyTaRMZoceGun7G
# CT2Rl653uUj+F+g94c63AhzSq4khdL4HlFIP2ePv29smfUnHtGq6yYFDLnT0q/Y+
# Di3jwloF8EWkkHRtSuXlFUbTmwr/lDDgbpZiKhLS7CBTDj32I0L5i532+uHczw82
# oZDmYmYmIUSMbZOgS65h797rj5JJ6OkeEUJoAVwwggd6MIIFYqADAgECAgphDpDS
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
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVZzCCFWMCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAd9r8C6Sp0q00AAAAAAB3zAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgZHlSSNiq
# xZ9HJaNkIa6OPWDxtbaQ05GAJGhZUdG9i0cwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQB5uvks4Y3dxCeEektsLdrn8mI5Pu3HX8+tfjhnZjNJ
# U++yu2QnaHEUyouCwYcqQXlI4mIiOhA4RNVZHST6V8Ygj2kuTisspzJop5b/5zuV
# RJTTy9Qf2GVJPnkeaDNs3XoqxC/5dKekqlhz/C9fakEly397CSbfzvza2RCUo6FN
# fm8RvVeNreBbFDVpuMonxXd1qw/6/845m07TMV5kDVSRLdFhSoiZ7Oyoa3vceHNl
# TWAu1vjPOW3wysI0G8tOWUy1zn86b8cjFMDYIoW9roXuwmJ36suxDLNXsx+nEvqn
# YBJsDbUlkUioRK8efqIoqilhMVPMcJOcnjBjK+IcVqJqoYIS8TCCEu0GCisGAQQB
# gjcDAwExghLdMIIS2QYJKoZIhvcNAQcCoIISyjCCEsYCAQMxDzANBglghkgBZQME
# AgEFADCCAVUGCyqGSIb3DQEJEAEEoIIBRASCAUAwggE8AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEICXEsmq9pUx3QVXdNyRQq0ssk6HlE2L4YgsY7k3t
# VRG+AgZhebGw6gIYEzIwMjExMDI5MDc1NTI2LjQ2NlowBIACAfSggdSkgdEwgc4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAsTIE1p
# Y3Jvc29mdCBPcGVyYXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMg
# VFNTIEVTTjpDNEJELUUzN0YtNUZGQzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgU2VydmljZaCCDkQwggT1MIID3aADAgECAhMzAAABV0QHYtxv6L4qAAAA
# AAFXMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEw
# MB4XDTIxMDExNDE5MDIxM1oXDTIyMDQxMTE5MDIxM1owgc4xCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVy
# YXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpDNEJE
# LUUzN0YtNUZGQzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAN5tA6dUZvnnwL9qQtXc
# wPANhB4ez+5CQrePp/Z8TH4NBr5vAfGMo0lV/lidBatKTgHErOuKH11xVAfBehHJ
# vH9T/OhOc83CJs9bzDhrld0Jdy3eJyC0yBdxVeucS+2a2ZBd50wBg/5/2YjQ2ylf
# D0dxKK6tQLxdODTuadQMbda05lPGnWGwZ3niSgIKVRgqqCVlhHzwNtRh1AH+Zxbf
# Se7t8z3oEKAdTAy7SsP8ykht3srjdh0BykPFdpaAgqwWCJJJmGk0gArSvHC8+vXt
# Go3MJhWQRe5JtzdD5kdaKH9uc9gnShsXyDEhGZjx3+b8cuqEO8bHv0WPX9MREfrf
# xvkCAwEAAaOCARswggEXMB0GA1UdDgQWBBRdMXu76DghnU/kPTMKdFkR9oCp2TAf
# BgNVHSMEGDAWgBTVYzpcijGQ80N7fEYbxTNoWoVtVTBWBgNVHR8ETzBNMEugSaBH
# hkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNU
# aW1TdGFQQ0FfMjAxMC0wNy0wMS5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUF
# BzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1RpbVN0
# YVBDQV8yMDEwLTA3LTAxLmNydDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsG
# AQUFBwMIMA0GCSqGSIb3DQEBCwUAA4IBAQAld3kAgG6XWiZyvdibLRmWr7yb6RSy
# cjVDg8tcCitS01sTVp4T8Ad2QeYfJWfK6DMEk7QRBfKgdN7oE8dXtmQVL+JcxLj0
# pUuy4NB5RchcteD5dRnTfKlRi8vgKUaxDcoFIzNEUz1EHpopeagDb4/uI9Uj5tIu
# wlik/qrv/sHAw7kM4gELLNOgdev9Z/7xo1JIwfe0eoQM3wxcCFLuf8S9OncttaFA
# WHtEER8IvgRAgLJ/WnluFz68+hrDfRyX/qqWSPIE0voE6qFx1z8UvLwKpm65QNyN
# DRMp/VmCpqRZrxB1o0RY7P+n4jSNGvbk2bR70kKt/dogFFRBHVVuUxf+MIIGcTCC
# BFmgAwIBAgIKYQmBKgAAAAAAAjANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJv
# b3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMTAwNzAxMjEzNjU1WhcN
# MjUwNzAxMjE0NjU1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKkdDbx3EYo6IOz8E5f1+n9plGt0
# VBDVpQoAgoX77XxoSyxfxcPlYcJ2tz5mK1vwFVMnBDEfQRsalR3OCROOfGEwWbEw
# RA/xYIiEVEMM1024OAizQt2TrNZzMFcmgqNFDdDq9UeBzb8kYDJYYEbyWEeGMoQe
# dGFnkV+BVLHPk0ySwcSmXdFhE24oxhr5hoC732H8RsEnHSRnEnIaIYqvS2SJUGKx
# Xf13Hz3wV3WsvYpCTUBR0Q+cBj5nf/VmwAOWRH7v0Ev9buWayrGo8noqCjHw2k4G
# kbaICDXoeByw6ZnNPOcvRLqn9NxkvaQBwSAJk3jN/LzAyURdXhacAQVPIk0CAwEA
# AaOCAeYwggHiMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBTVYzpcijGQ80N7
# fEYbxTNoWoVtVTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDCBoAYDVR0g
# AQH/BIGVMIGSMIGPBgkrBgEEAYI3LgMwgYEwPQYIKwYBBQUHAgEWMWh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9QS0kvZG9jcy9DUFMvZGVmYXVsdC5odG0wQAYIKwYB
# BQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AUABvAGwAaQBjAHkAXwBTAHQAYQB0AGUA
# bQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAAfmiFEN4sbgmD+BcQM9naOh
# IW+z66bM9TG+zwXiqf76V20ZMLPCxWbJat/15/B4vceoniXj+bzta1RXCCtRgkQS
# +7lTjMz0YBKKdsxAQEGb3FwX/1z5Xhc1mCRWS3TvQhDIr79/xn/yN31aPxzymXlK
# kVIArzgPF/UveYFl2am1a+THzvbKegBvSzBEJCI8z+0DpZaPWSm8tv0E4XCfMkon
# /VWvL/625Y4zu2JfmttXQOnxzplmkIz/amJ/3cVKC5Em4jnsGUpxY517IW3DnKOi
# PPp/fZZqkHimbdLhnPkd/DjYlPTGpQqWhqS9nhquBEKDuLWAmyI4ILUl5WTs9/S/
# fmNZJQ96LjlXdqJxqgaKD4kWumGnEcua2A5HmoDF0M2n0O99g/DhO3EJ3110mCII
# YdqwUB5vvfHhAN/nMQekkzr3ZUd46PioSKv33nJ+YWtvd6mBy6cJrDm77MbL2IK0
# cs0d9LiFAR6A+xuJKlQ5slvayA1VmXqHczsI5pgt6o3gMy4SKfXAL1QnIffIrE7a
# KLixqduWsqdCosnPGUFN4Ib5KpqjEWYw07t0MkvfY3v1mYovG8chr1m1rtxEPJdQ
# cdeh0sVV42neV8HR3jDA/czmTfsNv11P6Z0eGTgvvM9YBS7vDaBQNdrvCScc1bN+
# NR4Iuto229Nfj950iEkSoYIC0jCCAjsCAQEwgfyhgdSkgdEwgc4xCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAsTIE1pY3Jvc29mdCBP
# cGVyYXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpD
# NEJELUUzN0YtNUZGQzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
# dmljZaIjCgEBMAcGBSsOAwIaAxUAES34SWJ7DfbSG/gbIQwTrzgZ8PKggYMwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIF
# AOUlgZwwIhgPMjAyMTEwMjkwMDA3NTZaGA8yMDIxMTAzMDAwMDc1NlowdzA9Bgor
# BgEEAYRZCgQBMS8wLTAKAgUA5SWBnAIBADAKAgEAAgITCQIB/zAHAgEAAgIRUTAK
# AgUA5SbTHAIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIB
# AAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBAGNs9RhP/GDFKt3/
# Pj9tBkR494b0b6ZBI9rO4Zp08N3R5noUwgxHQDhclt0GTWOXtkEYcYguiJRb5U1w
# D23FqkcEZvCSA+5u1YufTYC+Jb8e0qWGADpgHGkmgyk4Sm4ic+IGSmP68KpQbk3+
# AE9mtP9hn9MxaT0eHmZC23hVZN7UMYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAFXRAdi3G/ovioAAAAAAVcwDQYJYIZIAWUD
# BAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0B
# CQQxIgQg821IK+8vsnX54hRiq74mc6WD92L3GqcHQNyRl9V07TkwgfoGCyqGSIb3
# DQEJEAIvMYHqMIHnMIHkMIG9BCAsWo0NQ6vzuuupUsZEMSJ4UsRjtQw2dFxZWkHt
# qRygEzCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB
# V0QHYtxv6L4qAAAAAAFXMCIEIDis/WP3owlzX53oIJbfKOB49AR8UCCcoxbSQxMa
# dj6zMA0GCSqGSIb3DQEBCwUABIIBADbLGPAks6m139IoW3PHXYMR60ZgzID6aqnZ
# QxaG77hHNTP8Ij9GHRM/HQhoNGAJUjIqchoZdUOev2qQUC3kJ4On9UoIqsX7JfGe
# l8M/M8HXEv2dfnVB9yBQWV9ZeXBhC8lwCaELnq7cOMD6YO9SelUuAXgwTF/41tm+
# 6eJcYcajzSFy/yXrX9W83QqWI9TMLJ5ILPLvV0Bp6SCox7WSE7Eui0qojcQfHefx
# cuqL75NJ9cLzOdA/r7wNP/EgUIgeVRG6P4+eBQVlI9V5O/4KKTEbeNhp8RWXHl0a
# gFpQM5hTDGDfVhIFkV5Kh+T4khMSN67n3bcwsQ2KBFX9nBTAy1Y=
# SIG # End signature block
