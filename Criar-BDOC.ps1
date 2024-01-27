

param (
    [Parameter(Mandatory=$true)] [string]$Gerencia
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Import-Module PnP.PowerShell -RequiredVersion 1.12.0
Import-Module ActiveDirectory

# Definindo constantes
$LogFile = "D:\Util\BibliotecaDocumentos\Logs\"+(Get-Date).ToString('yyyyMMdd')+"_BDOC-Criacao.csv"
$Sleep = 5



$tenantUrl = "https://petrobrasbrteste.sharepoint.com/teams/"
$hubUrl = "https://petrobrasbrteste.sharepoint.com/teams/bdoc"
$centralAdminUrl = "https://petrobrasbrteste-admin.sharepoint.com"
#$centralAdminUrl = "https://petrobras-admin.sharepoint.com"


# Definindo as credenciais de acesso 
$AdminEmail = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
$PlainPassword = "Ror66406"
$SecurePassword = ConvertTo-SecureString -String $PlainPassword -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminEmail, $SecurePassword


# ----------------------- Inicio das Fun��es Auxiliares -------------------- #

#Fun��o de Log
function Log([string]$Message, [bool]$Print){
    $Datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')

    if(($Message.contains("ERRO")) -and ($Print)){

        Write-Host "$Datetime;$Message" -ForegroundColor Red
    }
    elseif(($Message -like "ATENCAO") -and ($Print)){

        Write-Host "$Datetime;$Message" -ForegroundColor Yellow
    } elseif ($Print){

        Write-Host "$Datetime;$Message" -ForegroundColor Green
    }

    Add-Content -Path $LogFile -Value "$Datetime;$Message"
}

# ------------------------ Fim das Fun��es Auxiliares ---------------------- #

 
# ------------------------- Conectando ao SharePoint ----------------------- #
try{
    Connect-PnPOnline -Url $CentralAdminUrl -Credentials $Credential -InformationAction SilentlyContinue | Out-Null
    Log "SUCESSO: ao Conectar com o SharePoint 1/3"
} catch {
    Log "ATENCAO: ao conectar com o SharePoint 1/3"
    Start-Sleep $SLEEP
    try{
         Connect-PnPOnline -Url $CentralAdminUrl -Credential $Credential -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
        Log "SUCESSO: ao Conectar com o SharePoint 2/3"
    }catch{
        Log "ATENCAO: ao conectar com o SharePoint 2/3"
        Start-Sleep $SLEEP
        try {
             Connect-PnPOnline -Url $CentralAdminUrl -Credential $Credential -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
            Log "SUCESSO: ao Conectar com o SharePoint 3/3"
        } catch{
            Log "ERRO: ao conectar com o SharePoint 3/3"
            exit 1
        }
    }
#}
# ------------------- Fim da Conex�o com o SharePoint -------------------- #

 

# ------------------- Inicio das Fun��es de processamento ------------------ #

Function CreateBDOC(){

    # Capturando os membros da Gerencia
    $DeptoMembers = Get-ADUser -Filter "Department -eq '$Gerencia'" -SearchBase "OU=Usuarios,DC=petrobras,DC=biz" -Properties extensionAttribute1,extensionAttribute2 | Select-Object UserPrincipalName,extensionAttribute1,extensionAttribute2

    # Criando um array para armazenar os membros da gerencia
    $ArrayMembers = [System.Collections.ArrayList]::new()

    # # Inicializando as variáveis do Gerente
    $Manager = $null
    $ManagerRole = $null

    # Percorrendo os membros do departamento para separar os membros e gerentes
    foreach($Member in $DeptoMembers){

        # Capturando as propriedades do membro
        $upn = $Member.UserPrincipalName
        $ext1 = $Member.extensionAttribute1
        $ext2 = $Member.extensionAttribute2

        # Verificando a role do Membro
        if($ext1 -like "Gerente*" -or $ext1 -in "Advogado(a)-Geral da Petrobras","Ouvidor(a)-Geral","Secretario(a) da Presidencia","Chefe do Gabinete da Presidencia","Diretor(a)","Presidente"){

            # Setando as propriedades do Gerente
            $Manager = $upn

        } elseif($ext2 -in "F","N","T"){

            $ArrayMembers.Add($upn)
        }    
    }


    $datetime = (Get-Date).ToString('yyyyMMddHHmmss')
    $siteUrl = "bdoc_$datetime"


    # Criando constantes para prosseguimento do script
    $siteUrl = "bdoc_$GerenciaLinkName"
    $PageName = "Default.aspx"
    

    # Criando o Site da BDOC
    New-PnPTenantSite -Url ($tenantUrl + $siteUrl) -Owner $adminEmail -Title $Gerencia -Template "STS#3" -TimeZone 8 -RemoveDeletedSite -Lcid 1046 -StorageQuota 100000 -ResourceQuotaWarningLevel 98 -StorageQuotaWarningLevel 98 -ResourceQuota 300 -Wait:$true -ErrorAction Stop

    # Associando o Site ao HubSite
    Add-PnPHubSiteAssociation -Site ($tenantUrl + $siteUrl) -HubSite $hubUrl

    # Conectando ao HubSite
    Connect-PnPOnline ($tenantUrl + $siteUrl) -Credentials $Credential -ErrorAction Stop

    # Adicionando os usuarios como administradores da collection do site
    Add-PnPSiteCollectionAdmin -Owners $adminEmail

    # Criando a p�gina
    Add-PnPPage -Name $pageName -LayoutType Home -CommentsEnabled:$false -ErrorAction SilentlyContinue

    Start-Sleep 5

    # Adicionando uma nova se��o a p�gina
    Add-PnPPageSection -Page $pageName -SectionTemplate:OneColumn -Order 0 -ErrorAction SilentlyContinue

    # Definindo a P�gina Inicial
    Set-PnPHomePage -RootFolderRelativeUrl ("Documentos Compartilhados/Forms/AllItems.aspx")

    # Definindo o Logo da P�gina
    Set-PnPWeb -SiteLogoUrl "https://petrobrasbr.sharepoint.com/teams/bdoc/SiteAssets/__siteIcon__.png"

    Start-Sleep 5

    # Recuperando a lista de Documentos
    $list = Get-PnPList -Identity "Documentos"

    # Adicionando uma Web Part
    Add-PnPPageWebPart -Page "Default" -DefaultWebPartType "List" -WebPartProperties @{isDocumentLibrary="true";selectedListId=$list.Id}
    $web = Get-PnPWeb
    $web.QuickLaunchEnabled = $false
    $web.Update()
    Invoke-PnPQuery

    Start-Sleep 5

    # Publicando a Home Page
    Set-PnPClientSidePage -Identity $PageName -Publish:$true -LayoutType:Home

    # Configurando as permiss�es de Propriet�rio
    Add-PnPRoleDefinition -RoleName "Propriet�rios" -Include Open,ViewListItems,ApproveItems,OpenItems,ViewVersions,ViewPages,ManagePermissions,EnumeratePermissions,ApproveItems,BrowseDirectories,BrowseUserInfo | Out-Null

    # Associando os grupos as fun��es no site
    $web = Get-PnPWeb
    Set-PnPGroupPermissions -Identity $web.AssociatedOwnerGroup -RemoveRole 'Controle Total' -AddRole 'Propriet�rios'
    Set-PnPGroupPermissions -Identity $web.AssociatedMemberGroup -RemoveRole 'Editar' -AddRole 'Leitura'

    # Configurando permissoes da biblioteca de documentos
    # Orientacao Microsoft: Manter o n�vel de permissionamento de Proprietarios na lista
    Set-PnPList -Identity "Documentos" -BreakRoleInheritance:$true -CopyRoleAssignments:$false | Out-Null

    Set-PnPListPermission -Identity "Documentos" -Group $web.AssociatedOwnerGroup -AddRole "Propriet�rios"
    Set-PnPListPermission -Identity "Documentos" -Group $web.AssociatedOwnerGroup -AddRole "Colabora��o"
    Set-PnPListPermission -Identity "Documentos" -Group $web.AssociatedOwnerGroup -AddRole "Design"
    Set-PnPListPermission -Identity "Documentos" -Group $web.AssociatedMemberGroup -AddRole "Colabora��o"
    Set-PnPListPermission -Identity "Documentos" -Group $web.AssociatedVisitorGroup -AddRole "Leitura"

    # Desabilitando comentarios nas paginas"
    $adminConn = Connect-PnPOnline $centralAdminUrl -Credentials $Credential -ReturnConnection
    $spfxConn = Connect-PnPOnline ($tenantUrl + $siteUrl) -Credentials $Credential -ReturnConnection
    Set-PnPSite -CommentsOnSitePagesDisabled:$true -Connection $adminConn -Identity ($tenantUrl + $siteUrl)

    # Configurando o limite de versoes da Biblioteca de Documentos para 100
    Set-PnPList -Identity "Documentos" -EnableVersioning 1 -EnableMinorVersions 0 -MajorVersions 100 -MinorVersions 0 | Out-Null

    # Pastas criadas por padrao:
    # Ger�ncia: Proprietarios tem acesso de edicao, membros e visitantes nao possuem acesso
    # Conectando novamente ao Site
    Connect-PnPOnline ($tenantUrl + $siteUrl) -Credentials $Credential

    # Adicionando pasta padr�o na biblioteca
    Add-PnPFolder -Name "Ger�ncia" -Folder "Documentos Compartilhados" | Out-Null
    $members = Get-PnPGroup -AssociatedMemberGroup
    $visitors = Get-PnPGroup -AssociatedVisitorGroup
    $folder = Get-PnPFolder -Url "Documentos Compartilhados\Ger�ncia"
    Set-PnPFolderPermission -List "Documentos" -Identity $folder -Group $visitors.Title -RemoveRole "Leitura"
    Set-PnPFolderPermission -List "Documentos" -Identity $folder -Group $members.Title -RemoveRole "Colabora��o"


    # Adicionando as colunas Tamanho e Confidencialidade
    $view = Get-PnPView -List "Documentos" -Identity "Todos os Documentos"
    if($null -ne $view)
    {
        Set-PnPView -List "Documentos" -Identity $view.Id -Fields "DocIcon","LinkFilename","Modified","Editor","FileSizeDisplay","Confidencialidade" | Out-Null
    }
    else
    {
        Log "ERRO;View Todos os Documentos n�o encontrada"
    }

    # Niveis de compartilhamento:
    # Proprietario:	Pode compartilhar com pessoas de outras areas internas a Petrobras.
    # Membro: Pode compartilhar com pessoas de outras areas internas a Petrobras.
    # Visitante: Pode compartilhar com pessoas de outras areas internas a Petrobras, porem e necessario que um proprietario aprove este compartilhamento
    #Log "Configurando niveis de compartilhamento"
    Set-PnPSite -Connection $adminConn -Identity ($tenantUrl + $siteUrl) -Sharing:Disabled -DefaultLinkPermission:None -DefaultSharingLinkType:None -DisableCompanyWideSharingLinks:Disabled -SocialBarOnSitePagesDisabled:$true

    Start-Sleep -Seconds 5

    if ($Manager) {    

        Log "INFO;Propriet�rio;$tenantUrl$siteUrl;$Manager"
        $group = Get-PnPGroup -AssociatedOwnerGroup
        Add-PnPGroupMember -LoginName $Manager -Group $group.Id -Connection $spfxConn
    }

    if ($ArrayMembers) {

        $c = $ArrayMembers.Count
        Log "INFO;Membros;$tenantUrl$siteUrl;$c"
        $group = Get-PnPGroup -AssociatedMemberGroup

        foreach($membro in $ArrayMembers) {

            Add-PnPGroupMember -LoginName $membro -Group $group.Id -Connection $spfxConn -ErrorAction SilentlyContinue 

        }
    }
}

# -------------------- Fim das Fun��es de processamento -------------------- #