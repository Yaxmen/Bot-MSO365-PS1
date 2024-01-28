lParam (
    
    [Parameter( Mandatory=$true)] [String]$BDOC,
    [Parameter( Mandatory=$true)] [String]$Chaves

)


#Importando Módulos necessários
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Import-Module PnP.PowerShell -RequiredVersion 1.12.0
Import-Module ActiveDirectory




#Definindo Constantes de acesso ao Sharepoint

$tenantUrl = "https://petrobrasbr.sharepoint.com/teams/"
#$tenantUrl = "https://petrobrasbrteste.sharepoint.com/teams/"
#$AdminUrl = "https://petrobrasbrteste-admin.sharepoint.com/"

#Definindo Constantes de credenciais

# Credenciais de DEV
<#$UserName = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
$PlainPassword = "Ror66406"
$SecurePassword = ConvertTo-SecureString -String $PlainPassword -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword #>

# Credenciais de PROD
$username = "email"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop 


#Definindo o caminho do log
$LogFile = "D:\Util\BibliotecaDocumentos\Logs\" + (Get-Date).ToString('yyyyMMdd') + "_BDOC-Concessao.csv"

# -------------------- Função de log  -------------------- #

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

# --------------- Conectando ao SharePoint --------------- #




# Tratando o nome da BDOC

<#$GerenciaLinkName = $BDOC.replace('/','-').replace('.','')
$siteUrl = "bdoc_$GerenciaLinkName"#>


$CsvFile = Import-Csv -Path "D:\Util\BibliotecaDocumentos\BDOC-URL.csv" -Delimiter ";" | Select-Object -Property "Title","Url"

foreach($BDOC in $CsvFile){

    if($BDOC.'Title'.Equals($BDOC)){

        $Url = $BDOC.'Url'
    
    }
}



try{
    Connect-PnPOnline -Url $Url -Credentials $Credential -ErrorAction Stop
    Log -Message "Sucesso ao Conectar com o SharePoint Online 1/3" -Print:$false
} catch{
    Log -Message "ATENCAO ao contectar ao SharePoint Online 1/3" -Print:$false
    try {
        Connect-PnPOnline -Url $Url -Credentials $Credential -ErrorAction Stop
        Log -Message "Sucesso ao Conectar com o SharePoint Online 2/3" -Print:$false
    } catch {
        Log -Message "ATENCAO ao contectar ao SharePoint Online 2/3" -Print:$false

        try {
            Connect-PnPOnline -Url $Url -Credentials $Credential -ErrorAction Stop
            Log -Message "Sucesso ao Conectar com o SharePoint Online 3/3" -Print:$false
        } catch {
            Log -Message "ERRO: ao se conectar ao SharePoint Online" -Print:$true
            exit 1
        }
    }
} 



# ------------------ Funções Principais ------------------ #
function GiveAcces($BDOC, $Chaves){
    
    # Capturando o grupo de membros da BDOC
    $Group = Get-PnPGroup -AssociatedMemberGroup


    # Criando um array de Sucesso e Falha
    $SuccessArray = [System.Collections.ArrayList]::new()
    $FailedArray = [System.Collections.ArrayList]::new()
    $NotFoundArray = [System.Collections.ArrayList]::new()

    # Percorrendo a lista de chaves a serem adicionadas na BDOC
    foreach($User in $Chaves){
        
        # Capturando a mailbox do Usuário a ser adicionado na BDOC
        $ADUser = Get-ADUser -Filter "Name -eq '$User'"

        # Verificando se o usuário foi localizado
        if($ADUser){
            
            try{
                
                
                Add-PnPGroupMember -LoginName $ADUser.UserPrincipalName -Group $Group.Id -ErrorAction Stop
                $SuccessArray.Add( $ADUser.Name) | Out-Null

            }catch{
                
                $FailedArray.Add($ADUser.Name) | Out-Null
            }
        } else {

            $NotFoundArray.Add($User) | Out-Null
        
        }
    }

    # Verificando os casos de falha e Sucesso
    if($SuccessArray.Length -gt 0){

        Log "SUCESSO: Os seguintes usuários foram adicionados a BDOC" -Print:$true

        foreach($User in $SuccessArray){
            Log " -$User" -Print:$true
        }

        # Verificando se houveram falhas durante a execução
        if($NotFoundArray.length -gt 0){
            
            Log "Os seguintes usuários não foram adicionados por não serem localizados no Exchange:" -Print:$true
            foreach($User in $NotFoundArray){
                Log " -$User" -Print:$true
            }

        }

        if($FailedArray.Length -gt 0){
            
            Log "Os seguintes usuários não foram adicionados por um erro interno na execução do Comando:" -Print:$true
            foreach($User in $NotFoundArray){
                Log " -$User" -Print:$true
            }
        } 

    } else {
        
        # Caso não ocorra nenhum caso de sucesso
        Log "ERRO: Não houve nenhum caso de sucesso na execução" -print:$true

        if($NotFoundArray.length -gt 0){
            
            Log "Os seguintes usuários não foram adicionados por não serem localizados no Exchange:" -Print:$true
            foreach($User in $NotFoundArray){
                Log " -$User" -Print:$true
            }

        }

        if($FailedArray.Length -gt 0){
            
            Log "Os seguintes usuários não foram adicionados por um erro interno na execução do Comando :" -Print:$true
            foreach($User in $FailedArray){
                Log " -$User" -Print:$true
            }
        }
        
        exit 1 
    }
}

# --------------------- Processamento -------------------- #

$ArrayChaves =  $Chaves.split(",")

# Tornando a Sala "Unlock"
Set-PnPSite -LockState Unlock

# Realizando as alterações
GiveAcces -BDOC $BDOC -Chaves $ArrayChaves

# Tornando a Sala "ReadOnly"
Set-PnPSite -LockState ReadOnly

# Fechando a conexão com o SharePoint
Disconnect-PnPOnline
