Param (
    
    [Parameter( Mandatory=$true)] [String]$BDOC,
    [Parameter( Mandatory=$true)] [String]$Chave

)


#Importando Módulos necessários
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Import-Module PnP.PowerShell -RequiredVersion 1.12.0
Import-Module ActiveDirectory



#Definindo Constantes de acesso ao Sharepoint

#$tenantUrl = "http
#$AdminUrl = 

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
$LogFile = "D:\Util\BibliotecaDocumentos\Logs\" + (Get-Date).ToString('yyyyMMdd') + "_BDOC-CadastroSubstituto.csv"

# -------------------- Função de log  -------------------- #

function Log([string]$Message, [bool]$Print){
    $Datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')

    if(($Message -like "ERRO") -and ($Print)){

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

#--------------------------------------------------------------------------------------------------#

Function RegisterUserAsSubstitute( $BDOC, $User, $Connection){

    $Group = Get-PnPGroup -AssociatedOwnerGroup

    if($Group){

        try{
            
            # UPN para teste em DEV
            # $UPN = "andre.daltro@petrobrasbrteste.petrobras.com.br"

            Add-PnPGroupMember -LoginName $User.UserPrincipalName -Group $Group.Id -ErrorAction Stop
            Log -Message "SUCESSO: Agora o usuário $($User.UserPrincipalName) esta cadastrado como Gerente Substituto na Gerência $BDOC" -Print:$true

        } catch {

            Log -Message "ERRO: Ocorreu um erro ao inserir o usuário no grupo de Gerência da BDOC $BDOC" -Print:$true
            exit 1
        }
    
    } else {

        Log -Message "*ERRO*: Ocorreu um erro ao resgatar o grupo de gerência da BDOC $BDOC" -print:$true
        exit 1
    
    }

}

#--------------------------------------------------------------------------------------------------#


# Capturando usuário no AD
$User = Get-ADUser -Filter "Name -eq '$Chave'"

if($User){

    RegisterUserAsSubstitute -BDOC $BDOC -User $User 

} else {
    
    Log -Message "*ERRO*: Não foi Possível encontrar o Usuário $Chave no Active Directory" -Print:$true
    exit 1
}


