Param (

    [Parameter( Mandatory=$true)] [String]$AffectedUser,
    [Parameter( Mandatory=$true)] [String]$Action, 
    [Parameter( Mandatory=$true)] [String]$NameGroup,
    [Parameter( Mandatory=$true)] [String]$NameGroup2,
    [Parameter( Mandatory=$true)] [String]$NameGroup3,
    [Parameter( Mandatory=$true)] [String]$Members,
    [Parameter( Mandatory=$true)] [String]$Owners,
    [Parameter( Mandatory=$true)] [String]$GroupId,
    [Parameter( Mandatory=$true)] [String]$OldGroupName,
    [Parameter( Mandatory=$true)] [String]$NewGroupName

)

# Definindo Constantes
$Sleep = 20
# DEV ACCOUNT $ServiceAccountId = "9311cc06-c0fc-45ea-ab17-350da97dab0e"
$ServiceAccountId = "db94b0cd-ec5c-423f-9c60-07811c3801d9"


# Definindo arquivo de caminho de Log
$logFile = "d:\Util\Groups\logs\"+(Get-Date).ToString('yyyy')+"_groups_log.csv"


# Definindo as credenciais de acesso a Tenant de TESTE
#$username = email
#$PlainPassword="senha"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword


# Definindo credenciais de acesso a tenant
$username = "email"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop


# ----------------------- Inicio das Funções Auxiliares -------------------- #


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

# Função responsável por verificar se algum dos nomes fornecidos está disponível na Tenant
Function VerifyName($Names){
    foreach($Name in $Names){
        $GroupExists = Get-Group -Identity $Name -ErrorAction SilentlyContinue
        if(!$GroupExists){
            return $Name
        }
    }
}

# Função responsável por remover a conta de serviço da lista de membros e proprietários
Function RemoveServiceAccountFromGroup($Group, $ServiceAccountId){
    
    # Removendo a conta da lista de proprietários
    try{
        Remove-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Owner" -Links $ServiceAccountId -Confirm:$false -ErrorAction stop | Out-Null

    } catch {
        Log "*ERRO*: ao remover a conta de serviço da lista de proprietários" -Print:$true
    
    }

    # Removendo a conta da lista de membros
    try{
        Remove-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Member" -Links $ServiceAccountId -Confirm:$false -ErrorAction stop | Out-Null

    } catch {
        Log "*ERRO*: ao remover a conta de serviço da lista de membros" -Print:$true
    
    }
}

# Função responsável por verificar se o ´usuário que abriu o chamado é owner do grupo 
Function VerifyIfAffectedUserIsOwner($Group){


    # Capturando a caixa de Email do AffectedUser
    $AffectedUserMailBox = Get-Mailbox -Identity $AffectedUser

    if($AffectedUserMailBox){

        $CurrentOwners = Get-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Owner"

        foreach ($Owner in $CurrentOwners){

            if($Owner.name -eq $AffectedUserMailBox.Alias){

                return $true
            
            }
        }
    }

    return $null

}

# ------------------------ Fim das Funções Auxiliares ---------------------- #




# -------------------------- Conectando ao Exchange ------------------------ #
try{
    Connect-ExchangeOnline -Credential $Credential -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
    Log "SUCESSO: ao Conectar com o Exchange 1/3" -Print:$false
} catch {
    Log "ATENCAO: ao conectar com o Exchange 1/3" -Print:$false
    Start-Sleep $SLEEP
    try{
        Connect-ExchangeOnline -Credential $Credential -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
        Log "SUCESSO: ao Conectar com o Exchange 2/3" -print:$false
    }catch{
        Log "ATENCAO: ao conectar com o Exchange 2/3" -Print:$false
        Start-Sleep $SLEEP
        try {
            Connect-ExchangeOnline -Credential $Credential -InformationAction SilentlyContinue -ShowBanner:$false  | Out-Null
            Log "SUCESSO: ao Conectar com o Exchange 3/3" -Print:$false
        } catch{
            Log "ERRO: ao conectar com o Exchange 3/3" -Print:$true
            exit 1
        }
    }
}
# -------------------- Fim da Conexão com o Exchange --------------------- #





# ----------------- Inicio das Funções de Processamento ------------------ #

# Função responsável por inserir uma lista de usuários como membros em um grupo.
Function InsertMembers ($Group, $Members){
    
    # Separando os Membros em um array.
    $Members = $Members.Split(",")

    # Resgatando os Membros atuais do grupo
    $CurrentMembers = Get-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Member" | Out-Null

    # Percorrendo o Array de membros para realizar as operações
    foreach ($Member in $Members){
            
        # Resgatando a caixa de correio do membro
        $MemberMailBox = Get-Mailbox -Identity $Member -ErrorAction SilentlyContinue
        
        # Verificando se o usuário existe no exchange
        if($null -eq $MemberMailBox){
            
            Log "*ATENCAO*: $Member não localizado no Exchange e não inserido na lista de Membros" -Print:$true

        } else {

            if($null -eq $CurrentMembers){
                
                try{
                        # Inserindo os Membros no Grupo 
                        Add-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Member" -Links $Member | Out-Null
                    } catch {
                        Log "*ERRO*: Ao Inserir $Member como membro no Grupo" -Print:$true
           
                }
            } else {

                # Percorrendo os membros atuais para verificar se o usuário já não é membro do grupo
                foreach($CurrentMember in $CurrentMembers){
            
                    # Verificando se o Membro já está no grupo
                    if($CurrentMember -ne $MemberMailBox.Alias){

                        try{
                            # Inserindo os Membros no Grupo 
                            Add-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Member" -Links $Member | Out-Null
                        } catch {
                            Log "*ERRO*: Ao Inserir $Member como membro no Grupo" -Print:$true
           
                        }

                    } else {

                        # Retornando mensagem caso o membro não seja incluso por já ser membro do grupo
                        Log "*ATENÇÃO*: $Member já é membro do grupo $($Group.DisplayName)" -Print:$true
                    }
                }
            }
        }           
    }
}

# Função responsável por inserir uma lista de usuários como Owner em um grupo. 
Function InsertOwners ($Group, $Owners){
    
    # Separando os Membros em um array.
    $Owners = $Owners.Split(",")

    # Resgatando os Membros atuais do grupo
    $CurrentOwners = Get-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Owner"

    # Percorrendo o Array de membros para realizar as operações
    foreach ($Owner in $Owners){
        
        # Percorrendo os membros atuais para verificar se o usuário já não é membro do grupo
        foreach($CurrentOwner in $CurrentOwners){
            
            # Verificando se o Membro já está no grupo
            if($CurrentOwner -ne $Owner){

                try{
                    # Inserindo os Membros no Grupo
                    Add-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Owner" -Links $Owner -ErrorAction Stop | Out-Null
                } catch {
                    Log "*ERRO*: Ao Inserir os proprietários no Grupo" -Print:$true
                    exit 1
                }

            } else {

                # Retornando mensagem caso o membro não seja incluso por já ser membro do grupo
                Log "*ATENÇÃO*: $Owner já é membro do grupo $($Group.DisplayName)" -Print:$true
            }
        }        
    }
}

# Função responsável por realizar o atendimento da demanda de Criação do Grupo
Function CreateGroup($Name, $Members, $Owners){
    


    #Cria o Grupo
    try {
        New-UnifiedGroup -DisplayName $Name -Alias $Name | Out-Null #Verificar o Alias
    }
    catch {
        Log "*ERRO*: ao criar o grupo" -Print:$true
        exit 1
    }
    
    # Inicializando variável de controle
    $NOK = $true

    # Inicializando loop de verificação de existencia do grupo
    while($NOK){
        
        # Inicia o Sleep de 20 segundos
        Start-Sleep $Sleep

        # Realiza tantaiva de capturar o grupo
        $Group = Get-UnifiedGroup -Identity $Name

        # Verifica se o Grupo conseguiu ser capturado
        if($null -eq $Group){

            $Sleept = $Sleep + $Sleept

            # Para o código caso não localize a caixa em até 60 segundos
            if($Sleept -gt 59){
                Log "*ERRO*: Grupo não localizado no Exchange " -Print:$true
                exit 1
            }
        } else {
            # Encerra o loop se localizar a sala
            $NOK = $false
        }
    }
    

    # Inserindo os Membros do Grupo
    InsertMembers -Group $Group -Members $Members | Out-Null

    Start-sleep $Sleep

    # Inserindo os Proprietários como Membros do Grupo
    <# IMPORTANTE: de acordo com as regras do exchange para grupos, um proprietário só pode ser 
    adicionado ao grupo se antes ele for membro do mesmo grupo, desta forma se faz necessário a
    chamada da função de inserção de membros passando como parâmetro a lista de owners, fazendo 
    com que os owners atinjam a condição de membros antes de owners #>
    InsertMembers -Group $Group -Members $Owners | Out-Null

    # Inserindo os Proprietários no Grupo
    InsertOwners -Group $Group -Owners $Owners | Out-Null

    # Aguardando os membros serem inseridos para garantir que não havera erros na remoção da conta de serviço
    Start-Sleep $Sleep

    # Removendo a conta de serviço da lista de membros e proprietários
    RemoveServiceAccountFromGroup -Group $Group -ServiceAccountId $ServiceAccountId

    # Colocando a criação da sala nos Logs
    Log "SUCESSO: Grupo $($Group.DisplayName) foi criado e populado com sucesso" -print:$true
}

# Função responsável por remover uma lista de usuários da lista de membros do grupo
Function RemoveMembers($Group, $Members){

    # Criando Array de Membros
    $Members = $Members.Split(",")

    # Resgatando lista de Owners
    $CurrentOwners = Get-UnifiedGroupLinks -Identity $Group.Alias -LinkType "Owner"

    # Percorrendo Array para realizar a operação de exclusão
    foreach($Member in $Members){
        
        # Percorrendo a lista de owners para conferir se o membro não é um owner
        foreach($Owner in $CurrentOwners){
            
            # Verificando se o Membro é um Owner
            if($Owner.name -ne $Member){
                
                try{
                    # Removendo o membro do grupo
                    Remove-UnifiedGroupLinks -Identity $Group -LinkType "Member" -Links $Member -Confirm:$false -ErrorAction stop | Out-Null
                } catch{ 
                    Log "*ATENCAO*: Não foi possível remover o membro $Member" -Print:$true
                }
            } else {
                
                # Retornando mensagem de atenção caso Membro seja um owner
                Log "*ATENCAO*: Não foi possível excluir o membro $Member. O mesmo é proprietário do grupo" -print:$true
            }
        }
    }
}

# Função responsável por remover uma lista de usuários da lista de owners do grupo
Function RemoveOwners($Group, $Owners){

    # Criando Array de Membros
    $Owners = $Owners.Split(",")

    # Percorrendo Array para realizar a operação de exclusão
    foreach($Owner in $Owners){
         
         try{
            # Removendo os Owners do grupo
            Remove-UnifiedGroupLinks -Identity $Group -LinkType "Owner" -Links $Owner -Confirm:$false -ErrorAction stop | Out-Null
         } catch{ 
            Log "*ATENCAO*: Erro ao remover o proprietário $Owner" -Print:$true
         }
    }

}

# Função responsável por excluir totalmente um grupo
Function RemoveGroup($Group){
    
    # Remover o Grupo
    try {
        Remove-UnifiedGroup -Identity $Group.Alias -Confirm:$false
        Log "SUCESSO: O Grupo $($Group.DisplayName) foi excluído." -Print:$true
    }
    catch {
        Log "*ERRO*: ao excluir o grupo $($Group.Alias)" -Print:$true
        exit 1
    }
}

# Função responsável por Renomear um grupo
Function ChangeName($Group, $NewGroupName){
     
    # Renomeando o grupo
    try {
        Set-UnifiedGroup -Identity $Group.Alias -DisplayName $NewGroupName | Out-Null

        Log "SUCESSO: O Grupo $($Group.DisplayName) agora se chama $NewGroupName" -print:$true
    }
    catch {
        Log "*ERRO*: ao renomear o grupo $($Group.DisplayName)" -Print:$true
        exit 1
    }
}

# ------------------- Fim das Funções de Processamento ------------------- #





# ----------------------- Inicio do Processamento ------------------------ #

# Verificando qual a ação que o script vai tomar

# Ação de Criação
if($Action -eq "CL_CRIAR"){
    
    # Extraindo o nome do grupo do email
    $NameGroup = $NameGroup.Replace("@petrobras.com.br", "")
    $NameGroup2 = $NameGroup2.Replace("@petrobras.com.br", "")
    $NameGroup3 = $NameGroup3.Replace("@petrobras.com.br", "")

    # Criando um array de nomes
    $Names = @($NameGroup, $NameGroup2, $NameGroup3)

    # Capturando um nome disponível
    $Name = VerifyName -Names $Names

    # Verificando se o nome está disponível
    if($null -eq $Name){

        Log "*ERRO*: Ao criar o grupo. Nenhuma das opções de nome estão disponíveis" -Print:$true
        exit 1
    
    } else {
        
        # Criando o grupo
        CreateGroup -Name $Name -Members $Members -Owners $Owners
    }

    

} elseif($Action -eq "CL_ALTERA_NOME"){
    
    # Resgatando o grupo
    $Group = Get-UnifiedGroup -Identity $OldGroupName -ErrorAction SilentlyContinue

    if($null -ne $group){
        
        # Verificando se o solicitante é proprietário do grupo
        $IsOwner = VerifyIfAffectedUserIsOwner -Group $Group
            
        if($IsOwner){

            # Alterando o nome da Sala
            ChangeName -Group $Group -NewGroupName $NewGroupName

        } else {
            
            Log "ATENCAO: O Usuário $AffectedUser não é proprietário do Grupo $($Group.DisplayName). Essa operação só pode ser realizada por um proprietário do grupo $($Group.DisplayName)" -Print:$true

        }
        

    } else {
        Log "*ERRO*: Grupo $OldGroupName não foi localizado, impossível realizar a alteração de nome" -Print:$true
        exit 1
    }
} else {
    
    # Resgatando o grupo que vai sofrer alteração ou exclusão
    $Group = Get-UnifiedGroup -Identity $GroupId -ErrorAction SilentlyContinue

    # Verifica se o grupo a ser alterado/excluído existe
    if ($null -ne $Group){

        # Ação de Alterar
        if($Action -eq "CL_ALTERAR"){

            # Verificando se o solicitante é proprietário do grupo
            $IsOwner = VerifyIfAffectedUserIsOwner -Group $Group
            
            if($IsOwner){

                # Incluir Apenas Membros
                if(($Members[0] -ne "N") -and ($Members[2] -ne "A") -and ($Owners[0] -eq "N") -and ($Owners[2] -eq "A")){
                
                    # Inserindo Membros no grupo
                    InsertMembers -Group $Group -Members $Members

                    Log "SUCESSO: A lista de membros foi atualizada" -Print:$true
                }

                # Incluir Apenas Proprietários
                if(($Members[0] -eq "N") -and ($Members[2] -eq "A") -and ($Owners[0] -ne "N") -and ($Owners[2] -ne "A")){
                
                    # Inserindo os Owners como Membro
                    <# IMPORTANTE: de acordo com as regras do exchange para grupos, um proprietário só pode ser 
                    adicionado ao grupo se antes ele for membro do mesmo grupo, desta forma se faz necessário a
                    chamada da função de inserção de membros passando como parâmetro a lista de owners, fazendo 
                    com que os owners atinjam a condição de membros antes de owners #>
                    InsertMembers -Group $Group -Members $Owners

                    #Inseriondo os Owners no grupo
                    InsertOwners -Group $Group -Owners $Owners
                
                    Log "SUCESSO: A lista de proprietários foi atualizada" -Print:$true             
                }

                #Inclusão de Membros e Proprietarios
                if(($Members[0] -ne "N") -and ($Members[2] -ne "A") -and ($Owners[0] -ne "N") -and ($Owners[2] -ne "A")){
            
                    # Inserindo Membros no grupo
                    InsertMembers -Group $Group -Members $Members

                    # Inserindo os Owners como Membro
                    <# IMPORTANTE: de acordo com as regras do exchange para grupos, um proprietário só pode ser 
                    adicionado ao grupo se antes ele for membro do mesmo grupo, desta forma se faz necessário a
                    chamada da função de inserção de membros passando como parâmetro a lista de owners, fazendo 
                    com que os owners atinjam a condição de membros antes de owners #>
                    InsertMembers -Group $Group -Members $Owners

                    #Inseriondo os Owners no grupo
                    InsertOwners -Group $Group -Owners $Owners
                
                    Log "SUCESSO: A lista de membros e proprietários foi atualizada"  -Print:$true
                }

            } else {
                
                Log "ATENCAO: O Usuário $AffectedUser não é proprietário do Grupo $($Group.DisplayName). Essa operação só pode ser realizada por um proprietário do grupo $($Group.DisplayName)" -Print:$true

            }
              
        }


        # Ação de Excluir
        if($Action -eq "CL_EXCLUIR"){

            # Verificando se o solicitante é proprietário do grupo
            $IsOwner = VerifyIfAffectedUserIsOwner -Group $Group
            
            if($IsOwner){

                # Verifica o que vai ser excluido

                # Exclusão do Grupo
                if(($Members[0] -eq "N") -and ($Members[2] -eq "A") -and ($Owners[0] -eq "N") -and ($Owners[2] -eq "A")){
                
                    # Removendo o Grupo
                    RemoveGroup -Group $Group
                }

                # Exclusão de Membros apenas
                if(($Members[0] -ne "N") -and ($Members[2] -ne "A") -and ($Owners[0] -eq "N") -and ($Owners[2] -eq "A")){
                
                    # Excluindo Membros
                    RemoveMembers -Group $Group -Members $Members

                    Log "SUCESSO: Lista de Membros Atualizada" -Print:$true
                }

                # Exclusão de Proprietários apenas
                if(($Members[0] -eq "N") -and ($Members[2] -eq "A") -and ($Owners[0] -ne "N") -and ($Owners[2] -ne "A")){
                
                    # Excluindo Proprietários
                    RemoveOwners -Group $Group -Owners $Owners    
                
                    Log "SUCESSO: Lista de Proprietários Atualizada" -Print:$true         
                }

                #Exclusão de Membros e Proprietarios
                if(($Members[0] -ne "N") -and ($Members[3] -ne "A") -and ($Owners[0] -ne "N") -and ($Owners[3] -ne "A")){
            
                    # Excluindo Membros
                    RemoveMembers -Group $Group -Members $Members

                    # Excluindo Proprietários
                    RemoveOwners -Group $Group -Owners $Owners  

                    Log "SUCESSO: Lista de Proprietários e Membros atualizada" -Print:$true
                }
            
            } else {
            
                Log "ATENCAO: O Usuário $AffectedUser não é proprietário do Grupo $($Group.DisplayName). Essa operação só pode ser realizada por um proprietário do grupo $($Group.DisplayName)" -print:$true
            }
        }

    } else{
        
        Log "*ERRO*: Grupo informado para operação não existe" -Print:$true
        exit 1
    }
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null

# ----------------------- Fim do Processamento ------------------------ #













    
