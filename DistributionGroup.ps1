Param (
    
    [Parameter( Mandatory=$true)] [String]$AffectedUser,
    [Parameter( Mandatory=$true)] [String]$Name,
    [Parameter( Mandatory=$true)] [String]$Name2, 
    [Parameter( Mandatory=$true)] [String]$acao,
    [Parameter( Mandatory=$true)] [String]$Email, 
    [Parameter( Mandatory=$true)] [String]$Remetente,
    [Parameter( Mandatory=$true)] [String]$Member,
    [Parameter( Mandatory=$true)] [String]$Owner,
    [Parameter( Mandatory=$true)] [String]$Entrada,
    [Parameter( Mandatory=$true)] [String]$Saida	

)

#------------------------------------------------------------------------------------------#
# Este script tem a finalidade de atender as solicitações de "Criação/Alteração/Exclusão   #
# de lista de distribuição - Outlook"                                                      #
#                                                                                          #
# Ele recebe por parametro os Nomes e Email atual da lista e o novo nome e novo email da   #
# lista e executa o seguinte:                                                              #
#   1) Verifica se a lista de distribuição existe                                          #
#   2) Verifica qual é a demanda                                                           #
#   3) Verifica se o Novo Email está disponível                                            #
#   4) Executa as alterações                                                               #
#------------------------------------------------------------------------------------------#

$username = "email"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

#$username = "email"
#$PlainPassword="senha"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword

$logFile = "d:\Util\DistributionGroup\logs\"+(Get-Date).ToString('yyyyMMdd')+"_distributionGroup_log.csv"

$SLEEP = 60

function Log([string]$message){

    $datetime = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    if ( $message -like '*ERRO*') {
        Write-Host "$message" -ForegroundColor Red
    }
    else { 
        if ( $message -like '*ATENCAO*') {
            Write-Host "$message" -ForegroundColor Yellow
        }
        else {
            Write-Host "$message" -ForegroundColor Green
        }
    }
    Add-Content -Path $logFile -Value "$datetime;$message"
}

try {
    Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -InformationAction SilentlyContinue | Out-Null 
}
catch {
    Start-Sleep $SLEEP
    try {
        Connect-ExchangeOnline -Credential $credential -InformationAction SilentlyContinue | Out-Null
    }
    catch {
        Start-Sleep $SLEEP
        try {
            Connect-ExchangeOnline -Credential $credential -InformationAction SilentlyContinue | Out-Null
        }
        catch {    
            Log "ERRO: ao conectar Exchange"
            exit 1
        }        
    }
}

#Função para adicionar membros a um grupo Recém criado (Que não possui nenhum membro ainda)
function AddMembersToNewDistributionGroup($DistributionGroup, $Members){

    # Criando Arrays de Sucesso e falha 
    $SuccessListMembers = [System.Collections.ArrayList]::new()
    $FailedListMembers = [System.Collections.ArrayList]::new()


    # Percorrendo o Array de Membros para executar a inserção
    foreach($Member in $Members){
        
        $MemberMailBox = Get-Mailbox -Identity $Member -ErrorAction SilentlyContinue

        if($MemberMailBox){
            
            Add-DistributionGroupMember -Identity $DistributionGroup.Alias -Member $MemberMailBox.Alias | Out-Null

            $SuccessListMembers.Add($MemberMailBox.Alias) | Out-Null

        } else {
            
            $FailedListMembers.Add($MemberMailBox.Alias) | Out-Null
        }
    }

    if($SuccessListMembers.Length -gt 0){
    
        Log -message "Sucesso na Inserção dos membros no grupo:"

        foreach($Member in $SuccessListMembers){
            
            Log " -$Member"
        }

        if($FailedListMembers -gt 0){
            
            Log "Atenção: Os seguintes usuários não foram inseridos por não serem localizados no Exchange:"

            foreach($Member in $FailedListMembers){
                
                Log " -$Member"
            }
        }
        
    } else {
        
        Log "*Erro*: Não houveram casos de sucesso na Inserção de Membros"
        exit 1
    }


}

# Função para adicionar membros a uma lista de distribuição
function AddMembersToDistributionGroup($DistributionGroup, $Members) {

    # Capturando os Membros
    $DistributionGroupMembers = Get-DistributionGroupMember -Identity $DistributionGroup.alias

    # Criando array de Sucesso e Falha
    $SuccessListMembers = [System.Collections.ArrayList]::new()
    $AlreadyIncludeMemberList = [System.Collections.ArrayList]::new()
    $FailedListMembers = [System.Collections.ArrayList]::new()

    # Checando se existem membros para serem inseridos
    if(($Members[0] -eq "N") -or ($Members[1] -eq "A")){
        
        log "Não foi solicitado a inclusão de nenhum membro na Lista de Distribuição"
    
    } else {
        
        # Percorrendo todos os Membros
        foreach ($Member in $Members){
            
            # Capturando o usuário atual
            $MemberMailBox = Get-Mailbox -Identity $Member -ErrorAction SilentlyContinue

            # Verifica se o Usuário existe
            if($MemberMailBox){
                
                # Verifica se o usuário já está na lista de distribuição
                if($DistributionGroupMembers.Name.Contains($MemberMailBox.Alias)){

                    $AlreadyIncludeMemberList.Add($MemberMailBox.Alias)
                } else {
                    
                    Add-DistributionGroupMember -Identity $DistributionGroup.Alias -Member $MemberMailBox.Alias | Out-Null

                    $SuccessListMembers.Add($MemberMailBox.Alias) | Out-Null
                }
            } else {
                    
                $FailedListMembers.Add($Member) | Out-Null
            }
        }


        # Verificando se houveram usuários não inseridos por já serem membros do grupo
        if ($AlreadyIncludeMemberList.Length -gt 0){
        
            Log "Atenção, os seguintes usuários não foram inclusos pois já eram membros da lista de distribuição:"

            foreach ($User in $AlreadyIncludeMemberList){
            
                Log "-$User"
            }
        }

        # Verificando se houveram usuários não inseridos pois não foram localizados no Exchange
        if ($FailedListMembers.Length -gt 0){
        
            Log "Atenção, os seguintes usuários não foram inclusos pois não foram localizados no Exchange:"

            foreach ($User in $FailedListMembers){
            
                Log "-$User"
            }
        }

        # Verificando se houveram casos de sucesso e finalizando o chamado
        if ($SuccessListMembers.Length -gt 0){
        
            Log "Sucesso: A lista de Membros foi atualizada os seguintes membros foram adicionados a lista:"

            foreach($User in $SuccessListMembers){
                Log "-$User"
            }

        } else {
        
            Log "ERRO: Não houve nenhum caso de sucesso na inserção de Membros"
            exit 1
        }
    }
}

# Função para adicionar proprietários a uma lista de distribuição
function AddOwnersToDistributionGroup($DistributionGroup, $Owners){
    
    # Criando Array de Falha e Sucesso para Onwers
    $SuccessListOwners = [System.Collections.ArrayList]::new()
    $FailedListOwners = [System.Collections.ArrayList]::new()
    $AlreadyIsOwnerList = [System.Collections.ArrayList]::new()

    # Checando se foi solicitado a inclusão de proprietários
    if(($Owners[0] -eq "N") -or ($Owners[1] -eq "A")){

        Log "Não foi solicitado a inclusão de nenhum Proprietário na Lista de Distribuição"

    } else {
        
        # Percorrendo a lista de proprietários
        foreach($Owner in $Owners){
            
            # Validando se a chave existe no Exchange
            $OwnerBox = Get-Mailbox -Identity $Owner -ErrorAction SilentlyContinue

            if($OwnerBox){

                # Verificando se o usuário já é proprietário
                if($DistributionGroup.ManagedBy.Contains($OwnerBox.Alias)){
                    
                    $AlreadyIsOwnerList.Add($OwnerBox.Alias) | Out-Null
                
                } else{
                    
                    # Adiciona o usuário na lista de proprietários
                    $DistributionGroup.ManagedBy.Add($OwnerBox.Alias) | Out-Null

                    $SuccessListOwners.Add($OWnerBox.Alias) | Out-Null
                }
            
            } else {
                
                $FailedListOwners.Add($Owner) | Out-Null
            }
        }

        # Atualizando a lista de proprietários na Tenant
        Set-DistributionGroup -Identity $DistributionGroup.Alias -ManagedBy $DistributionGroup.ManagedBy | Out-Null


        # Verificando se houveram usuários não inseridos pois já eram proprietários da lista de distribuição
        if($AlreadyIsOwnerList.Length -gt 0){

            Log "ATENÇÃO: os seguintes usuários não foram inclusos na lista de distribuição pois já eram proprietários"

            foreach($User in $AlreadyIsOwnerList){
                
                Log "-$User"
            
            }
        }

        # Verificando se houveram usuários não inseridos pois não foram localizados no Exchange
        if($FailedListOwners.Length -gt 0){
            
            Log "ATENÇÃO: os seguintes usuários não foram inclusos na lista de dsitrbuição pois não foram localizados no Exchange"

            foreach($User in $FailedListOwners){

                Log "-$User"
            }
        }

        # Verificando se houveram usuários inseridos na Lista de distirbuição e encerrando o chamado
        if($SuccessListOwners.Length -gt 0){
            
            Log "SUCESSO: Os seguintes usuários foram incluídos como proprietários da Lista de Distribuição"

            foreach($User in $SuccessListOwners){
                
                Log "-$User"
            }
        } else {
            
            Log "ERRO: Não houveram nenhum caso de sucesso na inclusão de proprietários na lista de distribuição"
            exit 1
        }
    }
}

# Função para adicionar Remetentes a uma lista de distribuição
function AddSendersToDistributionGroup($DistributionGroup, $Senders){

    # Criando Array de Falha e Sucesso para Senders
    $SuccessListSenders = [System.Collections.ArrayList]::new()
    $FailedListSenders = [System.Collections.ArrayList]::new()
    $AlreadyIsSenderList = [System.Collections.ArrayList]::new()

    if(($Senders[0] -eq "N") -or ($Senders[1] -eq "A")){
		Log "Não foi solicitado a inserção de nenhum remetente a Lista de Distribuição"
	} else {
        
        # Percorrendo os Remetentes
        foreach($Sender in $Senders) {
            
            $SenderMailBox = Get-Mailbox -Identity $Sender -ErrorAction SilentlyContinue

            if($SenderMailBox){

                # Verificando se o usuário já é um Remetente
                if($DistributionGroup.AcceptMessagesOnlyFrom.Contains($SenderMailBox.Alias)){
                    
                    $AlreadyIsSenderList.Add($SenderMailBox.Alias) | Out-Null

                } else {
                    
                    # Adicionando o usuário a lista de Remetentes
                    $DistributionGroup.AcceptMessagesOnlyFrom.Add($SenderMailBox.Alias)

                    $SuccessListSenders.Add($SenderMailBox.Alias) | Out-Null

                }

            } else {

                $FailedListSenders.Add($SenderMailBox.Alias) | Out-Null

            }
        }


        # Verificando se houveram usuários que não foram inseridos pois já eram Remetentes
        if($AlreadyIsSenderList.Length -gt 0){
        
            Log "Atenção: os seguintes usuários não foram adicionados a lista de remetentes pois já eram remetentes"

            foreach($User in $AlreadyIsSenderList){
                Log "-$User"
            }
        }
        # Verificando se houveram usuários que não foram inseridos pois não foram localizados no Exchange
        if($FailedListSenders.Length -gt 0){
        
            Log "Atenção: os seguintes usuários não foram adicionados a lista de distribuição poisn não foram localizados no Exchange"

            foreach($User in $FailedListSenders.Length -gt 0){

                Log "-$User"
            }
        }

        # Verificando se os usuários foram inseridos e encerrando a inserção
        if ($SuccessListSenders.Length -gt 0){
        
            Log "Sucesso: os seguintes usuários foram inseridos como Remetentes da Lista de distribuição"

            foreach($User in $SuccessListSenders){

                Log " -$User"
            }
        } else {
        
            Log "ERRO: Não houve nenhum caso de sucesso na inserção de remetentes na Lista de distribuição"
            exit 1
        }
    }

    # Definindo a nova lista de remetentes
    Set-DistributionGroup -Identity $DistributionGroup.Alias -AcceptMessagesOnlyFrom $DistributionGroup.AcceptMessagesOnlyFrom | Out-Null

}

# Função para criar uma lista de distribuição
function CreateDistributionGroup($Members, $Name, $Owners, $DepartRestriction, $JoinRestriction, $Senders) {

        
            
            # Criando alias
            $AliasMail = $Name.Replace(" ","")

            # Criando o Email
            $MailAddress = $AliasMail + "@petrobras.com.br"

            # Criar Grupo
            try{
                New-DistributionGroup -Name $Name -Alias $AliasMail -PrimarySmtpAddress $MailAddress -ErrorAction Stop | Out-Null
            } catch {
                Write-output "Erro na Criação do Grupo $Name"
            }

            Start-Sleep 20

            # Resgatando a lista recém criada
            $DistributionGroup = Get-DistributionGroup -Identity $AliasMail
			
			if($null -eq $DistributionGroup){
				Log "Ocorreu um erro ao buscar a Lista de Distribuição recém criada."
				exit 1	
			}

            # Chamando a função de Adicionar Membros
            AddMembersToNewDistributionGroup -DistributionGroup $DistributionGroup -Members $Members

            # Chamando a função de adicionar proprietários
            AddOwnersToDistributionGroup -DistributionGroup $DistributionGroup -Owners $Owners

            # Removendo a conta de serviço da lista de Proprietários
            # Id conta DEV: 9311cc06-c0fc-45ea-ab17-350da97dab0e
            $DistributionGroup.ManagedBy.Remove("deb4c82c-00cf-43b9-9a58-669f36dca86b")

            # Definindo os proprietários
            Set-DistributionGroup -Identity $DistributionGroup.Alias -ManagedBy $DistributionGroup.ManagedBy | Out-Null

            # Chamando a função para adicionar os remetentes
            AddSendersToDistributionGroup -DistributionGroup $DistributionGroup -Senders $Senders
                   
            # Definindo restrições de entrada e saída
			Set-DistributionGroup -Identity $DistributionGroup.Alias -MemberDepartRestriction $DepartRestriction -MemberJoinRestriction $JoinRestriction | Out-Null
            
            $GroupName = $DistributionGroup.DisplayName
               
            Write-Output "Sucesso na Criação do grupo $GroupName"
        
}


# Função para remover uma lista de distribuição
function RemoveDistributionGroup($DistributionGroup) {
    
    try {

        #Remove o grupo de distribuição
        Remove-DistributionGroup -Identity $DistributionGroup.alias -Confirm:$false

        Write-Output "Sucesso na remoção da Lista de Distribuição " $DistributionGroup.DisplayName

    } catch {

        Write-Output "Grupo " + $DistributionGroup.DisplayName + " não localizado"
    }
}

# Função para remover membros de um grupo de lista de distribuição
function RemoveMembersFromDistributionGroup($DistributionGroup, $Members, $Owners, $Senders) {
    
    # Capturando os membros do grupo
    $DistributionGroupMembers = Get-DistributionGroupMember -Identity $DistributionGroup.alias

    # Removendo membros
    foreach($Member in $Members){
  
        #Validando se o membro existe
        $MemberMailBox = Get-Mailbox -Identity $Member -ErrorAction SilentlyContinue

        if($MemberMailBox){
            
            #Validando se o membro está no grupo
            if($DistributionGroupMembers.Name.Contains($MemberMailBox.alias)){

                #Remover o membro
                Remove-DistributionGroupMember -Identity $DistributionGroup.alias -Member $MemberMailBox.alias -Confirm:$false | Out-Null
        
            } else{
                
                $GroupName = $DistributionGroup.DisplayName

                Write-Output "ERRO: Chave $Member não localizada no grupo $GroupName"
            }

        } else{
            
             Write-Output "ERRO ao deletar MEMBRO: Chave $Member não localizada"
        
        } 
    }

    # Removendo Owners
    foreach($Owner in $Owners){

        

        # Verificando se o usuario existe
        $OwnerMailBox = Get-Mailbox -Identity $Owner -ErrorAction SilentlyContinue

        if($OwnerMailBox){
        
            # Verificando se o usuário é um proprietário
            if($DistributionGroup.ManagedBy.Contains($OwnerMailBox.alias)){
                
                # Removendo o usuário da lista de proprietários
                $DistributionGroup.ManagedBy.Remove($OwnerMailBox.alias)
            } else {
            
                Write-Output "Erro: Usuário $Owner não localizado na lista de proprietários"
            }
        } else {
            
            Write-Output "ERRO ao deletar OWNER: Chave $Owner não localizado"
            
        }
    }

    # Definindo os Proprietários
    Set-DistributionGroup -Identity $DistributionGroup.alias -ManagedBy $DistributionGroup.ManagedBy | Out-Null
    
	if(($Senders[0] -eq "N") -or ($Senders[1] -eq "A")){
		Log "Não foi solicitado a remoção de nenhum remetente a Lista de Distribuição"
	} else {
		
		# Validando remetentes
		foreach($Sender in $Senders) {
                
			#Validando o Sender
			$ValidSender = Get-Mailbox -Identity $Sender -ErrorAction SilentlyContinue

			if($ValidSender){

				# Validando se o Sender é um remetente
				if($DistributionGroup.AcceptMessagesOnlyFrom.Contains($ValidSender.Alias)){
                
					#Adicionando a lista de remetentes
					$DistributionGroup.AcceptMessagesOnlyFrom.Remove($ValidSender.Alias)

				} else {
            
					Log "ERRO ao remover REMETENTE: $Sender não localizado na Lista "
				}
			} else{
            
				Log "Remetente $Sender não localizado"

			}    
		}
	}
    

    # Definindo os remetentes
    Set-DistributionGroup -Identity $DistributionGroup.Alias -AcceptMessagesOnlyFrom $DistributionGroup.AcceptMessagesOnlyFrom

    # Mensagem de Sucesso
    Log "Sucesso na Remoção de Membros, Proprietários e Remetentes" 

}


function ValidName($Name, $Name2){


    $AvailableName = Get-DistributionGroup -Identity $Name -ErrorAction SilentlyContinue

    if(!$AvailableName){
        
        return $Name

    } elseif($AvailableName) {

        $AvailableName2 = Get-DistributionGroup -Identity $Name2 -ErrorAction SilentlyContinue

        if(!$AvailableName2) {

            return $Name2
        
        } 
    } else {
                Write-Output "ERRO: Nenhum Nome Disponível para criação da lista"
                Exit
    }
}
	
function ValidJoinPermission($Entrada){

    if($Entrada -eq "CL_ABERTA"){
        return "Open"
    } elseif ($Entrada -eq "CL_FECHADA") {
        return "Closed"
    } elseif ($Entrada -eq "CL_PORAPROVACAO"){
        return "ApprovalRequired"
    }

}

function ValidDepartPermission($Saida){

    if($Saida -eq "CL_ABERTA"){
        return "Open"
    } elseif($Saida -eq "CL_FECHADA"){
        return "Closed"
    }
}

    

    #Criar Array com Membros
    $ArrayMembros = $Member.Split(",")

    #Criar Array com Owners
    $ArrayOwners = $Owner.Split(",")

    #Criar Array com Senders
    $ArraySenders = $Remetente.Split(",")



if($acao -eq "CRIAR LISTA COM ATÉ 20 MEMBROS"){
      
      #Validar o Nome
      $DistributionName = ValidName -Name $Name -Name2 $Name2 -Name3 $Name3

      #Validar JoinRestriction
      $JoinRestriction = ValidJoinPermission -Entrada $Entrada

      #Validar DepartRestriction
      $DepartRestriction = ValidDepartPermission -Saida $Saida

      CreateDistributionGroup -Members $ArrayMembros -Owners $ArrayOwners -Name $DistributionName -DepartRestriction $DepartRestriction -JoinRestriction $JoinRestriction -Senders $ArraySenders

} elseif($acao -eq "CRIAR LISTA COM MEMBROS DE UMA GERÊNCIA") {

    Log "**Atenção**: Esse script não atende demandas de criação de lista utilizando membros de uma gerência, favor realizar o tratamento manual dessa demanda."
    exit 1

} elseif($acao -eq "CRIAR LISTA COM MAIS DE 20 MEMBROS"){
    
    Log "**Atenção**: Esse script não atende demandas de criação de lista utilizando com mais de 20 membros, favor realizar o tratamento manual dessa demanda."
    exit 1
}



$DistributionGroup = Get-DistributionGroup -Identity $Email -ErrorAction SilentlyContinue

if($DistributionGroup){

    # Checando se o solicitante é proprietário da Lista de Distribuição

    if($DistributionGroup.ManagedBy -notcontains $AffectedUser){

        Log "*ERRO*: Usuário solicitante não é proprietário da Lista de distribuição $($DistributionGroup.DisplayName)"
            exit 1
    }



    # Realizando Atendimento para Inclusão de Chaves
    if($acao -eq "INCLUIR CHAVES"){

        # Checando se chegaram Membros, Proprietários e Remetentes para incluir na Lista de Distribuição

         if($ArrayMembros[0] -ne "N"){

            # Chamando função de inclusão de Membros
            AddMembersToDistributionGroup -DistributionGroup $DistributionGroup -Members $ArrayMembros
                
         }

         if($ArrayOwners[0] -ne "N"){

            # Chamando a função de inclusão de Proprietários
            AddOwnersToDistributionGroup -DistributionGroup $DistributionGroup -Owners $ArrayOwners
                
         }

         if($ArraySenders[0] -ne "N"){

            # Chamando a função de inclusão de Remetentes
            AddSendersToDistributionGroup -DistributionGroup $DistributionGroup -Senders $ArraySenders
                
         }
    }


    # Realizando atendimento de Excluir Chaves
    if($acao -eq "EXCLUIR CHAVES"){
    
        # Chamando Função de Exclusão de membros
        RemoveMembersFromDistributionGroup -DistributionGroup $DistributionGroup -Members $ArrayMembros -Owners $ArrayOwners -Senders $ArraySenders
    
    }


    # Realizando atendimento de Excluir a Lista de Distribuição
    if($acao -eq "EXCLUIR A LISTA"){

        # Chamando função de remoção do Grupo
        RemoveDistributionGroup -DistributionGroup $DistributionGroup
    
    }

} else {

    Log "Erro: Lista de ditribuição $Email não localizada no exchange"
    exit 1

}
