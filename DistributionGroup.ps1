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
# Este script tem a finalidade de atender as solicita��es de "Cria��o/Altera��o/Exclus�o   #
# de lista de distribui��o - Outlook"                                                      #
#                                                                                          #
# Ele recebe por parametro os Nomes e Email atual da lista e o novo nome e novo email da���#
# lista e executa o seguinte:������������������������������������������������������������� #
#�� 1) Verifica se a lista de distribui��o existe����������������������������������������� #
#�� 2) Verifica qual � a demanda���������������������������������������������������������� #
#�� 3) Verifica se o Novo Email est� dispon�vel������������������������������������������� #
#�� 4) Executa as altera��es�������������������������������������������������������������� #
#------------------------------------------------------------------------------------------#

$username = "SAMSAZU@petrobras.com.br"
$msolKeyAuth = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content "D:\Password\password.txt" -ErrorAction Stop | ConvertTo-SecureString -Key $msolKeyAuth -ErrorAction Stop
$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password -ErrorAction Stop

#$username = "SAN3MSOFFICE@petrobrasbrteste.petrobras.com.br"
#$PlainPassword="Ror66406"
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

#Fun��o para adicionar membros a um grupo Rec�m criado (Que n�o possui nenhum membro ainda)
function AddMembersToNewDistributionGroup($DistributionGroup, $Members){

    # Criando Arrays de Sucesso e falha 
    $SuccessListMembers = [System.Collections.ArrayList]::new()
    $FailedListMembers = [System.Collections.ArrayList]::new()


    # Percorrendo o Array de Membros para executar a inser��o
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
    
        Log -message "Sucesso na Inser��o dos membros no grupo:"

        foreach($Member in $SuccessListMembers){
            
            Log " -$Member"
        }

        if($FailedListMembers -gt 0){
            
            Log "Aten��o: Os seguintes usu�rios n�o foram inseridos por n�o serem localizados no Exchange:"

            foreach($Member in $FailedListMembers){
                
                Log " -$Member"
            }
        }
        
    } else {
        
        Log "*Erro*: N�o houveram casos de sucesso na Inser��o de Membros"
        exit 1
    }


}

# Fun��o para adicionar membros a uma lista de distribui��o
function AddMembersToDistributionGroup($DistributionGroup, $Members) {

    # Capturando os Membros
    $DistributionGroupMembers = Get-DistributionGroupMember -Identity $DistributionGroup.alias

    # Criando array de Sucesso e Falha
    $SuccessListMembers = [System.Collections.ArrayList]::new()
    $AlreadyIncludeMemberList = [System.Collections.ArrayList]::new()
    $FailedListMembers = [System.Collections.ArrayList]::new()

    # Checando se existem membros para serem inseridos
    if(($Members[0] -eq "N") -or ($Members[1] -eq "A")){
        
        log "N�o foi solicitado a inclus�o de nenhum membro na Lista de Distribui��o"
    
    } else {
        
        # Percorrendo todos os Membros
        foreach ($Member in $Members){
            
            # Capturando o usu�rio atual
            $MemberMailBox = Get-Mailbox -Identity $Member -ErrorAction SilentlyContinue

            # Verifica se o Usu�rio existe
            if($MemberMailBox){
                
                # Verifica se o usu�rio j� est� na lista de distribui��o
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


        # Verificando se houveram usu�rios n�o inseridos por j� serem membros do grupo
        if ($AlreadyIncludeMemberList.Length -gt 0){
        
            Log "Aten��o, os seguintes usu�rios n�o foram inclusos pois j� eram membros da lista de distribui��o:"

            foreach ($User in $AlreadyIncludeMemberList){
            
                Log "-$User"
            }
        }

        # Verificando se houveram usu�rios n�o inseridos pois n�o foram localizados no Exchange
        if ($FailedListMembers.Length -gt 0){
        
            Log "Aten��o, os seguintes usu�rios n�o foram inclusos pois n�o foram localizados no Exchange:"

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
        
            Log "ERRO: N�o houve nenhum caso de sucesso na inser��o de Membros"
            exit 1
        }
    }
}

# Fun��o para adicionar propriet�rios a uma lista de distribui��o
function AddOwnersToDistributionGroup($DistributionGroup, $Owners){
    
    # Criando Array de Falha e Sucesso para Onwers
    $SuccessListOwners = [System.Collections.ArrayList]::new()
    $FailedListOwners = [System.Collections.ArrayList]::new()
    $AlreadyIsOwnerList = [System.Collections.ArrayList]::new()

    # Checando se foi solicitado a inclus�o de propriet�rios
    if(($Owners[0] -eq "N") -or ($Owners[1] -eq "A")){

        Log "N�o foi solicitado a inclus�o de nenhum Propriet�rio na Lista de Distribui��o"

    } else {
        
        # Percorrendo a lista de propriet�rios
        foreach($Owner in $Owners){
            
            # Validando se a chave existe no Exchange
            $OwnerBox = Get-Mailbox -Identity $Owner -ErrorAction SilentlyContinue

            if($OwnerBox){

                # Verificando se o usu�rio j� � propriet�rio
                if($DistributionGroup.ManagedBy.Contains($OwnerBox.Alias)){
                    
                    $AlreadyIsOwnerList.Add($OwnerBox.Alias) | Out-Null
                
                } else{
                    
                    # Adiciona o usu�rio na lista de propriet�rios
                    $DistributionGroup.ManagedBy.Add($OwnerBox.Alias) | Out-Null

                    $SuccessListOwners.Add($OWnerBox.Alias) | Out-Null
                }
            
            } else {
                
                $FailedListOwners.Add($Owner) | Out-Null
            }
        }

        # Atualizando a lista de propriet�rios na Tenant
        Set-DistributionGroup -Identity $DistributionGroup.Alias -ManagedBy $DistributionGroup.ManagedBy | Out-Null


        # Verificando se houveram usu�rios n�o inseridos pois j� eram propriet�rios da lista de distribui��o
        if($AlreadyIsOwnerList.Length -gt 0){

            Log "ATEN��O: os seguintes usu�rios n�o foram inclusos na lista de distribui��o pois j� eram propriet�rios"

            foreach($User in $AlreadyIsOwnerList){
                
                Log "-$User"
            
            }
        }

        # Verificando se houveram usu�rios n�o inseridos pois n�o foram localizados no Exchange
        if($FailedListOwners.Length -gt 0){
            
            Log "ATEN��O: os seguintes usu�rios n�o foram inclusos na lista de dsitrbui��o pois n�o foram localizados no Exchange"

            foreach($User in $FailedListOwners){

                Log "-$User"
            }
        }

        # Verificando se houveram usu�rios inseridos na Lista de distirbui��o e encerrando o chamado
        if($SuccessListOwners.Length -gt 0){
            
            Log "SUCESSO: Os seguintes usu�rios foram inclu�dos como propriet�rios da Lista de Distribui��o"

            foreach($User in $SuccessListOwners){
                
                Log "-$User"
            }
        } else {
            
            Log "ERRO: N�o houveram nenhum caso de sucesso na inclus�o de propriet�rios na lista de distribui��o"
            exit 1
        }
    }
}

# Fun��o para adicionar Remetentes a uma lista de distribui��o
function AddSendersToDistributionGroup($DistributionGroup, $Senders){

    # Criando Array de Falha e Sucesso para Senders
    $SuccessListSenders = [System.Collections.ArrayList]::new()
    $FailedListSenders = [System.Collections.ArrayList]::new()
    $AlreadyIsSenderList = [System.Collections.ArrayList]::new()

    if(($Senders[0] -eq "N") -or ($Senders[1] -eq "A")){
		Log "N�o foi solicitado a inser��o de nenhum remetente a Lista de Distribui��o"
	} else {
        
        # Percorrendo os Remetentes
        foreach($Sender in $Senders) {
            
            $SenderMailBox = Get-Mailbox -Identity $Sender -ErrorAction SilentlyContinue

            if($SenderMailBox){

                # Verificando se o usu�rio j� � um Remetente
                if($DistributionGroup.AcceptMessagesOnlyFrom.Contains($SenderMailBox.Alias)){
                    
                    $AlreadyIsSenderList.Add($SenderMailBox.Alias) | Out-Null

                } else {
                    
                    # Adicionando o usu�rio a lista de Remetentes
                    $DistributionGroup.AcceptMessagesOnlyFrom.Add($SenderMailBox.Alias)

                    $SuccessListSenders.Add($SenderMailBox.Alias) | Out-Null

                }

            } else {

                $FailedListSenders.Add($SenderMailBox.Alias) | Out-Null

            }
        }


        # Verificando se houveram usu�rios que n�o foram inseridos pois j� eram Remetentes
        if($AlreadyIsSenderList.Length -gt 0){
        
            Log "Aten��o: os seguintes usu�rios n�o foram adicionados a lista de remetentes pois j� eram remetentes"

            foreach($User in $AlreadyIsSenderList){
                Log "-$User"
            }
        }
        # Verificando se houveram usu�rios que n�o foram inseridos pois n�o foram localizados no Exchange
        if($FailedListSenders.Length -gt 0){
        
            Log "Aten��o: os seguintes usu�rios n�o foram adicionados a lista de distribui��o poisn n�o foram localizados no Exchange"

            foreach($User in $FailedListSenders.Length -gt 0){

                Log "-$User"
            }
        }

        # Verificando se os usu�rios foram inseridos e encerrando a inser��o
        if ($SuccessListSenders.Length -gt 0){
        
            Log "Sucesso: os seguintes usu�rios foram inseridos como Remetentes da Lista de distribui��o"

            foreach($User in $SuccessListSenders){

                Log " -$User"
            }
        } else {
        
            Log "ERRO: N�o houve nenhum caso de sucesso na inser��o de remetentes na Lista de distribui��o"
            exit 1
        }
    }

    # Definindo a nova lista de remetentes
    Set-DistributionGroup -Identity $DistributionGroup.Alias -AcceptMessagesOnlyFrom $DistributionGroup.AcceptMessagesOnlyFrom | Out-Null

}

# Fun��o para criar uma lista de distribui��o
function CreateDistributionGroup($Members, $Name, $Owners, $DepartRestriction, $JoinRestriction, $Senders) {

        
            
            # Criando alias
            $AliasMail = $Name.Replace(" ","")

            # Criando o Email
            $MailAddress = $AliasMail + "@petrobras.com.br"

            # Criar Grupo
            try{
                New-DistributionGroup -Name $Name -Alias $AliasMail -PrimarySmtpAddress $MailAddress -ErrorAction Stop | Out-Null
            } catch {
                Write-output "Erro na Cria��o do Grupo $Name"
            }

            Start-Sleep 20

            # Resgatando a lista rec�m criada
            $DistributionGroup = Get-DistributionGroup -Identity $AliasMail
			
			if($null -eq $DistributionGroup){
				Log "Ocorreu um erro ao buscar a Lista de Distribui��o rec�m criada."
				exit 1	
			}

            # Chamando a fun��o de Adicionar Membros
            AddMembersToNewDistributionGroup -DistributionGroup $DistributionGroup -Members $Members

            # Chamando a fun��o de adicionar propriet�rios
            AddOwnersToDistributionGroup -DistributionGroup $DistributionGroup -Owners $Owners

            # Removendo a conta de servi�o da lista de Propriet�rios
            # Id conta DEV: 9311cc06-c0fc-45ea-ab17-350da97dab0e
            $DistributionGroup.ManagedBy.Remove("deb4c82c-00cf-43b9-9a58-669f36dca86b")

            # Definindo os propriet�rios
            Set-DistributionGroup -Identity $DistributionGroup.Alias -ManagedBy $DistributionGroup.ManagedBy | Out-Null

            # Chamando a fun��o para adicionar os remetentes
            AddSendersToDistributionGroup -DistributionGroup $DistributionGroup -Senders $Senders
                   
            # Definindo restri��es de entrada e sa�da
			Set-DistributionGroup -Identity $DistributionGroup.Alias -MemberDepartRestriction $DepartRestriction -MemberJoinRestriction $JoinRestriction | Out-Null
            
            $GroupName = $DistributionGroup.DisplayName
               
            Write-Output "Sucesso na Cria��o do grupo $GroupName"
        
}


# Fun��o para remover uma lista de distribui��o
function RemoveDistributionGroup($DistributionGroup) {
    
    try {

        #Remove o grupo de distribui��o
        Remove-DistributionGroup -Identity $DistributionGroup.alias -Confirm:$false

        Write-Output "Sucesso na remo��o da Lista de Distribui��o " $DistributionGroup.DisplayName

    } catch {

        Write-Output "Grupo " + $DistributionGroup.DisplayName + " n�o localizado"
    }
}

# Fun��o para remover membros de um grupo de lista de distribui��o
function RemoveMembersFromDistributionGroup($DistributionGroup, $Members, $Owners, $Senders) {
    
    # Capturando os membros do grupo
    $DistributionGroupMembers = Get-DistributionGroupMember -Identity $DistributionGroup.alias

    # Removendo membros
    foreach($Member in $Members){
  
        #Validando se o membro existe
        $MemberMailBox = Get-Mailbox -Identity $Member -ErrorAction SilentlyContinue

        if($MemberMailBox){
            
            #Validando se o membro est� no grupo
            if($DistributionGroupMembers.Name.Contains($MemberMailBox.alias)){

                #Remover o membro
                Remove-DistributionGroupMember -Identity $DistributionGroup.alias -Member $MemberMailBox.alias -Confirm:$false | Out-Null
        
            } else{
                
                $GroupName = $DistributionGroup.DisplayName

                Write-Output "ERRO: Chave $Member n�o localizada no grupo $GroupName"
            }

        } else{
            
             Write-Output "ERRO ao deletar MEMBRO: Chave $Member n�o localizada"
        
        } 
    }

    # Removendo Owners
    foreach($Owner in $Owners){

        

        # Verificando se o usuario existe
        $OwnerMailBox = Get-Mailbox -Identity $Owner -ErrorAction SilentlyContinue

        if($OwnerMailBox){
        
            # Verificando se o usu�rio � um propriet�rio
            if($DistributionGroup.ManagedBy.Contains($OwnerMailBox.alias)){
                
                # Removendo o usu�rio da lista de propriet�rios
                $DistributionGroup.ManagedBy.Remove($OwnerMailBox.alias)
            } else {
            
                Write-Output "Erro: Usu�rio $Owner n�o localizado na lista de propriet�rios"
            }
        } else {
            
            Write-Output "ERRO ao deletar OWNER: Chave $Owner n�o localizado"
            
        }
    }

    # Definindo os Propriet�rios
    Set-DistributionGroup -Identity $DistributionGroup.alias -ManagedBy $DistributionGroup.ManagedBy | Out-Null
    
	if(($Senders[0] -eq "N") -or ($Senders[1] -eq "A")){
		Log "N�o foi solicitado a remo��o de nenhum remetente a Lista de Distribui��o"
	} else {
		
		# Validando remetentes
		foreach($Sender in $Senders) {
                
			#Validando o Sender
			$ValidSender = Get-Mailbox -Identity $Sender -ErrorAction SilentlyContinue

			if($ValidSender){

				# Validando se o Sender � um remetente
				if($DistributionGroup.AcceptMessagesOnlyFrom.Contains($ValidSender.Alias)){
                
					#Adicionando a lista de remetentes
					$DistributionGroup.AcceptMessagesOnlyFrom.Remove($ValidSender.Alias)

				} else {
            
					Log "ERRO ao remover REMETENTE: $Sender n�o localizado na Lista "
				}
			} else{
            
				Log "Remetente $Sender n�o localizado"

			}    
		}
	}
    

    # Definindo os remetentes
    Set-DistributionGroup -Identity $DistributionGroup.Alias -AcceptMessagesOnlyFrom $DistributionGroup.AcceptMessagesOnlyFrom

    # Mensagem de Sucesso
    Log "Sucesso na Remo��o de Membros, Propriet�rios e Remetentes" 

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
                Write-Output "ERRO: Nenhum Nome Dispon�vel para cria��o da lista"
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



if($acao -eq "CRIAR LISTA COM AT� 20 MEMBROS"){
      
      #Validar o Nome
      $DistributionName = ValidName -Name $Name -Name2 $Name2 -Name3 $Name3

      #Validar JoinRestriction
      $JoinRestriction = ValidJoinPermission -Entrada $Entrada

      #Validar DepartRestriction
      $DepartRestriction = ValidDepartPermission -Saida $Saida

      CreateDistributionGroup -Members $ArrayMembros -Owners $ArrayOwners -Name $DistributionName -DepartRestriction $DepartRestriction -JoinRestriction $JoinRestriction -Senders $ArraySenders

} elseif($acao -eq "CRIAR LISTA COM MEMBROS DE UMA GER�NCIA") {

    Log "**Aten��o**: Esse script n�o atende demandas de cria��o de lista utilizando membros de uma ger�ncia, favor realizar o tratamento manual dessa demanda."
    exit 1

} elseif($acao -eq "CRIAR LISTA COM MAIS DE 20 MEMBROS"){
    
    Log "**Aten��o**: Esse script n�o atende demandas de cria��o de lista utilizando com mais de 20 membros, favor realizar o tratamento manual dessa demanda."
    exit 1
}



$DistributionGroup = Get-DistributionGroup -Identity $Email -ErrorAction SilentlyContinue

if($DistributionGroup){

    # Checando se o solicitante � propriet�rio da Lista de Distribui��o

    if($DistributionGroup.ManagedBy -notcontains $AffectedUser){

        Log "*ERRO*: Usu�rio solicitante n�o � propriet�rio da Lista de distribui��o $($DistributionGroup.DisplayName)"
            exit 1
    }



    # Realizando Atendimento para Inclus�o de Chaves
    if($acao -eq "INCLUIR CHAVES"){

        # Checando se chegaram Membros, Propriet�rios e Remetentes para incluir na Lista de Distribui��o

         if($ArrayMembros[0] -ne "N"){

            # Chamando fun��o de inclus�o de Membros
            AddMembersToDistributionGroup -DistributionGroup $DistributionGroup -Members $ArrayMembros
                
         }

         if($ArrayOwners[0] -ne "N"){

            # Chamando a fun��o de inclus�o de Propriet�rios
            AddOwnersToDistributionGroup -DistributionGroup $DistributionGroup -Owners $ArrayOwners
                
         }

         if($ArraySenders[0] -ne "N"){

            # Chamando a fun��o de inclus�o de Remetentes
            AddSendersToDistributionGroup -DistributionGroup $DistributionGroup -Senders $ArraySenders
                
         }
    }


    # Realizando atendimento de Excluir Chaves
    if($acao -eq "EXCLUIR CHAVES"){
    
        # Chamando Fun��o de Exclus�o de membros
        RemoveMembersFromDistributionGroup -DistributionGroup $DistributionGroup -Members $ArrayMembros -Owners $ArrayOwners -Senders $ArraySenders
    
    }


    # Realizando atendimento de Excluir a Lista de Distribui��o
    if($acao -eq "EXCLUIR A LISTA"){

        # Chamando fun��o de remo��o do Grupo
        RemoveDistributionGroup -DistributionGroup $DistributionGroup
    
    }

} else {

    Log "Erro: Lista de ditribui��o $Email n�o localizada no exchange"
    exit 1

}
