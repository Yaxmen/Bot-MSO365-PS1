# Importa o módulo de conexão com o Banco de Dados
Import-Module "D:\Automation\Petro\TTR-Click\PowerShell\DatabaseModule.ps1"

# Recupera o tickets de acordo com o TicketType da demanda
$Tickets = GetWorkTickets -TicketType "CREATE_GROUP"

# Verifica se existem Tickets
if(!$Tickets){
    Write-Output "================================================================================="
    Write-output "Oferta: Criação/Alteração/Exclusão de lista de distribuição - Outlook"
    Write-Output "Ticket Type: CREATE_GROUP"
    Write-Output "Nenhum Ticket Localizado"
    Write-Output "=================================================================================="
    Write-Output " "
} else {
    
    # Percorre todos os chamados para o atendimento
    foreach($ticket in $tickets){

	    try{ 
	    $DistributionGroup = Get-DistributionGroup -Identity $ticket.IdentificadorDaLista -ErrorAction Stop
	    } catch {
		    InsertError -Solicitacao $ticket.solicitacao -ErrorMessage $_
	    }
			
        ArrayChaves = @{} 
        ArrayChaves = $ticket.Chaves.Split(",")
     
        foreach($Chave in $ArrayChaves){
				
                # Verifica se o grupo existe 
                $Group = Get-DistributionGroup -Identity $Email  
                if($Group.Contains($Email)){ 
                    try{    
                        
                        # Criar Grupo
                        New-DistributionGroup -Name $ticket.IdDaLista -Alias $Email -PrimarySmtpAddress $GroupEmailAddress

                        # Obter o grupo e adicionar o nome de exibição         
                        $Group = Get-DistributionGroup -Identity $ticket.IdDaLista
                        $Group | Set-DistributionGroup -DisplayName $ticket.IdDaLista

                        # Adicionando membros e proprietarios ao grupo           
                        Add-DistributionGroupMember -Identity $ticket.IdDaLista -Member $chave
                        Add-DistributionGroupOwner -Identity $ticket.IdDaLista -Owner $chave      
                    } catch {
               
                        InsertErrorMessage -Solicitacao $ticket.Solicitacao -NotaDeFechamento -Causa -CausaItem #verificar se o erro é fatal, se nao for, printar na tela (InsertError)
                    }
            
                } else {
                     # Inserindo as mensagens de erro no Sistema 
                     $ErrorMsg = "Erro na criação do grupo."
                     Write-Host "Erro: $ErrorMsg"           
                    }         
        }
                 
        $CloseNote = "Conforme solicitado, o grupo informado no chamado foi criado na Lista de Distribuição" 
        $ItemDeCausa = "Outlook"
        $Causa = "Criar"

        InsertCloseInfo -CloseNote $CloseNote -Solicitacao $Solicitacao -ItemDeCausa $ItemDeCausa -Causa $Causa
                 
         # Exibe Mensagem de Sucesso no terminal
         Write-Output "Ticket: $Solicitacao || Sucesso"
         Write-Output $DistributionGroup             
                 
                
        }
	}