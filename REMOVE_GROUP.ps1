# Importa o módulo de conexão com o Banco de Dados
Import-Module "D:\Automation\Petro\TTR-Click\PowerShell\DatabaseModule.ps1"

# Recupera o tickets de acordo com o TicketType da demanda
$Tickets = GetWorkTickets -TicketType "REMOVE_GROUP"

# Verifica se existem Tickets
if(!$Tickets){
    Write-Output "================================================================================="
    Write-output "Oferta: Criação/Alteração/Exclusão de lista de distribuição - Outlook"
    Write-Output "Ticket Type: REMOVE_GROUP"
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
                # Verifica se o grupo existe 
                $Group = Get-DistributionGroup -Identity $Email  
                if($Group.Contains($Email)){ 
                    try{

                       # Remover o grupo solicitado
                       Remove-DistributionGroup -Identity $ticket.IdDaLista                      
                    } catch { 
			
			    InsertErrorMessage -Solicitacao $ticket.Solicitacao -NotaDeFechamento -Causa -CausaItem #verificar se o erro é fatal, se nao for, printar na tela (InsertError)
        		       }

                } else {
                        # Inserindo as mensagens de erro no Sistema 
                        $ErrorMsg = "Grupo informado para exclusão não localizado."
                        Write-Host "Erro: $ErrorMsg"
                 }
            }
                    
            $CloseNote = "Conforme solicitado, o grupo informado no chamado foi removido da Lista de Distribuição"
            $ItemDeCausa = "Outlook"
            $Causa = "Remove"
            InsertCloseInfo -CloseNote $CloseNote -Solicitacao $Solicitacao -ItemDeCausa $ItemDeCausa -Causa $Causa
                    
            # Exibe Mensagem de Sucesso no terminal
            Write-Output "Ticket: $Solicitacao || Sucesso"
            Write-Output $DistributionGroup             
                
            }