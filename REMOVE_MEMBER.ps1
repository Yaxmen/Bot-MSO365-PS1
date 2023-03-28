# Importa o módulo de conexão com o Banco de Dados
Import-Module "D:\Automation\Petro\TTR-Click\PowerShell\DatabaseModule.ps1"

# Recupera o tickets de acordo com o TicketType da demanda
$Tickets = GetWorkTickets -TicketType "REMOVE_MEMBER"

# Verifica se existem Tickets
if(!$Tickets){
    Write-Output "================================================================================="
    Write-output "Oferta: Criação/Alteração/Exclusão de lista de distribuição - Outlook"
    Write-Output "Ticket Type: REMOVE_MEMBER"
    Write-Output "Nenhum Ticket Localizado"
    Write-Output "=================================================================================="
    Write-Output " "
} else {

    # Percorre todos os chamados para o atendimento
    foreach($ticket in $tickets){

	    try{ 
	    $DistribuitionGroup = Get-DistributionGroup -Identity $ticket.IdentificadorDaLista -ErrorAction Stop
	    } catch {
		    InsertError -Solicitacao $ticket.solicitacao -ErrorMessage $_
	    }
		    ArrayChaves = @{}
            ArrayChaves = $ticket.Chaves.Split(",")
 
            foreach($Chave in $ArrayChaves){

                # Verifica se já contém a chave no grupo 
                $Member = Get-DistributionGroupMember -Identity $Chave  
                if($Member.Contains($Chave)){ 
                    try{

                       # Remover a chave solicitada
                       Remove-DistributionGroupMember -Identity $ticket.IdDaLista -Member $chave                           
                    } catch {

                } else {
                     # Inserindo as mensagens de erro no Sistema 
                     $ErrorMsg = "Chave solicitada para remoção, não existe no grupo"
                     Write-Host "Erro: $ErrorMsg"

                # Verifica se já contém a chave no grupo 
                $Owner = Get-DistributionGroupOwner -Identity $Chave  
                if($Owner.Contains($Chave)){ 
                    try{

                       # Remover a chave solicitada
                       Remove-DistributionGroupOwner -Identity $ticket.IdDaLista -Owner $chave                           
                    } catch {                         

                        InsertErrorMessage -Solicitacao $ticket.Solicitacao -NotaDeFechamento -Causa -CausaItem #verificar se o erro é fatal, se nao for, printar na tela (InsertError)
        		       }

                    } else {
                        # Inserindo as mensagens de erro no Sistema 
                        $ErrorMsg = "Chave solicitada para remoção, não existe no grupo"
                        Write-Host "Erro: $ErrorMsg"
                 }
            }
                    
            $CloseNote = "Conforme solicitado, o(s) membros(os) informado no chamado fora removidos do grupo solicitado"
            $ItemDeCausa = "Outlook"
            $Causa = "Remove"
            InsertCloseInfo -CloseNote $CloseNote -Solicitacao $Solicitacao -ItemDeCausa $ItemDeCausa -Causa $Causa
                    
            # Exibe Mensagem de Sucesso no terminal
            Write-Output "Ticket: $Solicitacao || Sucesso"     
            Write-Output $DistribuitionGroup        
                
        }
    }
}
}