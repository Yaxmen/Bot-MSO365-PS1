# Importa o módulo de conexão com o Banco de Dados
Import-Module "D:\Automation\Petro\TTR-Click\PowerShell\DatabaseModule.ps1"

# Recupera o tickets de acordo com o TicketType da demanda
$Tickets = GetWorkTickets -TicketType "RENAME_NAME"

if(!$Tickets){
    Write-Output "=================================================================================="
    Write-output "Oferta: Alteração de nome/e-mail de caixa de correio compartilhada no Outlook"
    Write-Output "Ticket Type: RENAME_NAME"
    Write-Output "Nenhum Ticket Localizado"
    Write-Output "=================================================================================="
    Write-Output " "
} else {

    foreach($Ticket in $Tickets) {

        try{
    
            # Valida se a MailBox a ser alterada Existe
            $MailBoxExists = Get-MailBox -Identity $Ticket.EmailAtual
            
            # Capturando numero do ticket para exibir em tela
            $Solicitacao = $Ticket.solicitacao
    
            if($MailBoxExists){
    
                try {
                    
                    # Altera o Nome
                    Set-MailBox -Identity $Ticket.EmailAtual -DisplayName $Ticket.NovoNome -ErrorAction Stop

                    $CloseNote = "Conforme solicitado, a caixa de email " + $Ticket.EmailAtual + " teve seu nome de exibição alterado para " + $Ticket.NovoNome
                    $ItemDeCausa = "Outlook"
                    $Causa = "Alterar"
                    InsertCloseInfo -CloseNote $CloseNote -Solicitacao $Solicitacao -ItemDeCausa $ItemDeCausa -Causa $Causa
                    
                    # Exibe Mensagem de Sucesso no terminal
                    Write-Output "Ticket: $Solicitacao || Sucesso"             
                }
                catch {
                    InsertLog($BotName, "Exception", "Ticket: $Solicitacao, $_")
                    InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $_
                    Write-Output "Ticket: $Solicitacao|| Erro: $_"
                }
        
            } else {
        
                $ErrorMsg = "Mailbox nao localizada no Exchange Online"
                InsertLog($BotName, "Exception", "Ticket: $Solicitacao, $ErrorMsg")
                InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMsg
                Write-Output "Ticket: $Solicitacao || Erro: $ErrorMsg"
            }
    
        } catch{
            InsertLog($BotName, "Exception", "Ticket: $Solicitacao, $_")
            InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $_
            Write-Output "Ticket: $Solicitacao || Erro: $_"
        }
    
        
       
    }
}



