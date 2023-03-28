# Importa o módulo de conexão com o Banco de Dados
Import-Module "D:\Automation\Petro\TTR-Click\PowerShell\DatabaseModule.ps1"

# Recupera o tickets de acordo com o TicketType da demanda
$Tickets = GetWorkTickets -TicketType "RENAME_MAIL"

if(!$Tickets){
    Write-Output "=================================================================================="
    Write-output "Oferta: Alteração de nome/e-mail de caixa de correio compartilhada no Outlook"
    Write-Output "Ticket Type: RENAME_MAIL"
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
                    
                    # Se a MailBox existe checa se o Novo Nome e Email estão disponíveis
                    $MailNameAlreadyExists = Get-Mailbox -Identity $Ticket.NovoEmail
    
                    if($MailNameAlreadyExists){
                        
                        # Se o nome não está disponível dispara mensagem de erro e registra no banco de dados
                        $ErrorMsg = "Alias já está sendo utilizado por outro MailBox"
                        InsertLog($BotName, "Exception", "Ticket: $Solicitacao, $ErrorMsg")
                        InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMsg
                        Write-Output "Ticket: $Solicitacao || Erro: $ErrorMsg"
    
                    } else {
    
                        try {
                        
                            # Altera o E-mail
                            Set-MailBox -Identity $Ticket.EmailAtual -WindowsEmailAddress $Ticket.NovoEmail -ErrorAction Stop
                            Write-Output "Renomeou o Email"

                            $CloseNote = "Conforme solicitado, a caixa de email " + $Ticket.EmailAtual + " teve seu e-mail alterado para " + $Ticket.NovoEmail
                            $ItemDeCausa = "Outlook"
                            $Causa = "Alterar"
                            InsertCloseInfo -CloseNote $CloseNote -Solicitacao $Solicitacao -ItemDeCausa $ItemDeCausa -Causa $Causa

    
                            Write-Output "Ticket: $Solicitacao || Sucesso"
                            
                        }
                        catch {
                            # Insere a mensagem de erro no Banco de Dados
                            InsertLog($BotName, "Exception", "Ticket: $Solicitacao, $_")
                            InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $_
                            Write-Output "Ticket: $Solicitacao || Erro: $_"
                        }
                    }
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

