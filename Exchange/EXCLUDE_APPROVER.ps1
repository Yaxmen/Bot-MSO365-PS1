# Importa o módulo de conexão com o Banco de Dados
Import-Module "D:\Automation\Petro\TTR-Click\PowerShell\DatabaseModule.ps1"

# Recupera o tickets de acordo com o TicketType da demanda
$Tickets = GetWorkTickets -TicketType "EXCLUDE_APPROVER"

# Verifica se existem Tickets
if(!$Tickets){
    Write-Output "================================================================================="
    Write-output "Oferta: Incluir/Excluir aprovador de reserva de sala de reunião no Outlook"
    Write-Output "Ticket Type: EXCLUDE_APPROVER"
    Write-Output "Nenhum Ticket Localizado"
    Write-Output "=================================================================================="
    Write-Output " "
} else {

    # Percorre todos os chamados para o atendimento
    foreach($Ticket in $Tickets){

        try{
            # Resgata a Sala de Reunião a ser trabalhada
            $RoomMailBox = Get-Mailbox -Identity $Ticket.Of2_EmailDaSala -ErrorAction Stop

        } catch {
            InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg "MailBox não localizada no Exchange"
        }

        # Verifica se a Mailbox é elegível a atendimento
        if($RoomMailBox.DisplayName -match "Privativa"){

            # Verifica se existem chaves no campo de Chave
            if($Ticket.Of2_Chaves -eq ""){

                # Inserindo mensagem de erro na Base de Dados
                $ErrorMsg = "Nenhuma chave preenchida"
                InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMsg
                Write-Host "Erro: $ErrorMsg"

            } else {
                
                # Cria um Array com as chaves
                ArrayChaves = $Ticket.Of2_Chaves.Split(",")
                
                try {
                    # Recupera o Calendario da Sala
                    $Calendar = Get-CalendarProcessing -Identity $RoomMailBox.Alias -ErrorAction Stop

                    # Resgata o Resource Delegate
                    $Resource = $Calendar.ResourceDelegates

                    # Resgata o Book in Policy
                    $BookInPolicy = $Calendar.BookInPolicy
                }
                catch {
                    $ErrorMsg = "Erro ao recuperar o Calendário da Sala"
                    InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $_
                    Write-Host "Erro: $ErrorMsg"
                }
                
                # Percorre todas as chaves para realizar a ação
                foreach ($Chave in $ArrayChaves){
                    
                    # Captura a mailbox da chave
                    try{ $MailUser = Get-MailBox -Identity $Chave}
                    catch{
                        Write-Host "Não foi possível Localizar a chave $Chave no Exchange"
                    }
                    
                    # Recuperando o Alias da Caixa e o LegacyExchange
                    $AliasUser = $MailUser.Alias
                    $LegacyExchange = $MailUser.LegacyExchangeDN

                    # Verificando se o Usuário está cadastrado como aprovador
                    if(-not ($Resource.Contains($AliasUser)) -and -not ($BookInPolicy.Contains($LegacyExchange))){
                        
                        Write-Host "ATENÇÃO: Chave não localizada como aprovador da sala - $Chave"
                    } else {
                        
                        # Remove a chave da lista de aprovadores
                        $Resource.Remove($AliasUser)

                        # Remove a chave da lista do LegacyExchange
                        $BookInPolicy.Remove($LegacyExchange)

                    }

                }

                # Redefinindo a lista de aprovadores da sala de reunião
                Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource

                # Redefinindo lista de pré-aprovadores da sala
                Set-Mailbox -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy


                # Definindo as permissões de Calendário da sala
                $folder = $Ticket.NomeDaSala + ":\Calendar"

                try{ 
                    Set-MailboxFolderPermission -Identity $folder -User Default -AccessRights AvailabilityOnly | Out-Null
                } catch {
                    Write-Host "Erro ao atualizar o calendário"
                }

                # Informações de Fechamento do chamado
                $ItemDeCausa = "Outlook"
                $Item = "Outros"
                $CloseNote = "Conforme Solicitado as informações foram alteradas na sala de reunião: " + $Ticket.Of2_NomeDaSala

                # Inserindo as informações de fechamento 
                InsertCloseInfo -Solicitacao $Ticket.Solicitacao -ItemDeCausa $ItemDeCausa -Causa $Item -CloseNote $CloseNote
            }



        } else {
            # Inserindo as mensagens de erro no Sistema 
            $ErrorMsg = "Sala não localizada, ou não classificada como Privativa"
            InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMsg
            Write-Host "Erro: $ErrorMsg"
        }

    }
}

