# Importa o módulo de conexão com o Banco de Dados
Import-Module "D:\Automation\Petro\TTR-Click\PowerShell\DatabaseModule.ps1"

# Recupera o tickets de acordo com o TicketType da demanda
$Tickets = GetWorkTickets -TicketType "INCLUDE_APPROVER"

# Verifica se existem Tickets
if(!$Tickets){
    Write-Output "================================================================================="
    Write-output "Oferta: Incluir/Excluir aprovador de reserva de sala de reunião no Outlook"
    Write-Output "Ticket Type: INCLUDE_APPROVER"
    Write-Output "Nenhum Ticket Localizado"
    Write-Output "=================================================================================="
    Write-Output " "
} else {

    # Percorre todos os tickets para fazer o atendimento 
    foreach($Ticket in $Tickets){
        
        # Verifica se a mailbox a ser trabalhada existe
        try{
            $RoomMailBox = Get-Mailbox -Identity $Ticket.Of2_EmailDaSala -ErrorAction Stop
        } catch {
            $ErrorMessage = "MailBox não encontrada no Exchange"
            InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMessage 
            Write-Host $ErrorMessage
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

                try {

                    # Resgatando o calendário
                    $Calendar = Get-CalendarProcessing -Identity $RoomMailBox.Alias -ErrorAction Stop

                    # Capturando o Resource Delegate
                    $Resource = $Calendar.ResourceDelegates

                    # Capturando o Book in Policy
                    $BookInPolicy = $Calendar.BookInPolicy
                }
                catch {
                    $ErrorMessage = "Erro ao Resgatar dados do calendário"
                    InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMessage
                    Write-Host $Ticket.Solicitacao " || Erro: $ErrorMessage"
                }

                # Verificando se a sala aceita aprovadores
                if (($Calendar.AutomateProcessing -eq "AutoAccept") -and (-not($Calendar.AllBookInPolicy)) -and ($Calendar.AllRequestInPolicy)){
                    Continue
                } else {
                    
                    try {
                        # Setando a sala para aceitar apovadores
                        Set-CalendarProcessing -Identity $RoomMailBox.Alias -AutomateProcessing AutoAccept -AllBookInPolicy $false -AllRequestInPolicy $true -ErrorAction Stop
                    }
                    catch {
                        $ErrorMessage = "Não foi possível configurar a sala para aceitar aprovadores"
                        InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMessage
                        Write-Host $Ticket.Solicitacao + " || Erro: $ErrorMessage"
                    }
                }

                # Criando um array com as chaves
                $ArrayChaves = $Ticket.Chaves.Split(",")

                # Percorrendo todas as chaves para fazer a inclusão
                foreach($Chave in $ArrayChaves){
                    
                    
                    # Captura a mailbox da chave
                    try{ $MailUser = Get-MailBox -Identity $Chave}
                    catch{
                        $ErrorMessage = "Chave não localizada $Chave"
                        Write-Host "Erro não fatal: $ErrorMessage"
                        Continue
                    }

                    # Capturando o LegacyExchange e o Alias do usuario
                    $LegacyExchange = $MailUser.LegacyExchangeDN
                    $AliasUser = $MailUser.Alias


                    # Verifica se a Chave ja esta cadastrada como Aprovador
                    if(($Resource.Contains($AliasUser)) -and ($BookInPolicy.Contains($LegacyExchange))) {
                        Continue
                    } else {

                        # Incluindo como aprovador da Sala pelo Alias
                        $Resource.Add($AliasUser)

                        # Incluindo como Aprovador da Sala pelo LegacyExchange
                        $BookInPolicy.Add($LegacyExchange)

                    }

                    # Definindo a Folder do calendário 
                    $Folder = $RoomMailBox.DisplayName + ":\Calendar"

                    # Setando cada usuário como aprovador
                    Set-MailboxFolderPermission $Folder -User $AliasUser -AccessRights Editor -SharingPermissionFlags Delegate | Out-Null
                }
                
                try {
                    
                    # Definindo lista de aprovadores
                    Set-Mailbox -Identity $RoomMailBox.Alias -GrantSendOnBehalfTo $Resource
                    
                    # Defininfo Lista de pré-Aprovadores
                    Set-Mailbox -Identity $RoomMailBox.Alias -BookInPolicy $BookInPolicy
                }
                catch {
                    $ErrorMessage = "Erro ao atualizar a Lista de Aprovadores"
                    InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMessage
                    Write-Outline $Ticket.Solicitaco + " || Erro: $ErrorMessage"
                }
            }

            $Folder = $Ticket.NomeDaSala + ":\Calendar"

            try {
                
                # Configurar AvailabilityOnly para Usuario DEFAULT
                Set-MailboxFolderPermission $Folder -User Default -AccessRights AvailabilityOnly | Out-Null -ErrorAction Stop
            }
            catch {
                $ErrorMessage = "Erro ao definir AvailabilityOnly"
                InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMessage
                Write-Outline $Ticket.Solicitacao + " || Erro: $ErrorMessage"
            }

            # Defininfo variáveis de fechamento do chamado
            $CloseNote = "Conforme Solicitado as chaves informadas foram adicionadas como aprovadoras da sala " + $Ticket.Of2_NomeDaSala
            $Item = "Outros"
            $ItemDeCausa = "Outlook"

            InsertCloseInfo -Solicitacao $Ticket.Solicitacao -CloseNote $CloseNote -ItemDeCausa $ItemDeCausa -Item $Item
            
        } else {
            $ErrorMessage = "Sala não privativa"
            InsertErrorMessage -Solicitacao $Ticket.Solicitacao -ErrorMsg $ErrorMessage
            Print("Erro: $ErrorMessage")
        }

    }
}