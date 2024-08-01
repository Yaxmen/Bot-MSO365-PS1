

#Parametros de conexão
$SQLDBName = "TTR_CLICK"
$BotName = "TTR_Click"


function GetConnection {
    
    try {
        
        # Abrindo Conexão com o Banco de dados
        $connection=New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString ="server='localhost\SQLEXPRESS';database='$SQLDBName';trusted_connection=true"
        $connection.Open()
    
        return $connection
    }
    catch {
        InsertLog($BotName, "Exception", $_)
        Write-Output("[" +$currentQuery + "]  " + $_)
    }
    
}

#GetWorkTickets(): Método responsável por buscar os Tickets que serão trabalhados
function GetWorkTickets($TicketType){
    try{

        # Abrindo Conexao
        $connection = GetConnection

        # Criando Query para executar proc
        $query = "EXEC sp_GetWorkTicket '$TicketType'"
        $currentQuery = $query
        $command=$connection.CreateCommand()
        $command.CommandText =$query
        $result=$command.ExecuteReader() 
   
        #Guardando resultado da coleta de dados em uma tabela 
        $table=New-Object -TypeName System.Data.DataTable
        $table.Load($result)

        #Fechando coexão e retornando o resultado
        $connection.Close()
        return $table

    } catch {
        InsertLog($BotName, "Exception", $_)
        Write-Output("[" +$currentQuery + "]  " + $_)
    }
}

function InsertErrorMessage ($Solicitacao, $ErrorMsg){
    try {
        # Abrindo Conexao
        $connection = GetConnection
    
        # Criando Query para executar proc
        $query = "sp_InsertErrorMessage '$Solicitacao', '$ErrorMsg'"
        $currentQuery = $query
        $command=$connection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteReader()
        
        $connection.Close()
        
    }
    catch {
        InsertLog($BotName, "Exception", $_)
        Write-Output("[" +$currentQuery + "]  " + $_)
    }

}

function InsertCloseInfo ($Solicitacao, $CloseNote, $ItemDeCausa, $Causa) {
    try {
        # Abrindo Conexao
        $connection = GetConnection
    
        # Criando Query para executar proc
        $query = "sp_InsertCloseInfo '$Solicitacao', '$CloseNote', '$ItemDeCausa', '$Causa'"
        $currentQuery = $query
        $command=$connection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteReader()
        
        $connection.Close()
        
    }
    catch {
        InsertLog($BotName, "Exception", $_)
        Write-Output("[" +$currentQuery + "]  " + $_)
    }
}

function InsertLog($BotName, $Type, $Message) {
    
    try {
        # Abrindo Conexao
        $connection=New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString ="server='localhost\SQLEXPRESS';database='Petro';trusted_connection=true"
        $connection.Open()
    
        # Criando Query para executar proc
        $query = "sp_InsertLog '$BotName', '$Type', '$Message'"
        $currentQuery = $query
        $command=$connection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteReader()
        
        $connection.Close()
        
    }
    catch {
        Write-Output("[" +$currentQuery + "]  " + $_)
    }

}