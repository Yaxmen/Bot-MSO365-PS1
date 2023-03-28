
Remove-Variable * -ErrorAction SilentlyContinue

$credentials = @{}
$SQLServer = "avareport.westeurope.cloudapp.azure.com"
$SQLDBName = "Automation_TTR_Vale"
#$uid ="Gasbarro"
#$pwd = "@avanade7411"
$uid ="amaurilo"
$pwd = "@amaurilo2020"
$currentQuery = "";
#$credentials.add('Server','$SQLServer)



function GetTickets($ticketType)
{
    #Trazendo os chamados do banco de dados
   
   try
   {
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString ="server=$SQLServer,1433;database=$SQLDBName;uid=$uid;password=$pwd"
        $connection.Open()

        #creating and running the query
        $query = "EXEC sp_GetTickets '$ticketType'"
        $currentQuery = $query
        $command=$connection.CreateCommand()
        $command.CommandText =$query
        $result=$command.ExecuteReader() 
 
        #store result in the Table variable
        $table=New-Object -TypeName System.Data.DataTable
        $table.Load($result)

        $connection.Close()

        return $table
    }
    catch 
    {
        Write-Output("[" +$currentQuery + "]  " + $_)
    }
}

function UpdateTicket($WO, $Status, $errorMessage,  $closeNote)
{
    #Atualizando os chamados do banco de dados
    $currentQuery = "";
   
    $connection=New-Object -TypeName System.Data.SqlClient.SqlConnection
    $connection.ConnectionString ="server=$SQLServer,1433;database=$SQLDBName;uid=$uid;password=$pwd"
    $connection.Open()
    try
    {
        #creating and running the query
   
     if($null -eq $errorMessage)
     {
        $errorMessage = " "
     }
     if($null -eq $closeNote)
     {
        $closeNote = " "
     }
        $closeNote.Replace("'"," ")
        $query = "EXEC sp_UpdateWorkTickets '$WO', '$Status', '$errorMessage', '$closeNote'"
        $currentQuery = $query
        $command=$connection.CreateCommand()
        $command.CommandText =$query
        $command.ExecuteNonQuery() 
        $connection.Close()   
    }
    catch  [Exception]
    {
        Write-Output("[" +$currentQuery + "]  " + $_.Exceptption.Message + " " +$WO)
        Log($WO,"'$WO', '$Status', '$errorMessage', '$ticketType', '$CloseNote'")
    }
}

function InsertProvisioningCheck($WO, $UPN, $site, $ProvisioningDate)
{
    try
    {

        $connection=New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString ="server=$SQLServer,1433;database=$SQLDBName;uid=$uid;password=$pwd"
        $connection.Open()

        #creating and running the query
        $query = "EXEC sp_InsertProvisioningCheck '$WO', '$UPN', '$site', '$ProvisioningDate'"
        $currentQuery = $query
        $command=$connection.CreateCommand()
        $command.CommandText =$query
        $command.ExecuteNonQuery() 
 
        #store result in the Table variable
        #$table=New-Object -TypeName System.Data.DataTable
        #$table.Load($result)

        $connection.Close()
    }
    catch
    {
        Write-Output("[" +$currentQuery + "]  " + $_)
    }

}
function MoveTicket($WO)
{
    #Movendo os chamados do banco de dados
    try
    {

        $connection=New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString ="server=$SQLServer,1433;database=$SQLDBName;uid=$uid;password=$pwd"
        $connection.Open()

        #creating and running the query
        $query = "EXEC sp_MoveTickets '$WO'"
        $currentQuery = $query
        $command=$connection.CreateCommand()
        $command.CommandText =$query
        $command.ExecuteNonQuery() 
 
        #store result in the Table variable
        $table=New-Object -TypeName System.Data.DataTable
        $table.Load($result)

        $connection.Close()
    }
    catch
    {
        Write-Output("[" +$currentQuery + "]  " + $_)
    }
}


function Log($WO,$MSG)
{
  

    $connection=New-Object -TypeName System.Data.SqlClient.SqlConnection
    $connection.ConnectionString ="server=$SQLServer,1433;database=$SQLDBName;uid=$uid;password=$pwd"
    $connection.Open()


    #creating and running the query
    $query = "EXEC sp_LOG '$WO','Vale','TTR_VALE','ERROR','$MSG'"
    $command=$connection.CreateCommand()
    $command.CommandText =$query
    $command.ExecuteNonQuery() 


    $connection.Close()
}