[string]$corrid = Read-Host 'What is the Corelation ID to search for?'

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Data Source=EDM-GOA-SQL-240;Integrated Security=SSPI;Initial Catalog=master"
$SqlConnection.open()

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = "select [RowCreatedTime],[ProcessName],[Area],[Category],EventID,[Message] from [WSS_UsageApplication].[dbo].[ULSTraceLog] where CorrelationId=< pre>"+$corrid
$SqlCmd.Connection = $SqlConnection
$SqlCmd.CommandTimeout = 0

$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd

$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)

$SqlConnection.Close()
$DataSet.Tables[0]



<#
[string]$corrid = Read-Host 'What is the Corelation ID to search for?'
$connectionstring = "Data Source=EDM-GOA-SQL-240;Integrated Security=SSPI;Initial Catalog=master"
$connection = new-object system.data.SqlClient.SQLConnection($connectionstring);

[string]$qry ="select [RowCreatedTime],[ProcessName],[Area],[Category],EventID,[Message] from [WSS_UsageApplication].[dbo].[ULSTraceLog] where CorrelationId=< pre>"+$corrid

$cmd = new-object system.data.sqlclient.sqlcommand($qry, $connection);
$connection.Open();

$adapter = New-Object System.Data.sqlclient.sqlDataAdapter $cmd
$dataset = New-Object system.Data.DataSet

#$adapter.Fill($dataSet) | Out-Null
$adapter.Fill($dataset,"Tables")
$connection.Close()
$dataSet.Tables[0]
#>

