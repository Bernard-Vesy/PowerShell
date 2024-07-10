# Définir les paramètres de connexion
$dsn = "mtaprdmd_read" # Nom du DSN ODBC configuré
$user = "sysprogress"
$password = "sysprogress"

# Construire la chaîne de connexion
$connectionString = "DSN=$dsn;UID=$user;PWD=$password;"

# Créer et ouvrir la connexion ODBC
$connection = New-Object System.Data.Odbc.OdbcConnection
$connection.ConnectionString = $connectionString
$connection.Open()

# Définir la requête SQL à exécuter
$query = "SELECT * FROM pub.drwphys where drwphys_id > '20240000000000000' "
                                                       
#$query = "SELECT * FROM pub.drwphys"

# Créer une commande ODBC
$command = $connection.CreateCommand()
$command.CommandText = $query

# Exécuter la requête et récupérer les résultats
$adapter = New-Object System.Data.Odbc.OdbcDataAdapter $command
$dataSet = New-Object System.Data.DataSet
$adapter.Fill($dataSet)

# Afficher les résultats
#$dataSet.Tables[0] | Format-Table -AutoSize
$dataSet.Tables[0].Rows[1].drwphys_id

# Parcourir et afficher les résultats de 1 à X
for ($i = 0; $i -lt $dataSet.Tables[0].Rows.Count; $i++) {

    Write-Output "$($dataSet.Tables[0].Rows[$i].drwphys_name)  $($dataSet.Tables[0].Rows[$i].drwphys_id)"
}


# Fermer la connexion
$connection.Close()
