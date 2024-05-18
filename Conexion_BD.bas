Attribute VB_Name = "Conexion_BD"
Public Conn_Mysqldb As ADODB.Connection
Public Conn_MySqlRemoteBD As ADODB.Connection
Public Conn_MysqlLocalBD As ADODB.Connection
Public Conn_MysqldbAlmRep As ADODB.Connection
Public cmd_mysqldb As New ADODB.Command
Public rst_mysqldb As ADODB.Recordset
Public fld_mysqldb As ADODB.Field
Public sql_mysqldb As String
Public Sub Conn_BDaiosoft()
'connect to MySQL server using Connector/ODBC
Set Conn_Mysqldb = New ADODB.Connection
ConnBDCorrecta = True
On Error GoTo ErrorConexion
'Conn_Mysqldb.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};SERVER=localhost;DATABASE=bdaiosoft;UID=root;PWD=Aio2019Pass; OPTION=3"
Conn_Mysqldb.ConnectionString = "DSN=ConexionAiosoft;UID=root;PWD=Aio2019Pass;"
'Conn_Mysqldb.ConnectionString = "DSN=ConexionAiosoft;UID=aiosoftbd;PWD=yyphitgw3EVnmDt;"
Conn_Mysqldb.Open
Exit Sub
ErrorConexion:
If Err.Number = -2147467259 Then
    ConnBDCorrecta = False
    
End If

Resume Next

End Sub

Public Sub Conn_Remote_BDaiosfot()
Set Conn_MySqlRemoteBD = New ADODB.Connection
ConnBDCorrecta = True
On Error GoTo ErrorConexion
Conn_MySqlRemoteBD.ConnectionString = "DSN=ConexionPadronSunat;UID=aiosoftbd;PWD=yyphitgw3EVnmDt;"
Conn_MySqlRemoteBD.ConnectionTimeout = 30
Conn_MySqlRemoteBD.Open
Exit Sub

ErrorConexion:
If Err.Number = -2147467259 Then
    ConnBDCorrecta = False
End If

Resume Next
End Sub
Public Sub Conn_Local_BDaiosoft()
'connect to MySQL server using Connector/ODBC
Set Conn_MysqlLocalBD = New ADODB.Connection
ConnBDCorrecta = True
On Error GoTo ErrorConexion
'Conn_Mysqldb.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};SERVER=localhost;DATABASE=bdaiosoft;UID=root;PWD=Aio2019Pass; OPTION=3"
Conn_MysqlLocalBD.ConnectionString = "DSN=LocalCliente;UID=aiosoftbd;PWD=yyphitgw3EVnmDt;"
'Conn_MysqlLocalBD.ConnectionString = "DSN=LocalCliente;UID=root;PWD=Aio2019Pass;"

Conn_MysqlLocalBD.Open
Exit Sub
ErrorConexion:
If Err.Number = -2147467259 Then
    ConnBDCorrecta = False
End If

Resume Next
End Sub

Public Sub ConectarAlmacenRepuesto()
Set Conn_MysqldbAlmRep = New ADODB.Connection
ConnBDCorrecta = True
On Error GoTo ErrorConexion
Conn_MysqldbAlmRep.ConnectionString = "DSN=ConnAlmacenRep;UID=aiosoftbd;PWD=yyphitgw3EVnmDt;"
Conn_MysqldbAlmRep.Open
Exit Sub
ErrorConexion:
If Err.Number = -2147467259 Then
    ConnBDCorrecta = False
End If
Resume Next
    
End Sub
