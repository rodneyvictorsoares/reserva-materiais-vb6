Attribute VB_Name = "Module1"
Public gstrUsuarioLogado As String
Public gintIdUsuarioLogado As Integer
Public gstrTipoUsuario As String
Public gConn As ADODB.Connection

Public Sub AbrirConexao()
    Set gConn = New ADODB.Connection
    gConn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=school_db;Data Source=.\SQLEXPRESS"
    gConn.Open
End Sub

Public Sub FecharConexao()
    gConn.Close
    Set gConn = Nothing
End Sub
