Attribute VB_Name = "modDB"
Public Conn  As New adodb.Connection
Public dbPath As String


'����ACCESS���ݿ�
Sub DBConnect()
    strconn = "Provider=Microsoft.Jet.OLEDB.4.0;jet oledb:database Password=office;Data Source=" & dbPath
    If Conn.State <> 0 Then Conn.Close
    Conn.Open strconn
    
End Sub
