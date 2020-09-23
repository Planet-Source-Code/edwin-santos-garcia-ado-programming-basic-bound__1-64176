Attribute VB_Name = "modData"
Global adoConn As New ADODB.Connection
Global adoRS As New ADODB.Recordset
Global adoCmd As New ADODB.Command
Global strConn As String
Global strSQL As String

Global varName, varPass
Global varFound As Boolean
Global varCustNo

Sub OpenDatabase()
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    strConn = strConn & "FSS.mdb"
    
    adoConn.ConnectionString = strConn
    adoConn.Open
End Sub

Sub OpenEmployee()
    adoRS.Source = "Select * From Employee"
    adoRS.ActiveConnection = adoConn
    adoRS.Open
End Sub

Sub CloseDatabase()
    adoRS.Close
    Set adoRS = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub
