VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADOFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function OpenSQLForwardOnly(sql As String, Cnx As String) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    Dim cn As ADODB.Connection

    On Error GoTo errhandler
    Set cn = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    cn.Open Cnx
    Rs.Open sql, cn, 3, 1
    Set OpenSQLForwardOnly = Rs

    Exit Function

errhandler:
    Set OpenSQLForwardOnly = Nothing
    MsgBox Err.Number & " " & Err.Description, vbExclamation, App.Title
End Function
