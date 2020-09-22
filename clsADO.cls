VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'   Connection object
    Private adocn As New ADODB.Connection
'   Declare the recordset objects
    Private adoRs As New ADODB.Recordset
'   Private collection to hold table names
    Private colDBTables As New Collection
    
    Private mConnectionString$
Public Sub GetTables(Optional StrTable As String, Optional Connection$, Optional PWD$, Optional USERID$, Optional CloseTblRs As Integer)
'   If the connection is no longer needed close it
    If CloseTblRs = 0 Then
    
'   Close the connection
            adoRs.Close
    Else
'   If the connection is open then skip the open statement
        If adoRs.State <> adStateOpen Then
                  adocn.Open Connection$
'   Use OpenSchema to retrieve the database's table names
                Set adoRs = adocn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
'   Move to the first table in the database
                adoRs.MoveFirst
        End If
            Do Until adoRs.EOF
        
'   Add Tables to the collection
                addTable = adoRs!Table_Name
'   Move to the next table name
                adoRs.MoveNext
            Loop
End If
End Sub

Private Property Let addTable(TableName$)
     colDBTables.Add TableName$
End Property
    
Public Property Get TableNames() As Collection
   Set TableNames = colDBTables
End Property

Public Property Let ConnectionString(ConnString$)
    mConnectionString = ConnString$
End Property
Public Property Get ConnectionString$()
    ConnectionString$ = mConnectionString$
End Property

Private Sub Class_Terminate()
'   Release objects
    Set adocn = Nothing
    Set adoRs = Nothing
End Sub

