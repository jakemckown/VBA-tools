Attribute VB_Name = "MDatabase"
Option Explicit

Private pProvider As String
Private pDataSource As String
Private pDatabase As String
Private pUserID As String
Private pPassword As String

' Only dbOpenStatic supported for now
Public Enum dbCursorTypeEnum
    dbOpenDynamic = 2
    dbOpenForwardOnly = 0
    dbOpenKeyset = 1
    dbOpenStatic = 3
    dbOpenUnspecified = -1
End Enum

Public Enum dbDatabase
    dbMicrosoftAccessFileExample
    dbOracleDatabaseExample
    dbSQLServerDatabaseExample
End Enum

Public Function ExecuteQuery(ByVal Database As dbDatabase, _
                             ByVal Query As String) As Collection
    Dim Records As Collection
    Set Records = New Collection
    Call SetDatabase(Database)
    Dim Connection As Object
    Set Connection = CreateObject("ADODB.Connection")
    Dim Recordset As Object
    Set Recordset = CreateObject("ADODB.Recordset")
    On Error GoTo ErrorHandling
    With Connection
        Let .ConnectionString = vbNullString
        If pProvider <> vbNullString And pDataSource <> vbNullString Then
            Let .ConnectionString = .ConnectionString & "Provider=" & pProvider & ";"
            Let .ConnectionString = .ConnectionString & "Data Source=" & pDataSource & ";"
            If pDatabase <> vbNullString Then Let .ConnectionString = .ConnectionString & "Database=" & pDatabase & ";"
            If pUserID <> vbNullString Then Let .ConnectionString = .ConnectionString & "User ID=" & pUserID & ";"
            If pPassword <> vbNullString Then Let .ConnectionString = .ConnectionString & "Password=" & pPassword & ";"
        End If
        If .ConnectionString = vbNullString Then GoTo ErrorHandling
    End With
    Call Connection.Open
    Call Recordset.Open(Source:=Query, ActiveConnection:=Connection, CursorType:=dbOpenStatic)
    If Recordset.State <> 0 Then
        Call Recordset.MoveFirst
        Dim Record As Object, i As Long
        Do While Not Recordset.EOF
            Set Record = CreateObject("Scripting.Dictionary")
            For i = 0 To Recordset.Fields.Count - 1
                Call Record.Add(Key:=Recordset.Fields(i).Name, Item:=Recordset.Fields(i).Value)
            Next i
            Call Records.Add(Record)
            Call Recordset.MoveNext
        Loop
    End If
ErrorHandling:
    If Recordset.State <> 0 Then Call Recordset.Close
    If Connection.State <> 0 Then Call Connection.Close
    Set ExecuteQuery = Records
End Function

Private Sub SetDatabase(ByVal Database As dbDatabase)
    Select Case Database
        Case dbMicrosoftAccessFileExample
            Let pProvider = "Microsoft.ACE.OLEDB.12.0"
            Let pDataSource = "\\File\Path\To\MicrosoftAccessFile.accdb"
            Let pDatabase = vbNullString
            Let pUserID = "UserID" ' or vbNullString if no user ID
            Let pPassword = "Password" ' or vbNullString if no password
        Case dbOracleDatabaseExample
            Let pProvider = "MSDAORA"
            Let pDataSource = "DataSource"
            Let pDatabase = "Database"
            Let pUserID = "UserID" ' or vbNullString if no user ID
            Let pPassword = "Password" ' or vbNullString if no password
        Case dbSQLServerDatabaseExample
            Let pProvider = "SQLOLEDB"
            Let pDataSource = "DataSource"
            Let pDatabase = "Database"
            Let pUserID = "UserID" ' or vbNullString if no user ID
            Let pPassword = "Password" ' or vbNullString if no password
        Case Else
            Let pProvider = vbNullString
            Let pDataSource = vbNullString
            Let pDatabase = vbNullString
            Let pUserID = vbNullString
            Let pPassword = vbNullString
    End Select
End Sub
