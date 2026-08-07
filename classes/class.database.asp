<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.database.asp
' Part    : 1
'====================================================================
'Option Explicit

Class Database

    Private mConn
    Private mConnectionString
    Private mCommandTimeout
    Private mConnectionTimeout
    Private mLastError
    Private mLastSQL
    Private mAffectedRows
    Private mInTransaction

    Private Const adCmdText        = 1
    Private Const adOpenForwardOnly= 0
    Private Const adOpenStatic     = 3
    Private Const adLockReadOnly   = 1
    Private Const adLockOptimistic = 3
    Private Const adUseClient      = 3

    Private Sub Class_Initialize()

        mConnectionString  = ""
        mCommandTimeout    = 60
        mConnectionTimeout = 30
        mLastError         = ""
        mLastSQL           = ""
        mAffectedRows      = 0
        mInTransaction     = False

    End Sub

    Private Sub Class_Terminate()

        Close

    End Sub



    Public Property Let ConnectionString(Value)

        mConnectionString = CStr(Value)

    End Property



    Public Property Get ConnectionString()

        ConnectionString = mConnectionString

    End Property



    Public Property Let CommandTimeout(Value)

        mCommandTimeout = CLng(Value)

    End Property



    Public Property Get CommandTimeout()

        CommandTimeout = mCommandTimeout

    End Property



    Public Property Get LastError()

        LastError = mLastError

    End Property



    Public Property Get LastSQL()

        LastSQL = mLastSQL

    End Property



    Public Property Get AffectedRows()

        AffectedRows = mAffectedRows

    End Property



    Public Function Open()

        On Error Resume Next

        If IsObject(mConn) Then

            If mConn.State = 1 Then

                Open = True
                Exit Function

            End If

        End If

        Set mConn = Server.CreateObject("ADODB.Connection")

        mConn.ConnectionTimeout = mConnectionTimeout
        mConn.CommandTimeout    = mCommandTimeout

        mConn.Open mConnectionString

        If Err.Number <> 0 Then

            mLastError = Err.Description
            Open = False
            Err.Clear
            Exit Function

        End If

        Open = True

    End Function



    Public Sub Close()

        On Error Resume Next

        If IsObject(mConn) Then

            If mConn.State = 1 Then
                mConn.Close
            End If

            Set mConn = Nothing

        End If

        Err.Clear

    End Sub



    Public Function IsOpen()

        If IsObject(mConn) Then

            IsOpen = (mConn.State = 1)

        Else

            IsOpen = False

        End If

    End Function



    Public Function Execute(SQL)

        Dim Records

        Execute = False

        If Not Open() Then Exit Function

        mLastSQL = SQL
        mAffectedRows = 0

        On Error Resume Next

        mConn.Execute SQL, Records, adCmdText

        If Err.Number <> 0 Then

            mLastError = Err.Description
            Err.Clear
            Exit Function

        End If

        mAffectedRows = CLng(Records)

        Execute = True

    End Function



    Public Function Query(SQL)

        Dim RS

        If Not Open() Then

            Set Query = Nothing
            Exit Function

        End If

        mLastSQL = SQL

        Set RS = Server.CreateObject("ADODB.Recordset")

        RS.CursorLocation = adUseClient

        On Error Resume Next

        RS.Open SQL, _
                mConn, _
                adOpenStatic, _
                adLockReadOnly, _
                adCmdText

        If Err.Number <> 0 Then

            mLastError = Err.Description

            Set RS = Nothing

            Err.Clear

        End If

        Set Query = RS

    End Function
 
     Public Function Scalar(SQL)

        Dim RS

        Scalar = Null

        Set RS = Query(SQL)

        If IsObject(RS) Then

            If Not RS.EOF Then
                Scalar = RS(0).Value
            End If

            RS.Close
            Set RS = Nothing

        End If

    End Function



    Public Function GetRow(SQL)

        Dim RS
        Dim D
        Dim i

        Set D = Server.CreateObject("Scripting.Dictionary")
        D.CompareMode = 1

        Set RS = Query(SQL)

        If IsObject(RS) Then

            If Not RS.EOF Then

                For i = 0 To RS.Fields.Count - 1
                    D.Add RS.Fields(i).Name, RS.Fields(i).Value
                Next

            End If

            RS.Close
            Set RS = Nothing

        End If

        Set GetRow = D

    End Function



    Public Function GetRows(SQL)

        Dim RS
        Dim Rows()
        Dim Row
        Dim i
        Dim n

        n = -1

        ReDim Rows(-1)

        Set RS = Query(SQL)

        If IsObject(RS) Then

            Do Until RS.EOF

                n = n + 1

                ReDim Preserve Rows(n)

                Set Row = Server.CreateObject("Scripting.Dictionary")
                Row.CompareMode = 1

                For i = 0 To RS.Fields.Count - 1
                    Row.Add RS.Fields(i).Name, RS.Fields(i).Value
                Next

                Set Rows(n) = Row

                RS.MoveNext

            Loop

            RS.Close
            Set RS = Nothing

        End If

        GetRows = Rows

    End Function



    Public Function RecordExists(SQL)

        Dim RS

        RecordExists = False

        Set RS = Query(SQL)

        If IsObject(RS) Then

            RecordExists = Not RS.EOF

            RS.Close
            Set RS = Nothing

        End If

    End Function



    Public Function BeginTransaction()

        If Not Open() Then
            BeginTransaction = False
            Exit Function
        End If

        On Error Resume Next

        mConn.BeginTrans

        If Err.Number <> 0 Then

            mLastError = Err.Description
            BeginTransaction = False

            Err.Clear

            Exit Function

        End If

        mInTransaction = True

        BeginTransaction = True

    End Function



    Public Function Commit()

        Commit = False

        If Not mInTransaction Then Exit Function

        On Error Resume Next

        mConn.CommitTrans

        If Err.Number <> 0 Then

            mLastError = Err.Description

            Err.Clear

            Exit Function

        End If

        mInTransaction = False

        Commit = True

    End Function



    Public Function Rollback()

        Rollback = False

        If Not mInTransaction Then Exit Function

        On Error Resume Next

        mConn.RollbackTrans

        If Err.Number <> 0 Then

            mLastError = Err.Description

            Err.Clear

            Exit Function

        End If

        mInTransaction = False

        Rollback = True

    End Function
 
     Public Function Escape(Value)

        If IsNull(Value) Then

            Escape = "NULL"

        Else

            Escape = "'" & Replace(CStr(Value), "'", "''") & "'"

        End If

    End Function



    Public Function EscapeLike(Value)

        Dim S

        S = CStr(Value)

        S = Replace(S, "'", "''")
        S = Replace(S, "%", "[%]")
        S = Replace(S, "_", "[_]")

        EscapeLike = S

    End Function



    Public Function InsertID()

        InsertID = Scalar("SELECT SCOPE_IDENTITY()")

    End Function



    Public Function ServerDate()

        ServerDate = Scalar("SELECT GETDATE()")

    End Function



    Public Function TableExists(TableName)

        TableExists = RecordExists( _
            "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME=" & Escape(TableName))

    End Function



    Public Function ColumnExists(TableName, ColumnName)

        ColumnExists = RecordExists( _
            "SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS " & _
            "WHERE TABLE_NAME=" & Escape(TableName) & _
            " AND COLUMN_NAME=" & Escape(ColumnName))

    End Function



    Public Function Connection()

        If Open() Then
            Set Connection = mConn
        Else
            Set Connection = Nothing
        End If

    End Function



    Public Function RecordCount(SQL)

        Dim RS

        RecordCount = 0

        Set RS = Query(SQL)

        If IsObject(RS) Then

            If Not RS.EOF Then

                RS.MoveLast
                RecordCount = RS.RecordCount

            End If

            RS.Close
            Set RS = Nothing

        End If

    End Function



    Public Function DatabaseName()

        DatabaseName = Scalar("SELECT DB_NAME()")

    End Function



    Public Function DatabaseVersion()

        DatabaseVersion = Scalar("SELECT @@VERSION")

    End Function



    Public Function TestConnection()

        TestConnection = Open()

    End Function



End Class
%>