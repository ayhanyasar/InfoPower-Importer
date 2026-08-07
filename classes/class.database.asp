<%
'#######################################################################
'
' InfoPower Importer PRO v3.0
' File : class.database.asp
' Description : Database Layer
'
'#######################################################################

Option Explicit

Dim Conn
Dim ConnStr

'=======================================================================
' OPEN CONNECTION
'=======================================================================

Sub DB_Open()

    If IsObject(Conn) Then Exit Sub

    ConnStr = Trim(Application("ConnectionString"))

    If ConnStr = "" Then

        Err.Raise vbObjectError + 1000, _
                  "DB_Open", _
                  "Application(""ConnectionString"") boş."

    End If

    Set Conn = Server.CreateObject("ADODB.Connection")

    Conn.ConnectionTimeout = 30
    Conn.CommandTimeout = 600

    Conn.Open ConnStr

End Sub

'=======================================================================
' CLOSE CONNECTION
'=======================================================================

Sub DB_Close()

    On Error Resume Next

    If IsObject(Conn) Then

        If Conn.State = 1 Then
            Conn.Close
        End If

        Set Conn = Nothing

    End If

End Sub

'=======================================================================
' BEGIN TRANSACTION
'=======================================================================

Sub DB_Begin()

    If Conn.State = 1 Then
        Conn.BeginTrans
    End If

End Sub

'=======================================================================
' COMMIT
'=======================================================================

Sub DB_Commit()

    If Conn.State = 1 Then
        Conn.CommitTrans
    End If

End Sub

'=======================================================================
' ROLLBACK
'=======================================================================

Sub DB_Rollback()

    If Conn.State = 1 Then
        Conn.RollbackTrans
    End If

End Sub

'=======================================================================
' EXECUTE SQL
'=======================================================================

Sub DB_Execute(SQL)

    Conn.Execute SQL

End Sub

'=======================================================================
' GET RECORDSET
'=======================================================================

Function DB_Recordset(SQL)

    Dim RS

    Set RS = Server.CreateObject("ADODB.Recordset")

    RS.CursorLocation = 3

    RS.Open SQL, _
            Conn, _
            0, _
            1

    Set DB_Recordset = RS

End Function

'=======================================================================
' GET FIRST VALUE
'=======================================================================

Function DB_Value(SQL)

    Dim RS

    DB_Value = Null

    Set RS = Conn.Execute(SQL)

    If Not RS.EOF Then
        DB_Value = RS(0)
    End If

    RS.Close
    Set RS = Nothing

End Function

'=======================================================================
' EXISTS
'=======================================================================

Function DB_Exists(SQL)

    Dim RS

    DB_Exists = False

    Set RS = Conn.Execute(SQL)

    If Not RS.EOF Then
        DB_Exists = True
    End If

    RS.Close
    Set RS = Nothing

End Function

'=======================================================================
' LAST IDENTITY
'=======================================================================

Function DB_LastIdentity()

    DB_LastIdentity = CLng(DB_Value("SELECT @@IDENTITY"))

End Function

'=======================================================================
' SAFE STRING
'=======================================================================

Function DB_String(Value)

    If IsNull(Value) Then

        DB_String = ""

        Exit Function

    End If

    DB_String = Replace(CStr(Value), "'", "''")

End Function

'=======================================================================
' SAFE NUMBER
'=======================================================================

Function DB_Number(Value)

    If IsNull(Value) Then

        DB_Number = "NULL"

        Exit Function

    End If

    If Trim(CStr(Value)) = "" Then

        DB_Number = "NULL"

        Exit Function

    End If

    If IsNumeric(Value) Then

        DB_Number = Replace(CStr(Value), ",", ".")

    Else

        DB_Number = "NULL"

    End If

End Function

'=======================================================================
' SAFE DATE
'=======================================================================

Function DB_Date(Value)

    If IsDate(Value) Then

        DB_Date = "'" & _
                  Year(Value) & "-" & _
                  Right("0" & Month(Value),2) & "-" & _
                  Right("0" & Day(Value),2) & " " & _
                  Right("0" & Hour(Value),2) & ":" & _
                  Right("0" & Minute(Value),2) & ":" & _
                  Right("0" & Second(Value),2) & "'"

    Else

        DB_Date = "NULL"

    End If

End Function

'=======================================================================
' CONNECTION TEST
'=======================================================================

Function DB_Test()

    On Error Resume Next

    DB_Test = False

    DB_Open

    If Err.Number = 0 Then
        DB_Test = True
    End If

    Err.Clear

End Function

'#######################################################################
' END OF FILE
'#######################################################################