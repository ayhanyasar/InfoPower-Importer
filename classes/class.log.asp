<%
'#######################################################################
'
' InfoPower Importer PRO v3.0
' File : class.log.asp
' Description : Logging Engine
'
'#######################################################################

Option Explicit

'=======================================================================
' HTML HEADER
'=======================================================================

Sub Log_Start()

    If Not LOG_BROWSER Then Exit Sub

    Response.Write "<!doctype html>"
    Response.Write "<html>"
    Response.Write "<head>"
    Response.Write "<meta charset=""utf-8"">"
    Response.Write "<title>" & APP_NAME & " " & APP_VERSION & "</title>"

    Response.Write "<style>"
    Response.Write "body{margin:0;padding:20px;background:#f4f4f4;font-family:Segoe UI,Arial;font-size:13px;}"
    Response.Write ".panel{background:#fff;border:1px solid #ddd;padding:20px;}"
    Response.Write ".title{font-size:22px;font-weight:bold;margin-bottom:15px;}"
    Response.Write ".log{padding:4px 0;border-bottom:1px solid #eee;font-family:Consolas;font-size:12px;}"
    Response.Write ".ok{color:#008000;}"
    Response.Write ".warn{color:#d98300;}"
    Response.Write ".err{color:#cc0000;}"
    Response.Write ".info{color:#0066cc;}"
    Response.Write "</style>"

    Response.Write "</head>"
    Response.Write "<body>"
    Response.Write "<div class=""panel"">"
    Response.Write "<div class=""title"">"
    Response.Write APP_NAME & " v" & APP_VERSION
    Response.Write "</div>"

End Sub

'=======================================================================
' HTML FOOTER
'=======================================================================

Sub Log_Finish()

    If Not LOG_BROWSER Then Exit Sub

    Response.Write "</div>"
    Response.Write "</body>"
    Response.Write "</html>"

End Sub

'=======================================================================
' INFO
'=======================================================================

Sub Log_Info(Text)

    If Not LOG_BROWSER Then Exit Sub

    Response.Write "<div class=""log info"">"

    Response.Write Server.HTMLEncode(CStr(Text))

    Response.Write "</div>"

    Response.Flush

End Sub

'=======================================================================
' SUCCESS
'=======================================================================

Sub Log_Success(Text)

    If Not LOG_BROWSER Then Exit Sub

    Response.Write "<div class=""log ok"">"

    Response.Write "✔ "

    Response.Write Server.HTMLEncode(CStr(Text))

    Response.Write "</div>"

    Response.Flush

End Sub

'=======================================================================
' WARNING
'=======================================================================

Sub Log_Warning(Text)

    If Not LOG_BROWSER Then Exit Sub

    Response.Write "<div class=""log warn"">"

    Response.Write "⚠ "

    Response.Write Server.HTMLEncode(CStr(Text))

    Response.Write "</div>"

    Response.Flush

End Sub

'=======================================================================
' ERROR
'=======================================================================

Sub Log_Error(Text)

    ErrorProducts = ErrorProducts + 1

    If Not LOG_BROWSER Then Exit Sub

    Response.Write "<div class=""log err"">"

    Response.Write "✖ "

    Response.Write Server.HTMLEncode(CStr(Text))

    Response.Write "</div>"

    Response.Flush

End Sub

'=======================================================================
' DATABASE LOG
'=======================================================================

Sub Log_Database(Level,Message)

    If Not LOG_DATABASE Then Exit Sub

    Dim SQL

    SQL = ""

    SQL = SQL & "INSERT INTO " & TBL_LOGS & "("
    SQL = SQL & "log_date,"
    SQL = SQL & "log_level,"
    SQL = SQL & "log_message"
    SQL = SQL & ") VALUES ("

    SQL = SQL & "GETDATE(),"

    SQL = SQL & "'" & DB_String(Level) & "',"

    SQL = SQL & "'" & DB_String(Message) & "'"

    SQL = SQL & ")"

    DB_Execute SQL

End Sub

'=======================================================================
' BOTH
'=======================================================================

Sub Log_Write(Level,Message)

    Select Case UCase(Level)

        Case "INFO"

            Log_Info Message

        Case "SUCCESS"

            Log_Success Message

        Case "WARNING"

            Log_Warning Message

        Case "ERROR"

            Log_Error Message

        Case Else

            Log_Info Message

    End Select

    Log_Database Level,Message

End Sub

'#######################################################################
' END OF FILE
'#######################################################################