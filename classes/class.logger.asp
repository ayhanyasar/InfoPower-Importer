<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.logger.asp
' Part    : 1
'====================================================================
'Option Explicit

Class Logger

    Private Const LOG_DEBUG = 1
    Private Const LOG_INFO  = 2
    Private Const LOG_WARN  = 3
    Private Const LOG_ERROR = 4
    Private Const LOG_FATAL = 5

    Private mFSO
    Private mLogFolder
    Private mLogFile
    Private mMinLevel
    Private mEnabled
    Private mConsole
    Private mStartTick
    Private mDateFormat
    Private mFilePrefix
    Private mMaxSizeMB
    Private mKeepDays

    Private Sub Class_Initialize()

        Set mFSO = Server.CreateObject("Scripting.FileSystemObject")

        mLogFolder = Server.MapPath("/logs")
        mFilePrefix = "importer"

        mEnabled = True
        mConsole = False

        mMinLevel = LOG_DEBUG

        mDateFormat = "yyyy-mm-dd hh:nn:ss"

        mMaxSizeMB = 20
        mKeepDays = 30

        EnsureFolder

        mLogFile = BuildLogFile()

        mStartTick = Timer()

    End Sub

    Private Sub Class_Terminate()

        Set mFSO = Nothing

    End Sub



    Public Property Let Enabled(Value)

        mEnabled = CBool(Value)

    End Property



    Public Property Get Enabled()

        Enabled = mEnabled

    End Property



    Public Property Let MinimumLevel(Value)

        mMinLevel = CLng(Value)

    End Property



    Public Property Get MinimumLevel()

        MinimumLevel = mMinLevel

    End Property



    Public Property Let Folder(Value)

        mLogFolder = CStr(Value)

        EnsureFolder

        mLogFile = BuildLogFile()

    End Property



    Public Property Get Folder()

        Folder = mLogFolder

    End Property



    Public Property Let Prefix(Value)

        mFilePrefix = Trim(CStr(Value))

        mLogFile = BuildLogFile()

    End Property



    Public Property Get Prefix()

        Prefix = mFilePrefix

    End Property



    Public Property Let MaxSizeMB(Value)

        mMaxSizeMB = CLng(Value)

    End Property



    Public Property Get MaxSizeMB()

        MaxSizeMB = mMaxSizeMB

    End Property



    Public Property Let KeepDays(Value)

        mKeepDays = CLng(Value)

    End Property



    Public Property Get KeepDays()

        KeepDays = mKeepDays

    End Property



    Public Sub Debug(Message)

        WriteLog LOG_DEBUG,"DEBUG",Message

    End Sub



    Public Sub Info(Message)

        WriteLog LOG_INFO,"INFO",Message

    End Sub



    Public Sub Warn(Message)

        WriteLog LOG_WARN,"WARN",Message

    End Sub



    Public Sub [Error](Message)

        WriteLog LOG_ERROR,"ERROR",Message

    End Sub



    Public Sub Fatal(Message)

        WriteLog LOG_FATAL,"FATAL",Message

    End Sub



    Public Sub Separator()

        WriteRaw String(100,"-")

    End Sub



    Public Sub Blank()

        WriteRaw ""

    End Sub



    Public Sub SessionStart()

        Separator
        Info "IMPORT SESSION STARTED"
        Info "DATE : " & Now()
        Separator

    End Sub



    Public Sub SessionEnd()

        Separator
        Info "IMPORT SESSION FINISHED"
        Info "ELAPSED : " & FormatNumber(Timer()-mStartTick,2) & " sec"
        Separator
        Blank

    End Sub
 
     Public Sub WriteException(ProcedureName, Number, Description)

        Dim Msg

        Msg = ""
        Msg = Msg & "Procedure : " & ProcedureName & vbCrLf
        Msg = Msg & "Number    : " & Number & vbCrLf
        Msg = Msg & "Message   : " & Description

        WriteLog LOG_ERROR, "EXCEPTION", Msg

    End Sub



    Public Sub WriteSQL(SQL)

        WriteLog LOG_DEBUG, "SQL", SQL

    End Sub



    Public Sub WriteURL(URL)

        WriteLog LOG_INFO, "HTTP", URL

    End Sub



    Public Sub WriteFile(FileName)

        WriteLog LOG_INFO, "FILE", FileName

    End Sub



    Public Sub Rotate()

        Dim F

        Dim MaxBytes

        MaxBytes = CLng(mMaxSizeMB) * 1024 * 1024

        If Not mFSO.FileExists(mLogFile) Then Exit Sub

        Set F = mFSO.GetFile(mLogFile)

        If CLng(F.Size) < MaxBytes Then Exit Sub

        On Error Resume Next

        mFSO.MoveFile _
            mLogFile, _
            Replace(mLogFile, ".log", "_" & _
            Replace(Replace(Replace(CStr(Now()),":",""),"/","")," ","_") & ".log")

        Err.Clear

        mLogFile = BuildLogFile()

    End Sub



    Public Sub Cleanup()

        Dim Folder
        Dim File

        Set Folder = mFSO.GetFolder(mLogFolder)

        For Each File In Folder.Files

            If LCase(Right(File.Name,4)) = ".log" Then

                If DateDiff("d", File.DateLastModified, Now()) > mKeepDays Then

                    On Error Resume Next

                    File.Delete True

                    Err.Clear

                End If

            End If

        Next

    End Sub



    Public Function Elapsed()

        Elapsed = Round(Timer() - mStartTick, 3)

    End Function



    Public Sub ResetTimer()

        mStartTick = Timer()

    End Sub



    Private Sub WriteLog(LevelNo, LevelName, Message)

        Dim Line

        If Not mEnabled Then Exit Sub

        If LevelNo < mMinLevel Then Exit Sub

        Rotate

        Line = "[" & _
               FormatDateTimeString(Now()) & _
               "] [" & _
               LevelName & _
               "] " & _
               CStr(Message)

        WriteRaw Line

    End Sub



    Private Sub WriteRaw(Text)

        Dim TS

        EnsureFolder

        Set TS = mFSO.OpenTextFile(mLogFile, 8, True, -1)

        TS.WriteLine Text

        TS.Close

        Set TS = Nothing

    End Sub
 
     Private Function BuildLogFile()

        BuildLogFile = _
            mLogFolder & "\" & _
            mFilePrefix & "_" & _
            Year(Date()) & _
            Right("0" & Month(Date()),2) & _
            Right("0" & Day(Date()),2) & _
            ".log"

    End Function



    Private Sub EnsureFolder()

        If Not mFSO.FolderExists(mLogFolder) Then
            mFSO.CreateFolder mLogFolder
        End If

    End Sub



    Private Function FormatDateTimeString(D)

        FormatDateTimeString = _
            Year(D) & "-" & _
            Right("0" & Month(D),2) & "-" & _
            Right("0" & Day(D),2) & " " & _
            Right("0" & Hour(D),2) & ":" & _
            Right("0" & Minute(D),2) & ":" & _
            Right("0" & Second(D),2)

    End Function



    Public Function LevelName(LevelNo)

        Select Case CLng(LevelNo)

            Case LOG_DEBUG
                LevelName = "DEBUG"

            Case LOG_INFO
                LevelName = "INFO"

            Case LOG_WARN
                LevelName = "WARN"

            Case LOG_ERROR
                LevelName = "ERROR"

            Case LOG_FATAL
                LevelName = "FATAL"

            Case Else
                LevelName = "UNKNOWN"

        End Select

    End Function



    Public Function LogExists()

        LogExists = mFSO.FileExists(mLogFile)

    End Function



    Public Function LogFile()

        LogFile = mLogFile

    End Function



    Public Function LogSize()

        If mFSO.FileExists(mLogFile) Then
            LogSize = CLng(mFSO.GetFile(mLogFile).Size)
        Else
            LogSize = 0
        End If

    End Function



    Public Sub Clear()

        On Error Resume Next

        If mFSO.FileExists(mLogFile) Then
            mFSO.DeleteFile mLogFile, True
        End If

        Err.Clear

    End Sub



    Public Function ReadAll()

        Dim TS

        ReadAll = ""

        If Not mFSO.FileExists(mLogFile) Then Exit Function

        Set TS = mFSO.OpenTextFile(mLogFile, 1, False, -1)

        ReadAll = TS.ReadAll

        TS.Close

        Set TS = Nothing

    End Function

End Class
%>