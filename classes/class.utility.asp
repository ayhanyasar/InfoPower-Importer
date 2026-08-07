<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.utility.asp
' Part    : 1
'====================================================================
'Option Explicit

Class Utility

    Private mFSO

    Private Sub Class_Initialize()

        Set mFSO = Server.CreateObject("Scripting.FileSystemObject")

        Randomize Timer

    End Sub

    Private Sub Class_Terminate()

        Set mFSO = Nothing

    End Sub

    Public Function IsNullOrEmpty(v)

        If IsNull(v) Then
            IsNullOrEmpty = True
        ElseIf IsEmpty(v) Then
            IsNullOrEmpty = True
        ElseIf Trim(CStr(v)) = "" Then
            IsNullOrEmpty = True
        Else
            IsNullOrEmpty = False
        End If

    End Function

    Public Function Nz(v, DefaultValue)

        If IsNullOrEmpty(v) Then
            Nz = DefaultValue
        Else
            Nz = v
        End If

    End Function

    Public Function ToInt(v)

        If IsNumeric(v) Then
            ToInt = CLng(v)
        Else
            ToInt = 0
        End If

    End Function

    Public Function ToDouble(v)

        If IsNumeric(v) Then
            ToDouble = CDbl(v)
        Else
            ToDouble = 0
        End If

    End Function

    Public Function ToBool(v)

        Select Case LCase(Trim(CStr(v)))
            Case "1","true","yes","on","evet"
                ToBool = True
            Case Else
                ToBool = False
        End Select

    End Function

    Public Function HtmlEncode(Text)

        Text = CStr(Text)

        Text = Replace(Text,"&","&amp;")
        Text = Replace(Text,"<","&lt;")
        Text = Replace(Text,">","&gt;")
        Text = Replace(Text,Chr(34),"&quot;")
        Text = Replace(Text,"'","&#39;")

        HtmlEncode = Text

    End Function

    Public Function HtmlDecode(Text)

        Text = CStr(Text)

        Text = Replace(Text,"&lt;","<")
        Text = Replace(Text,"&gt;",">")
        Text = Replace(Text,"&quot;",Chr(34))
        Text = Replace(Text,"&#39;","'")
        Text = Replace(Text,"&amp;","&")

        HtmlDecode = Text

    End Function

    Public Function UrlEncode(Text)

        Dim i
        Dim c
        Dim n
        Dim s

        s = ""

        For i = 1 To Len(Text)

            c = Mid(Text,i,1)
            n = AscW(c)

            If (n>=48 And n<=57) _
            Or (n>=65 And n<=90) _
            Or (n>=97 And n<=122) Then

                s = s & c

            ElseIf c=" " Then

                s = s & "+"

            Else

                s = s & "%" & Right("0" & Hex(n),2)

            End If

        Next

        UrlEncode = s

    End Function

    Public Function TrimAll(Text)

        Dim s

        s = Trim(CStr(Text))

        Do While InStr(s,"  ")>0
            s = Replace(s,"  "," ")
        Loop

        s = Replace(s,vbCr," ")
        s = Replace(s,vbLf," ")
        s = Replace(s,vbTab," ")

        Do While InStr(s,"  ")>0
            s = Replace(s,"  "," ")
        Loop

        TrimAll = Trim(s)

    End Function

    Public Function StripTags(Html)

        Dim RE

        Set RE = New RegExp

        RE.Global = True
        RE.IgnoreCase = True
        RE.Pattern = "<[^>]*>"

        StripTags = RE.Replace(Html,"")

        Set RE = Nothing

    End Function

    Public Function FileExists(Path)

        FileExists = mFSO.FileExists(Path)

    End Function

    Public Function FolderExists(Path)

        FolderExists = mFSO.FolderExists(Path)

    End Function

    Public Sub CreateFolder(Path)

        If Not mFSO.FolderExists(Path) Then
            mFSO.CreateFolder Path
        End If

    End Sub

    Public Function ReadFile(Path)

        Dim ts

        ReadFile = ""

        If Not FileExists(Path) Then Exit Function

        Set ts = mFSO.OpenTextFile(Path,1,False,-1)

        ReadFile = ts.ReadAll

        ts.Close

        Set ts = Nothing

    End Function

    Public Function WriteFile(Path,Content)

        Dim ts

        On Error Resume Next

        Set ts = mFSO.CreateTextFile(Path,True,True)

        ts.Write Content

        ts.Close

        Set ts = Nothing

        WriteFile = (Err.Number=0)

        Err.Clear

    End Function
				
    Public Function AppendFile(Path, Content)

        Dim ts

        On Error Resume Next

        If mFSO.FileExists(Path) Then
            Set ts = mFSO.OpenTextFile(Path,8,True,-1)
        Else
            Set ts = mFSO.CreateTextFile(Path,True,True)
        End If

        ts.Write Content
        ts.Close

        Set ts = Nothing

        AppendFile = (Err.Number = 0)

        Err.Clear

    End Function



    Public Function DeleteFile(Path)

        On Error Resume Next

        If mFSO.FileExists(Path) Then
            mFSO.DeleteFile Path, True
        End If

        DeleteFile = (Err.Number = 0)

        Err.Clear

    End Function



    Public Function CopyFile(SourcePath, TargetPath)

        On Error Resume Next

        mFSO.CopyFile SourcePath, TargetPath, True

        CopyFile = (Err.Number = 0)

        Err.Clear

    End Function



    Public Function MoveFile(SourcePath, TargetPath)

        On Error Resume Next

        mFSO.MoveFile SourcePath, TargetPath

        MoveFile = (Err.Number = 0)

        Err.Clear

    End Function



    Public Function DeleteFolder(Path)

        On Error Resume Next

        If mFSO.FolderExists(Path) Then
            mFSO.DeleteFolder Path, True
        End If

        DeleteFolder = (Err.Number = 0)

        Err.Clear

    End Function



    Public Function RandomString(LenValue)

        Dim s
        Dim c
        Dim i

        c = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        s = ""

        For i = 1 To CLng(LenValue)

            s = s & Mid(c, Int(Rnd() * Len(c)) + 1, 1)

        Next

        RandomString = s

    End Function



    Public Function Guid()

        Dim t

        Set t = Server.CreateObject("Scriptlet.TypeLib")

        Guid = Mid(t.Guid,2,36)

        Set t = Nothing

    End Function



    Public Function MD5(Text)

        Dim xml
        Dim node

        Set xml = Server.CreateObject("MSXML2.DOMDocument.6.0")
        Set node = xml.createElement("tmp")

        node.DataType = "bin.hex"

        node.nodeTypedValue = Text

        MD5 = LCase(node.Text)

        Set node = Nothing
        Set xml = Nothing

    End Function



    Public Function UnixTime()

        UnixTime = DateDiff("s","01.01.1970 00:00:00",Now())

    End Function



    Public Function FormatBytes(ByteCount)

        Dim n

        n = CDbl(ByteCount)

        If n < 1024 Then

            FormatBytes = CLng(n) & " B"

        ElseIf n < 1048576 Then

            FormatBytes = FormatNumber(n / 1024,2) & " KB"

        ElseIf n < 1073741824 Then

            FormatBytes = FormatNumber(n / 1048576,2) & " MB"

        Else

            FormatBytes = FormatNumber(n / 1073741824,2) & " GB"

        End If

    End Function



    Public Function FileExtension(FileName)

        Dim p

        p = InStrRev(FileName,".")

        If p > 0 Then
            FileExtension = LCase(Mid(FileName,p+1))
        Else
            FileExtension = ""
        End If

    End Function



    Public Function FileNameOnly(FileName)

        Dim p

        p = InStrRev(FileName,"\")

        If p > 0 Then
            FileNameOnly = Mid(FileName,p+1)
        Else
            FileNameOnly = FileName
        End If

    End Function



    Public Function RemoveExtension(FileName)

        Dim p

        p = InStrRev(FileName,".")

        If p > 0 Then
            RemoveExtension = Left(FileName,p-1)
        Else
            RemoveExtension = FileName
        End If

    End Function

    Public Function GetFileSize(Path)

        If mFSO.FileExists(Path) Then
            GetFileSize = CLng(mFSO.GetFile(Path).Size)
        Else
            GetFileSize = 0
        End If

    End Function



    Public Function GetFileDate(Path)

        If mFSO.FileExists(Path) Then
            GetFileDate = mFSO.GetFile(Path).DateLastModified
        Else
            GetFileDate = Null
        End If

    End Function



    Public Function EnsureTrailingSlash(Path)

        If Right(Path,1)="\" Then
            EnsureTrailingSlash = Path
        Else
            EnsureTrailingSlash = Path & "\"
        End If

    End Function



    Public Function ReplaceCRLF(Text)

        Text = Replace(Text,vbCrLf,vbLf)
        Text = Replace(Text,vbCr,vbLf)

        ReplaceCRLF = Text

    End Function



    Public Function Repeat(Text, Count)

        Dim i
        Dim s

        s = ""

        For i = 1 To CLng(Count)
            s = s & Text
        Next

        Repeat = s

    End Function



    Public Function PadLeft(Text, TotalLength, PadChar)

        Dim s

        s = CStr(Text)

        Do While Len(s) < CLng(TotalLength)
            s = PadChar & s
        Loop

        PadLeft = s

    End Function



    Public Function PadRight(Text, TotalLength, PadChar)

        Dim s

        s = CStr(Text)

        Do While Len(s) < CLng(TotalLength)
            s = s & PadChar
        Loop

        PadRight = s

    End Function



    Public Function StartsWith(Text, Find)

        StartsWith = (Left(CStr(Text),Len(CStr(Find))) = CStr(Find))

    End Function



    Public Function EndsWith(Text, Find)

        EndsWith = (Right(CStr(Text),Len(CStr(Find))) = CStr(Find))

    End Function



    Public Function Contains(Text, Find)

        Contains = (InStr(1,CStr(Text),CStr(Find),vbTextCompare) > 0)

    End Function



    Public Function Reverse(Text)

        Dim i
        Dim s

        s = ""

        For i = Len(Text) To 1 Step -1
            s = s & Mid(Text,i,1)
        Next

        Reverse = s

    End Function



    Public Function SafeFileName(Text)

        Dim s

        s = CStr(Text)

        s = Replace(s,"\","_")
        s = Replace(s,"/","_")
        s = Replace(s,":","_")
        s = Replace(s,"*","_")
        s = Replace(s,"?","_")
        s = Replace(s,"""","_")
        s = Replace(s,"<","_")
        s = Replace(s,">","_")
        s = Replace(s,"|","_")

        SafeFileName = Trim(s)

    End Function



    Public Function CreateFolderTree(Path)

        Dim arr
        Dim p
        Dim i

        arr = Split(Path,"\")

        p = arr(0)

        For i = 1 To UBound(arr)

            p = p & "\" & arr(i)

            If Not mFSO.FolderExists(p) Then
                On Error Resume Next
                mFSO.CreateFolder p
                Err.Clear
            End If

        Next

        CreateFolderTree = mFSO.FolderExists(Path)

    End Function

End Class
%>