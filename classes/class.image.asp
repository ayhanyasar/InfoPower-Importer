<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.image.asp
' Part    : 1
'====================================================================
Option Explicit

Class ImageManager

    Private mHttp
    Private mFSO

    Private mSavePath
    Private mBaseUrl

    Private mOverwrite
    Private mCreateFolders

    Private mAllowed

    Private Sub Class_Initialize()

        Set mHttp = New HttpClient
        Set mFSO  = Server.CreateObject("Scripting.FileSystemObject")

        mSavePath = Server.MapPath("/images/products")
        mBaseUrl  = "/images/products"

        mOverwrite = True
        mCreateFolders = True

        mAllowed = Array( _
            "jpg", _
            "jpeg", _
            "png", _
            "gif", _
            "webp", _
            "bmp" _
        )

        EnsureFolder mSavePath

    End Sub



    Private Sub Class_Terminate()

        Set mHttp = Nothing
        Set mFSO = Nothing

    End Sub



    Public Property Let SavePath(Value)

        mSavePath = Trim(CStr(Value))

        EnsureFolder mSavePath

    End Property



    Public Property Get SavePath()

        SavePath = mSavePath

    End Property



    Public Property Let BaseUrl(Value)

        mBaseUrl = Trim(CStr(Value))

    End Property



    Public Property Get BaseUrl()

        BaseUrl = mBaseUrl

    End Property



    Public Property Let Overwrite(Value)

        mOverwrite = CBool(Value)

    End Property



    Public Property Get Overwrite()

        Overwrite = mOverwrite

    End Property



    Public Function Download(ImageUrl)

        Dim Resp

        Set Resp = mHttp.Get(ImageUrl)

        If Resp Is Nothing Then

            Set Download = Nothing
            Exit Function

        End If

        If Not Resp.Success Then

            Set Download = Nothing
            Exit Function

        End If

        Set Download = Resp

    End Function



    Public Function Save(ImageUrl)

        Dim Resp
        Dim FileName
        Dim FullName
        Dim Stm

        Save = ""

        Set Resp = Download(ImageUrl)

        If Resp Is Nothing Then Exit Function

        FileName = BuildFileName(ImageUrl)

        FullName = mSavePath & "\" & FileName

        If mFSO.FileExists(FullName) Then

            If Not mOverwrite Then

                Save = FullName
                Exit Function

            End If

        End If

        Set Stm = Server.CreateObject("ADODB.Stream")

        Stm.Type = 1
        Stm.Open

        Stm.Write Resp.Bytes

        Stm.SaveToFile FullName,2

        Stm.Close

        Set Stm = Nothing

        Save = FullName

    End Function



    Public Function SaveAs(ImageUrl,FileName)

        Dim Resp
        Dim FullName
        Dim Stm

        SaveAs = ""

        Set Resp = Download(ImageUrl)

        If Resp Is Nothing Then Exit Function

        FullName = mSavePath & "\" & SafeFileName(FileName)

        Set Stm = Server.CreateObject("ADODB.Stream")

        Stm.Type = 1
        Stm.Open

        Stm.Write Resp.Bytes

        Stm.SaveToFile FullName,2

        Stm.Close

        Set Stm = Nothing

        SaveAs = FullName

    End Function
 
     Public Function Exists(FileName)

        Exists = mFSO.FileExists(mSavePath & "\" & FileName)

    End Function



    Public Function Delete(FileName)

        Delete = False

        On Error Resume Next

        If Exists(FileName) Then

            mFSO.DeleteFile mSavePath & "\" & FileName, True

            Delete = (Err.Number = 0)

        End If

        Err.Clear

    End Function



    Public Function Url(FileName)

        Url = mBaseUrl & "/" & FileName

    End Function



    Public Function BuildFileName(ImageUrl)

        Dim FileName
        Dim Ext
        Dim p

        FileName = ImageUrl

        p = InStrRev(FileName, "/")

        If p > 0 Then
            FileName = Mid(FileName, p + 1)
        End If

        p = InStr(FileName, "?")

        If p > 0 Then
            FileName = Left(FileName, p - 1)
        End If

        Ext = LCase(GetExtension(FileName))

        If Not IsAllowedExtension(Ext) Then
            FileName = GetBaseName(FileName) & ".jpg"
        End If

        BuildFileName = SafeFileName(FileName)

    End Function



    Private Function GetExtension(FileName)

        Dim p

        p = InStrRev(FileName, ".")

        If p > 0 Then
            GetExtension = Mid(FileName, p + 1)
        Else
            GetExtension = ""
        End If

    End Function



    Private Function GetBaseName(FileName)

        Dim p

        p = InStrRev(FileName, ".")

        If p > 0 Then
            GetBaseName = Left(FileName, p - 1)
        Else
            GetBaseName = FileName
        End If

    End Function



    Private Function IsAllowedExtension(Ext)

        Dim i

        IsAllowedExtension = False

        For i = LBound(mAllowed) To UBound(mAllowed)

            If LCase(Ext) = mAllowed(i) Then
                IsAllowedExtension = True
                Exit Function
            End If

        Next

    End Function



    Private Function SafeFileName(FileName)

        Dim S

        S = Trim(CStr(FileName))

        S = Replace(S, "\", "_")
        S = Replace(S, "/", "_")
        S = Replace(S, ":", "_")
        S = Replace(S, "*", "_")
        S = Replace(S, "?", "_")
        S = Replace(S, """", "_")
        S = Replace(S, "<", "_")
        S = Replace(S, ">", "_")
        S = Replace(S, "|", "_")

        SafeFileName = S

    End Function
 
     Private Sub EnsureFolder(Path)

        If mFSO.FolderExists(Path) Then Exit Sub

        CreateFolderTree Path

    End Sub



    Private Sub CreateFolderTree(Path)

        Dim Parts
        Dim Current
        Dim i

        Parts = Split(Path, "\")

        If UBound(Parts) < 0 Then Exit Sub

        Current = Parts(0)

        If Right(Current,1) = ":" Then
            Current = Current & "\"
        End If

        For i = 1 To UBound(Parts)

            If Right(Current,1) <> "\" Then
                Current = Current & "\"
            End If

            Current = Current & Parts(i)

            If Not mFSO.FolderExists(Current) Then

                On Error Resume Next
                mFSO.CreateFolder Current
                Err.Clear

            End If

        Next

    End Sub



    Public Function ImageInfo(FileName)

        Dim D
        Dim F

        Set D = Server.CreateObject("Scripting.Dictionary")
        D.CompareMode = 1

        D.Add "exists", False
        D.Add "filename", FileName
        D.Add "path", mSavePath & "\" & FileName
        D.Add "url", Url(FileName)
        D.Add "size", 0
        D.Add "date", Null

        If Exists(FileName) Then

            Set F = mFSO.GetFile(mSavePath & "\" & FileName)

            D("exists") = True
            D("size") = CLng(F.Size)
            D("date") = F.DateLastModified

        End If

        Set ImageInfo = D

    End Function



    Public Function FileCount()

        Dim Folder

        Set Folder = mFSO.GetFolder(mSavePath)

        FileCount = Folder.Files.Count

    End Function



    Public Sub Clear()

        Dim Folder
        Dim F

        Set Folder = mFSO.GetFolder(mSavePath)

        For Each F In Folder.Files

            On Error Resume Next

            F.Delete True

            Err.Clear

        Next

    End Sub

End Class
%>