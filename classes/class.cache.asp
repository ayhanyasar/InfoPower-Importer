<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.cache.asp
'====================================================================
Option Explicit

Class CacheItem

    Public Key
    Public Value
    Public Created
    Public Expires
    Public Hit

    Private Sub Class_Initialize()

        Key=""
        Value=""
        Created=Now()
        Expires=Now()
        Hit=0

    End Sub

End Class



Class Cache

    Private mFSO
    Private mRoot
    Private mExt

    Private Sub Class_Initialize()

        Set mFSO=Server.CreateObject("Scripting.FileSystemObject")

        mRoot=Server.MapPath("/cache")

        mExt=".cache"

        If Not mFSO.FolderExists(mRoot) Then
            mFSO.CreateFolder mRoot
        End If

    End Sub

    Private Sub Class_Terminate()

        Set mFSO=Nothing

    End Sub



    Public Property Let Root(Path)

        mRoot=Path

        If Not mFSO.FolderExists(mRoot) Then
            mFSO.CreateFolder mRoot
        End If

    End Property



    Public Property Get Root()

        Root=mRoot

    End Property



    Public Property Let Extension(Value)

        If Left(Value,1)="." Then
            mExt=Value
        Else
            mExt="." & Value
        End If

    End Property



    Public Property Get Extension()

        Extension=mExt

    End Property



    Public Function Exists(Key)

        Exists=mFSO.FileExists(BuildFile(Key))

    End Function



    Public Function Remove(Key)

        On Error Resume Next

        If Exists(Key) Then

            mFSO.DeleteFile BuildFile(Key),True

        End If

        Remove=(Err.Number=0)

        Err.Clear

    End Function



    Public Sub RemoveAll()

        Dim f

        If Not mFSO.FolderExists(mRoot) Then Exit Sub

        For Each f In mFSO.GetFolder(mRoot).Files

            If LCase(Right(f.Name,Len(mExt)))=LCase(mExt) Then

                On Error Resume Next
                f.Delete True
                Err.Clear

            End If

        Next

    End Sub



    Public Function Save(Key,Value)

        Dim ts

        On Error Resume Next

        Set ts=mFSO.CreateTextFile(BuildFile(Key),True,True)

        ts.Write Value

        ts.Close

        Set ts=Nothing

        Save=(Err.Number=0)

        Err.Clear

    End Function



    Public Function Append(Key,Value)

        Dim ts

        On Error Resume Next

        If Exists(Key) Then

            Set ts=mFSO.OpenTextFile(BuildFile(Key),8,True,-1)

        Else

            Set ts=mFSO.CreateTextFile(BuildFile(Key),True,True)

        End If

        ts.Write Value

        ts.Close

        Set ts=Nothing

        Append=(Err.Number=0)

        Err.Clear

    End Function



    Public Function Load(Key)

        Dim ts

        Load=""

        If Not Exists(Key) Then Exit Function

        On Error Resume Next

        Set ts=mFSO.OpenTextFile(BuildFile(Key),1,False,-1)

        Load=ts.ReadAll

        ts.Close

        Set ts=Nothing

        Err.Clear

    End Function



    Public Function LoadLines(Key)

        Dim ts
        Dim txt

        txt=""

        If Not Exists(Key) Then

            LoadLines=Array()

            Exit Function

        End If

        Set ts=mFSO.OpenTextFile(BuildFile(Key),1,False,-1)

        txt=ts.ReadAll

        ts.Close

        Set ts=Nothing

        LoadLines=Split(txt,vbCrLf)

    End Function



    Public Function Age(Key)

        If Not Exists(Key) Then

            Age=-1
            Exit Function

        End If

        Age=DateDiff("s",mFSO.GetFile(BuildFile(Key)).DateLastModified,Now())

    End Function



    Public Function IsExpired(Key,Minute)

        If Not Exists(Key) Then

            IsExpired=True
            Exit Function

        End If

        IsExpired=(DateDiff("n",mFSO.GetFile(BuildFile(Key)).DateLastModified,Now())>=CLng(Minute))

    End Function
	
    Public Function LastModified(Key)

        If Exists(Key) Then
            LastModified = mFSO.GetFile(BuildFile(Key)).DateLastModified
        Else
            LastModified = Null
        End If

    End Function



    Public Function Created(Key)

        If Exists(Key) Then
            Created = mFSO.GetFile(BuildFile(Key)).DateCreated
        Else
            Created = Null
        End If

    End Function



    Public Function Size(Key)

        If Exists(Key) Then
            Size = CLng(mFSO.GetFile(BuildFile(Key)).Size)
        Else
            Size = 0
        End If

    End Function



    Public Function ReadBytes(Key)

        Dim stm

        ReadBytes = Null

        If Not Exists(Key) Then Exit Function

        Set stm = Server.CreateObject("ADODB.Stream")

        stm.Type = 1
        stm.Open
        stm.LoadFromFile BuildFile(Key)

        ReadBytes = stm.Read

        stm.Close
        Set stm = Nothing

    End Function



    Public Function SaveBytes(Key, Bytes)

        Dim stm

        On Error Resume Next

        Set stm = Server.CreateObject("ADODB.Stream")

        stm.Type = 1
        stm.Open
        stm.Write Bytes
        stm.SaveToFile BuildFile(Key),2
        stm.Close

        Set stm = Nothing

        SaveBytes = (Err.Number = 0)

        Err.Clear

    End Function



    Public Function Touch(Key)

        Dim txt

        If Not Exists(Key) Then
            Touch = False
            Exit Function
        End If

        txt = Load(Key)

        Touch = Save(Key, txt)

    End Function



    Public Function Copy(SourceKey, TargetKey)

        On Error Resume Next

        If Not Exists(SourceKey) Then
            Copy = False
            Exit Function
        End If

        mFSO.CopyFile _
            BuildFile(SourceKey), _
            BuildFile(TargetKey), _
            True

        Copy = (Err.Number = 0)

        Err.Clear

    End Function



    Public Function Move(SourceKey, TargetKey)

        On Error Resume Next

        If Not Exists(SourceKey) Then
            Move = False
            Exit Function
        End If

        mFSO.MoveFile _
            BuildFile(SourceKey), _
            BuildFile(TargetKey)

        Move = (Err.Number = 0)

        Err.Clear

    End Function



    Public Function Rename(SourceKey, NewKey)

        Rename = Move(SourceKey, NewKey)

    End Function



    Public Function FileName(Key)

        FileName = SafeKey(Key) & mExt

    End Function



    Public Function FilePath(Key)

        FilePath = BuildFile(Key)

    End Function



    Public Function Keys()

        Dim Folder
        Dim File
        Dim Arr()
        Dim i
        Dim NameOnly

        Set Folder = mFSO.GetFolder(mRoot)

        ReDim Arr(-1)

        i = -1

        For Each File In Folder.Files

            If LCase(Right(File.Name, Len(mExt))) = LCase(mExt) Then

                i = i + 1
                ReDim Preserve Arr(i)

                NameOnly = Left(File.Name, Len(File.Name) - Len(mExt))

                Arr(i) = NameOnly

            End If

        Next

        Keys = Arr

    End Function



    Public Function Count()

        Dim Folder
        Dim File
        Dim n

        n = 0

        Set Folder = mFSO.GetFolder(mRoot)

        For Each File In Folder.Files

            If LCase(Right(File.Name, Len(mExt))) = LCase(mExt) Then
                n = n + 1
            End If

        Next

        Count = n

    End Function
    Public Function TotalSize()

        Dim Folder
        Dim File
        Dim n

        n = 0

        If Not mFSO.FolderExists(mRoot) Then
            TotalSize = 0
            Exit Function
        End If

        Set Folder = mFSO.GetFolder(mRoot)

        For Each File In Folder.Files

            If LCase(Right(File.Name, Len(mExt))) = LCase(mExt) Then
                n = n + CLng(File.Size)
            End If

        Next

        TotalSize = n

    End Function



    Public Function GetInfo(Key)

        Dim d

        Set d = Server.CreateObject("Scripting.Dictionary")

        d.CompareMode = 1

        d.Add "exists", Exists(Key)

        If Exists(Key) Then

            d.Add "path", BuildFile(Key)
            d.Add "name", FileName(Key)
            d.Add "size", Size(Key)
            d.Add "created", Created(Key)
            d.Add "modified", LastModified(Key)
            d.Add "age", Age(Key)

        End If

        Set GetInfo = d

    End Function



    Public Function ReadIfValid(Key, ExpireMinute)

        If IsExpired(Key, ExpireMinute) Then
            ReadIfValid = ""
        Else
            ReadIfValid = Load(Key)
        End If

    End Function



    Public Function SaveIfDifferent(Key, Value)

        If Exists(Key) Then

            If Load(Key) = CStr(Value) Then
                SaveIfDifferent = True
                Exit Function
            End If

        End If

        SaveIfDifferent = Save(Key, Value)

    End Function



    Public Function PurgeExpired(ExpireMinute)

        Dim Folder
        Dim File
        Dim Deleted

        Deleted = 0

        If Not mFSO.FolderExists(mRoot) Then
            PurgeExpired = 0
            Exit Function
        End If

        Set Folder = mFSO.GetFolder(mRoot)

        For Each File In Folder.Files

            If LCase(Right(File.Name, Len(mExt))) = LCase(mExt) Then

                If DateDiff("n", File.DateLastModified, Now()) >= CLng(ExpireMinute) Then

                    On Error Resume Next

                    File.Delete True

                    If Err.Number = 0 Then
                        Deleted = Deleted + 1
                    End If

                    Err.Clear

                End If

            End If

        Next

        PurgeExpired = Deleted

    End Function



    Private Function BuildFile(Key)

        BuildFile = mRoot & "\" & SafeKey(Key) & mExt

    End Function



    Private Function SafeKey(Key)

        Dim s

        s = Trim(CStr(Key))

        s = Replace(s, "\", "_")
        s = Replace(s, "/", "_")
        s = Replace(s, ":", "_")
        s = Replace(s, "*", "_")
        s = Replace(s, "?", "_")
        s = Replace(s, """", "_")
        s = Replace(s, "<", "_")
        s = Replace(s, ">", "_")
        s = Replace(s, "|", "_")
        s = Replace(s, "&", "_")
        s = Replace(s, "%", "_")
        s = Replace(s, "#", "_")
        s = Replace(s, "=", "_")
        s = Replace(s, " ", "_")

        SafeKey = s

    End Function

End Class
%>