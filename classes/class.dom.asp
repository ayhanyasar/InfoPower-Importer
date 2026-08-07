<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.dom.asp
' Part    : 1
'====================================================================
'Option Explicit

Class Dom

    Private mHtml
    Private mDocument
    Private mReady

    Private Sub Class_Initialize()

        mHtml = ""
        mReady = False

        Set mDocument = Server.CreateObject("HTMLFILE")

    End Sub

    Private Sub Class_Terminate()

        Set mDocument = Nothing

    End Sub



    Public Function LoadHtml(Html)

        On Error Resume Next

        mHtml = CStr(Html)

        mDocument.Open
        mDocument.Write mHtml
        mDocument.Close

        mReady = (Err.Number = 0)

        Err.Clear

        LoadHtml = mReady

    End Function



    Public Function LoadFile(FileName)

        Dim FSO
        Dim TS
        Dim Html

        LoadFile = False

        Set FSO = Server.CreateObject("Scripting.FileSystemObject")

        If Not FSO.FileExists(FileName) Then
            Set FSO = Nothing
            Exit Function
        End If

        Set TS = FSO.OpenTextFile(FileName,1,False,-1)

        Html = TS.ReadAll

        TS.Close

        Set TS = Nothing
        Set FSO = Nothing

        LoadFile = LoadHtml(Html)

    End Function



    Public Property Get Document()

        Set Document = mDocument

    End Property



    Public Function Html()

        Html = mHtml

    End Function



    Public Function Title()

        On Error Resume Next

        Title = Trim(mDocument.Title)

        Err.Clear

    End Function



    Public Function BodyText()

        On Error Resume Next

        BodyText = mDocument.Body.InnerText

        Err.Clear

    End Function



    Public Function BodyHtml()

        On Error Resume Next

        BodyHtml = mDocument.Body.InnerHTML

        Err.Clear

    End Function



    Public Function ElementById(Id)

        On Error Resume Next

        Set ElementById = mDocument.getElementById(Id)

        Err.Clear

    End Function



    Public Function ElementsByTag(TagName)

        On Error Resume Next

        Set ElementsByTag = mDocument.getElementsByTagName(TagName)

        Err.Clear

    End Function



    Public Function ElementsByClass(ClassName)

        On Error Resume Next

        Set ElementsByClass = mDocument.getElementsByClassName(ClassName)

        Err.Clear

    End Function



    Public Function Links()

        On Error Resume Next

        Set Links = mDocument.Links

        Err.Clear

    End Function



    Public Function Images()

        On Error Resume Next

        Set Images = mDocument.Images

        Err.Clear

    End Function
 
     Public Function Forms()

        On Error Resume Next

        Set Forms = mDocument.Forms

        Err.Clear

    End Function



    Public Function Scripts()

        On Error Resume Next

        Set Scripts = mDocument.getElementsByTagName("script")

        Err.Clear

    End Function



    Public Function MetaTags()

        On Error Resume Next

        Set MetaTags = mDocument.getElementsByTagName("meta")

        Err.Clear

    End Function



    Public Function ExistsById(Id)

        ExistsById = Not (ElementById(Id) Is Nothing)

    End Function



    Public Function Count(TagName)

        On Error Resume Next

        Count = mDocument.getElementsByTagName(TagName).Length

        Err.Clear

    End Function



    Public Function GetAttributeById(Id, Attr)

        Dim E

        GetAttributeById = ""

        Set E = ElementById(Id)

        If E Is Nothing Then Exit Function

        On Error Resume Next

        GetAttributeById = E.getAttribute(Attr)

        Err.Clear

    End Function



    Public Function InnerTextById(Id)

        Dim E

        InnerTextById = ""

        Set E = ElementById(Id)

        If E Is Nothing Then Exit Function

        On Error Resume Next

        InnerTextById = E.innerText

        Err.Clear

    End Function



    Public Function InnerHtmlById(Id)

        Dim E

        InnerHtmlById = ""

        Set E = ElementById(Id)

        If E Is Nothing Then Exit Function

        On Error Resume Next

        InnerHtmlById = E.innerHTML

        Err.Clear

    End Function



    Public Function OuterHtmlById(Id)

        Dim E

        OuterHtmlById = ""

        Set E = ElementById(Id)

        If E Is Nothing Then Exit Function

        On Error Resume Next

        OuterHtmlById = E.outerHTML

        Err.Clear

    End Function



    Public Function GetMeta(Name)

        Dim M
        Dim i

        GetMeta = ""

        Set M = MetaTags()

        If M Is Nothing Then Exit Function

        For i = 0 To M.Length - 1

            If LCase("" & M(i).getAttribute("name")) = LCase(Name) Then

                GetMeta = M(i).getAttribute("content")
                Exit Function

            End If

        Next

    End Function



    Public Function GetElementsText(TagName)

        Dim Col
        Dim Arr()
        Dim i

        Set Col = ElementsByTag(TagName)

        If Col Is Nothing Then

            ReDim Arr(-1)
            GetElementsText = Arr
            Exit Function

        End If

        ReDim Arr(Col.Length - 1)

        For i = 0 To Col.Length - 1
            Arr(i) = Col(i).innerText
        Next

        GetElementsText = Arr

    End Function
 
     Public Function GetElementsHtml(TagName)

        Dim Col
        Dim Arr()
        Dim i

        Set Col = ElementsByTag(TagName)

        If Col Is Nothing Then

            ReDim Arr(-1)
            GetElementsHtml = Arr
            Exit Function

        End If

        ReDim Arr(Col.Length - 1)

        For i = 0 To Col.Length - 1
            Arr(i) = Col(i).outerHTML
        Next

        GetElementsHtml = Arr

    End Function



    Public Function GetLinks()

        Dim Col
        Dim Arr()
        Dim i

        Set Col = Links()

        If Col Is Nothing Then
            ReDim Arr(-1)
            GetLinks = Arr
            Exit Function
        End If

        ReDim Arr(Col.Length - 1)

        For i = 0 To Col.Length - 1
            Arr(i) = Col(i).href
        Next

        GetLinks = Arr

    End Function



    Public Function GetImages()

        Dim Col
        Dim Arr()
        Dim i

        Set Col = Images()

        If Col Is Nothing Then
            ReDim Arr(-1)
            GetImages = Arr
            Exit Function
        End If

        ReDim Arr(Col.Length - 1)

        For i = 0 To Col.Length - 1
            Arr(i) = Col(i).src
        Next

        GetImages = Arr

    End Function



    Public Function QuerySelector(CssSelector)

        On Error Resume Next

        Set QuerySelector = mDocument.querySelector(CssSelector)

        Err.Clear

    End Function



    Public Function QuerySelectorAll(CssSelector)

        On Error Resume Next

        Set QuerySelectorAll = mDocument.querySelectorAll(CssSelector)

        Err.Clear

    End Function



    Public Function Ready()

        Ready = mReady

    End Function



    Public Sub Clear()

        mHtml = ""
        mReady = False

        Set mDocument = Nothing
        Set mDocument = Server.CreateObject("HTMLFILE")

    End Sub

End Class
%>