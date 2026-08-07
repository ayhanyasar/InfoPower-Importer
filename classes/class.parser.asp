<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.parser.asp
' Part    : 1
'====================================================================
Option Explicit

Class Parser

    Private mDom
    Private mHtml
    Private mCards
    Private mProducts

    Private Sub Class_Initialize()

        Set mDom = New Dom
        Set mCards = Server.CreateObject("Scripting.Dictionary")

        mCards.CompareMode = 1

        ReDim mProducts(-1)

        mHtml = ""

    End Sub



    Private Sub Class_Terminate()

        Set mDom = Nothing

        If IsObject(mCards) Then
            mCards.RemoveAll
            Set mCards = Nothing
        End If

    End Sub



    Public Function Load(Html)

        mHtml = Html

        Load = mDom.LoadHtml(Html)

    End Function



    Public Function LoadFile(FileName)

        LoadFile = mDom.LoadFile(FileName)

    End Function



    Public Function Html()

        Html = mHtml

    End Function



    Public Function Title()

        Title = mDom.Title()

    End Function



    Public Function ProductCardCount()

        Dim Col

        Set Col = mDom.QuerySelectorAll("div.fixedkarscard")

        If Col Is Nothing Then

            ProductCardCount = 0

        Else

            ProductCardCount = Col.Length

        End If

    End Function



    Public Function ProductCards()

        Set ProductCards = mDom.QuerySelectorAll("div.fixedkarscard")

    End Function



    Public Function ParseCards()

        Dim Nodes
        Dim Card
        Dim Item

        Dim Index

        Index = -1

        Set Nodes = ProductCards()

        If Nodes Is Nothing Then

            ReDim mProducts(-1)

            ParseCards = 0

            Exit Function

        End If



        For Each Card In Nodes

            Index = Index + 1

            ReDim Preserve mProducts(Index)

            Set Item = ParseCard(Card)

            Set mProducts(Index) = Item

        Next

        ParseCards = Index + 1

    End Function



    Private Function ParseCard(Node)

        Dim D

        Set D = Server.CreateObject("Scripting.Dictionary")

        D.CompareMode = 1

        D.Add "id"            ,GetAttribute(Node,"data-urunid")
        D.Add "class"         ,GetAttribute(Node,"className")
        D.Add "title"         ,""
        D.Add "model"         ,""
        D.Add "url"           ,""
        D.Add "image"         ,""
        D.Add "description"   ,""
        D.Add "power"         ,""
        D.Add "engine"        ,""
        D.Add "alternator"    ,""
        D.Add "frequency"     ,""
        D.Add "rpm"           ,""

        Call ParseBasic(Node,D)

        Set ParseCard = D

    End Function



    Private Sub ParseBasic(Node,ByRef D)

        Dim A
        Dim Img
        Dim H

        Set A = First(Node,"a")

        If Not A Is Nothing Then

            D("url") = GetAttribute(A,"href")

            D("title") = Clean(A.innerText)

        End If



        Set Img = First(Node,"img")

        If Not Img Is Nothing Then

            D("image") = GetAttribute(Img,"src")

            If D("title")="" Then

                D("title") = GetAttribute(Img,"alt")

            End If

        End If



        Set H = First(Node,"h1")

        If H Is Nothing Then Set H = First(Node,"h2")
        If H Is Nothing Then Set H = First(Node,"h3")
        If H Is Nothing Then Set H = First(Node,"h4")

        If Not H Is Nothing Then

            D("title") = Clean(H.innerText)

        End If

    End Sub
 
     Private Function First(Node, TagName)

        On Error Resume Next

        Dim Col

        Set Col = Node.getElementsByTagName(TagName)

        If Err.Number <> 0 Then

            Err.Clear
            Set First = Nothing
            Exit Function

        End If

        If Col.Length > 0 Then
            Set First = Col(0)
        Else
            Set First = Nothing
        End If

    End Function



    Private Function GetAttribute(Node, AttrName)

        On Error Resume Next

        GetAttribute = ""

        If Node Is Nothing Then Exit Function

        Select Case LCase(AttrName)

            Case "classname"

                GetAttribute = "" & Node.className

            Case Else

                GetAttribute = "" & Node.getAttribute(AttrName)

        End Select

        Err.Clear

    End Function



    Private Function Clean(Value)

        Dim S

        S = CStr(Value)

        S = Replace(S, vbCr, " ")
        S = Replace(S, vbLf, " ")
        S = Replace(S, vbTab, " ")
        S = Replace(S, "&nbsp;", " ")

        Do While InStr(S, "  ") > 0
            S = Replace(S, "  ", " ")
        Loop

        Clean = Trim(S)

    End Function



    Public Function Products()

        Products = mProducts

    End Function



    Public Function Product(Index)

        If Index < 0 Then Exit Function

        If Index > UBound(mProducts) Then Exit Function

        Set Product = mProducts(Index)

    End Function



    Public Function ProductCount()

        If UBound(mProducts) < 0 Then

            ProductCount = 0

        Else

            ProductCount = UBound(mProducts) + 1

        End If

    End Function



    Public Function FindById(ProductID)

        Dim i

        Set FindById = Nothing

        For i = 0 To UBound(mProducts)

            If CStr(mProducts(i)("id")) = CStr(ProductID) Then

                Set FindById = mProducts(i)
                Exit Function

            End If

        Next

    End Function



    Public Function FindByTitle(Title)

        Dim i

        Set FindByTitle = Nothing

        For i = 0 To UBound(mProducts)

            If LCase(mProducts(i)("title")) = LCase(Title) Then

                Set FindByTitle = mProducts(i)
                Exit Function

            End If

        Next

    End Function
 
     Public Function ParseProductDetail()

        Dim D
        Dim Tables
        Dim i

        Set D = Server.CreateObject("Scripting.Dictionary")
        D.CompareMode = 1

        D.Add "title"        , mDom.Title()
        D.Add "description"  , ""
        D.Add "image"        , ""
        D.Add "gallery"      , Array()
        D.Add "pdf"          , Array()
        D.Add "specifications", Server.CreateObject("Scripting.Dictionary")

        ' Ana resim
        On Error Resume Next

        Dim Img
        Set Img = mDom.QuerySelector("img")

        If Not Img Is Nothing Then
            D("image") = Clean(GetAttribute(Img, "src"))
        End If

        ' Açıklama
        Dim P
        Set P = mDom.QuerySelector("p")

        If Not P Is Nothing Then
            D("description") = Clean(P.innerText)
        End If

        ' Teknik özellik tabloları
        Set Tables = mDom.ElementsByTag("table")

        If Not Tables Is Nothing Then

            For i = 0 To Tables.Length - 1
                ParseSpecificationTable Tables(i), D("specifications")
            Next

        End If

        Set ParseProductDetail = D

    End Function



    Private Sub ParseSpecificationTable(TableNode, ByRef Specs)

        Dim Rows
        Dim r
        Dim Cells
        Dim Key
        Dim Value

        On Error Resume Next

        Set Rows = TableNode.getElementsByTagName("tr")

        If Rows Is Nothing Then Exit Sub

        For Each r In Rows

            Set Cells = r.getElementsByTagName("td")

            If Cells.Length >= 2 Then

                Key = Clean(Cells(0).innerText)
                Value = Clean(Cells(1).innerText)

                If Len(Key) > 0 Then

                    If Specs.Exists(Key) Then
                        Specs(Key) = Value
                    Else
                        Specs.Add Key, Value
                    End If

                End If

            End If

        Next

    End Sub



    Public Sub Clear()

        ReDim mProducts(-1)

        mCards.RemoveAll

        mHtml = ""

        Set mDom = New Dom

    End Sub

End Class
%>