<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.product.asp
' Part    : 1
'====================================================================
'Option Explicit

Class Product

    Private mData
    Private mSpecs
    Private mImages
    Private mPdfFiles

    Private Sub Class_Initialize()

        Set mData = Server.CreateObject("Scripting.Dictionary")
        Set mSpecs = Server.CreateObject("Scripting.Dictionary")

        mData.CompareMode = 1
        mSpecs.CompareMode = 1

        ReDim mImages(-1)
        ReDim mPdfFiles(-1)

        mData.Add "id", 0
        mData.Add "category_id", 0
        mData.Add "brand_id", 0

        mData.Add "title", ""
        mData.Add "slug", ""
        mData.Add "url", ""

        mData.Add "short_description", ""
        mData.Add "description", ""

        mData.Add "image", ""

        mData.Add "model", ""
        mData.Add "power", ""
        mData.Add "engine", ""
        mData.Add "alternator", ""
        mData.Add "frequency", ""
        mData.Add "rpm", ""

        mData.Add "meta_title", ""
        mData.Add "meta_keywords", ""
        mData.Add "meta_description", ""

        mData.Add "created_at", Now()
        mData.Add "updated_at", Now()

    End Sub



    Private Sub Class_Terminate()

        If IsObject(mData) Then
            mData.RemoveAll
            Set mData = Nothing
        End If

        If IsObject(mSpecs) Then
            mSpecs.RemoveAll
            Set mSpecs = Nothing
        End If

    End Sub



    Public Property Let Id(Value)
        mData("id") = CLng(Value)
    End Property

    Public Property Get Id()
        Id = mData("id")
    End Property



    Public Property Let CategoryId(Value)
        mData("category_id") = CLng(Value)
    End Property

    Public Property Get CategoryId()
        CategoryId = mData("category_id")
    End Property



    Public Property Let BrandId(Value)
        mData("brand_id") = CLng(Value)
    End Property

    Public Property Get BrandId()
        BrandId = mData("brand_id")
    End Property



    Public Property Let Title(Value)
        mData("title") = Trim(CStr(Value))
    End Property

    Public Property Get Title()
        Title = mData("title")
    End Property



    Public Property Let Slug(Value)
        mData("slug") = Trim(CStr(Value))
    End Property

    Public Property Get Slug()
        Slug = mData("slug")
    End Property



    Public Property Let Url(Value)
        mData("url") = Trim(CStr(Value))
    End Property

    Public Property Get Url()
        Url = mData("url")
    End Property



    Public Property Let ShortDescription(Value)
        mData("short_description") = CStr(Value)
    End Property

    Public Property Get ShortDescription()
        ShortDescription = mData("short_description")
    End Property



    Public Property Let Description(Value)
        mData("description") = CStr(Value)
    End Property

    Public Property Get Description()
        Description = mData("description")
    End Property



    Public Property Let Image(Value)
        mData("image") = Trim(CStr(Value))
    End Property

    Public Property Get Image()
        Image = mData("image")
    End Property



    Public Property Let Model(Value)
        mData("model") = Trim(CStr(Value))
    End Property

    Public Property Get Model()
        Model = mData("model")
    End Property



    Public Property Let Power(Value)
        mData("power") = Trim(CStr(Value))
    End Property

    Public Property Get Power()
        Power = mData("power")
    End Property



    Public Property Let Engine(Value)
        mData("engine") = Trim(CStr(Value))
    End Property

    Public Property Get Engine()
        Engine = mData("engine")
    End Property
 
     Public Property Let Alternator(Value)
        mData("alternator") = Trim(CStr(Value))
    End Property

    Public Property Get Alternator()
        Alternator = mData("alternator")
    End Property



    Public Property Let Frequency(Value)
        mData("frequency") = Trim(CStr(Value))
    End Property

    Public Property Get Frequency()
        Frequency = mData("frequency")
    End Property



    Public Property Let RPM(Value)
        mData("rpm") = Trim(CStr(Value))
    End Property

    Public Property Get RPM()
        RPM = mData("rpm")
    End Property



    Public Property Let MetaTitle(Value)
        mData("meta_title") = Trim(CStr(Value))
    End Property

    Public Property Get MetaTitle()
        MetaTitle = mData("meta_title")
    End Property



    Public Property Let MetaKeywords(Value)
        mData("meta_keywords") = Trim(CStr(Value))
    End Property

    Public Property Get MetaKeywords()
        MetaKeywords = mData("meta_keywords")
    End Property



    Public Property Let MetaDescription(Value)
        mData("meta_description") = Trim(CStr(Value))
    End Property

    Public Property Get MetaDescription()
        MetaDescription = mData("meta_description")
    End Property



    Public Sub SetValue(Key, Value)

        Key = LCase(Trim(CStr(Key)))

        If mData.Exists(Key) Then
            mData(Key) = Value
        Else
            mData.Add Key, Value
        End If

    End Sub



    Public Function GetValue(Key)

        Key = LCase(Trim(CStr(Key)))

        If mData.Exists(Key) Then
            GetValue = mData(Key)
        Else
            GetValue = Null
        End If

    End Function



    Public Sub AddSpecification(Name, Value)

        Name = Trim(CStr(Name))

        If Len(Name) = 0 Then Exit Sub

        If mSpecs.Exists(Name) Then
            mSpecs(Name) = Value
        Else
            mSpecs.Add Name, Value
        End If

    End Sub



    Public Function Specification(Name)

        If mSpecs.Exists(Name) Then
            Specification = mSpecs(Name)
        Else
            Specification = ""
        End If

    End Function



    Public Function Specifications()

        Set Specifications = mSpecs

    End Function



    Public Function SpecificationCount()

        SpecificationCount = mSpecs.Count

    End Function



    Public Sub AddImage(ImageUrl)

        Dim n

        n = UBound(mImages) + 1

        ReDim Preserve mImages(n)

        mImages(n) = Trim(CStr(ImageUrl))

    End Sub



    Public Function Images()

        Images = mImages

    End Function



    Public Function ImageCount()

        If UBound(mImages) < 0 Then
            ImageCount = 0
        Else
            ImageCount = UBound(mImages) + 1
        End If

    End Function



    Public Sub AddPdf(PdfUrl)

        Dim n

        n = UBound(mPdfFiles) + 1

        ReDim Preserve mPdfFiles(n)

        mPdfFiles(n) = Trim(CStr(PdfUrl))

    End Sub



    Public Function PdfFiles()

        PdfFiles = mPdfFiles

    End Function
 
     Public Function PdfCount()

        If UBound(mPdfFiles) < 0 Then
            PdfCount = 0
        Else
            PdfCount = UBound(mPdfFiles) + 1
        End If

    End Function



    Public Function ToDictionary()

        Set ToDictionary = mData

    End Function



    Public Sub FromDictionary(ByVal D)

        Dim K

        If D Is Nothing Then Exit Sub

        For Each K In D.Keys

            If mData.Exists(LCase(CStr(K))) Then
                mData(LCase(CStr(K))) = D(K)
            Else
                mData.Add LCase(CStr(K)), D(K)
            End If

        Next

    End Sub



    Public Function Exists(Key)

        Exists = mData.Exists(LCase(Trim(CStr(Key))))

    End Function



    Public Sub Remove(Key)

        Key = LCase(Trim(CStr(Key)))

        If mData.Exists(Key) Then
            mData.Remove Key
        End If

    End Sub



    Public Sub Clear()

        mData.RemoveAll
        mSpecs.RemoveAll

        ReDim mImages(-1)
        ReDim mPdfFiles(-1)

    End Sub



    Public Function IsValid()

        IsValid = _
            (Len(Trim(Title)) > 0) And _
            (Len(Trim(Model)) > 0)

    End Function



    Public Function Clone()

        Dim P
        Dim K
        Dim i

        Set P = New Product

        For Each K In mData.Keys
            P.SetValue K, mData(K)
        Next

        For Each K In mSpecs.Keys
            P.AddSpecification K, mSpecs(K)
        Next

        If UBound(mImages) >= 0 Then
            For i = 0 To UBound(mImages)
                P.AddImage mImages(i)
            Next
        End If

        If UBound(mPdfFiles) >= 0 Then
            For i = 0 To UBound(mPdfFiles)
                P.AddPdf mPdfFiles(i)
            Next
        End If

        Set Clone = P

    End Function

End Class
%>