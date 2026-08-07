<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.config.asp
'====================================================================
'Option Explicit

Class Config

    Private mItems

    Private Sub Class_Initialize()

        Set mItems = Server.CreateObject("Scripting.Dictionary")
        mItems.CompareMode = 1

        LoadDefaults

    End Sub

    Private Sub Class_Terminate()

        If IsObject(mItems) Then
            mItems.RemoveAll
            Set mItems = Nothing
        End If

    End Sub

    Private Sub LoadDefaults()

        '----------------------------------------------------------
        ' Application
        '----------------------------------------------------------
        SetValue "APP_NAME"             ,"InfoPower Importer PRO"
        SetValue "APP_VERSION"          ,"3.0"
        SetValue "APP_CHARSET"          ,"utf-8"
        SetValue "APP_TIMEZONE"         ,"Turkey Standard Time"
        SetValue "APP_LANGUAGE"         ,"tr"
        SetValue "APP_DEBUG"            ,False

        '----------------------------------------------------------
        ' Paths
        '----------------------------------------------------------
        SetValue "ROOT_PATH"            ,Server.MapPath("/")
        SetValue "CACHE_PATH"           ,Server.MapPath("/cache")
        SetValue "LOG_PATH"             ,Server.MapPath("/logs")
        SetValue "TEMP_PATH"            ,Server.MapPath("/temp")
        SetValue "IMAGE_PATH"           ,Server.MapPath("/images")
        SetValue "CLASS_PATH"           ,Server.MapPath("/classes")

        '----------------------------------------------------------
        ' Cache
        '----------------------------------------------------------
        SetValue "CACHE_ENABLED"        ,True
        SetValue "CACHE_HTML"           ,Server.MapPath("/cache/index.html")
        SetValue "CACHE_EXPIRE_MINUTE"  ,60

        '----------------------------------------------------------
        ' HTTP
        '----------------------------------------------------------
        SetValue "HTTP_TIMEOUT_RESOLVE" ,30000
        SetValue "HTTP_TIMEOUT_CONNECT" ,30000
        SetValue "HTTP_TIMEOUT_SEND"    ,60000
        SetValue "HTTP_TIMEOUT_RECEIVE" ,60000

        SetValue "HTTP_USER_AGENT", _
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0 Safari/537.36"

        SetValue "HTTP_ACCEPT"          ,"*/*"
        SetValue "HTTP_ENCODING"        ,"gzip, deflate"
        SetValue "HTTP_LANGUAGE"        ,"tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7"
        SetValue "HTTP_IGNORE_SSL"      ,True

        '----------------------------------------------------------
        ' Import
        '----------------------------------------------------------
        SetValue "IMPORT_BASE_URL"      ,"https://www.genpower.com.tr"
        SetValue "IMPORT_LIST_URL"      ,"https://www.genpower.com.tr/tr/urunler/dizel-jeneratorleri"
        SetValue "IMPORT_DELAY_MS"      ,0
        SetValue "IMPORT_RETRY"         ,3
        SetValue "IMPORT_SAVE_HTML"     ,True

        '----------------------------------------------------------
        ' Database
        '----------------------------------------------------------
        SetValue "DB_PROVIDER"          ,"SQLOLEDB"

        SetValue "DB_SERVER"            ,"localhost"
        SetValue "DB_DATABASE"          ,"InfoPower"
        SetValue "DB_USER"              ,"sa"
        SetValue "DB_PASSWORD"          ,""

        SetValue "DB_TIMEOUT"           ,30

        '----------------------------------------------------------
        ' Tables
        '----------------------------------------------------------
        SetValue "TABLE_PRODUCTS"       ,"products"
        SetValue "TABLE_CATEGORIES"     ,"categories"
        SetValue "TABLE_BRANDS"         ,"brands"
        SetValue "TABLE_IMAGES"         ,"product_images"
        SetValue "TABLE_LOG"            ,"import_logs"

    End Sub

    Public Sub SetValue(Key,Value)

        Key=LCase(Trim(CStr(Key)))

        If mItems.Exists(Key) Then
            mItems(Key)=Value
        Else
            mItems.Add Key,Value
        End If

    End Sub

    Public Function GetValue(Key)

        Key=LCase(Trim(CStr(Key)))

        If mItems.Exists(Key) Then
            GetValue=mItems(Key)
        Else
            GetValue=Null
        End If

    End Function

    Public Function Exists(Key)

        Exists=mItems.Exists(LCase(Trim(CStr(Key))))

    End Function

    Public Sub Remove(Key)

        Key=LCase(Trim(CStr(Key)))

        If mItems.Exists(Key) Then
            mItems.Remove Key
        End If

    End Sub

    Public Sub Clear()

        mItems.RemoveAll

    End Sub

    Public Property Get Count()

        Count=mItems.Count

    End Property

    Public Function ConnectionString()

        ConnectionString = _
            "Provider=" & GetValue("DB_PROVIDER") & ";" & _
            "Data Source=" & GetValue("DB_SERVER") & ";" & _
            "Initial Catalog=" & GetValue("DB_DATABASE") & ";" & _
            "User ID=" & GetValue("DB_USER") & ";" & _
            "Password=" & GetValue("DB_PASSWORD") & ";" & _
            "Persist Security Info=False;"

    End Function

    Public Function CacheFile(FileName)

        CacheFile = GetValue("CACHE_PATH") & "\" & FileName

    End Function

    Public Function LogFile(FileName)

        LogFile = GetValue("LOG_PATH") & "\" & FileName

    End Function

    Public Function TempFile(FileName)

        TempFile = GetValue("TEMP_PATH") & "\" & FileName

    End Function

End Class
%>