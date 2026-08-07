<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.main.asp
' Part    : 1
'====================================================================
Option Explicit

Class Importer

    Private mConfig
    Private mHttp
    Private mParser
    Private mDatabase
    Private mLogger
    Private mHistory
    Private mQueue
    Private mImage
    Private mPdf

    Private mProducts
    Private mStarted
    Private mFinished

    Private Sub Class_Initialize()

        Set mConfig   = New Config
        Set mHttp     = New HttpClient
        Set mParser   = New Parser
        Set mDatabase = New Database
        Set mLogger   = New Logger
        Set mHistory  = New History
        Set mQueue    = New Queue
        Set mImage    = New ImageManager
        Set mPdf      = New PdfManager

        ReDim mProducts(-1)

    End Sub



    Private Sub Class_Terminate()

        Set mConfig   = Nothing
        Set mHttp     = Nothing
        Set mParser   = Nothing
        Set mDatabase = Nothing
        Set mLogger   = Nothing
        Set mHistory  = Nothing
        Set mQueue    = Nothing
        Set mImage    = Nothing
        Set mPdf      = Nothing

    End Sub



    Public Property Get Config()

        Set Config = mConfig

    End Property



    Public Property Get Database()

        Set Database = mDatabase

    End Property



    Public Property Get Logger()

        Set Logger = mLogger

    End Property



    Public Property Get Parser()

        Set Parser = mParser

    End Property



    Public Property Get Queue()

        Set Queue = mQueue

    End Property



    Public Function LoadUrl(Url)

        Dim Resp

        Set Resp = mHttp.Get(Url)

        If Resp Is Nothing Then

            LoadUrl = False
            Exit Function

        End If

        If Not Resp.Success Then

            LoadUrl = False
            Exit Function

        End If

        LoadUrl = mParser.Load(Resp.Text)

    End Function



    Public Function LoadFile(FileName)

        LoadFile = mParser.LoadFile(FileName)

    End Function



    Public Function Parse()

        Parse = mParser.ParseCards()

    End Function
 
     Public Function BuildQueue()

        Dim i
        Dim P

        mQueue.Clear

        For i = 0 To mParser.ProductCount() - 1

            Set P = mParser.Product(i)

            mQueue.Enqueue P

        Next

        BuildQueue = mQueue.Count()

    End Function



    Public Function ProductCount()

        ProductCount = mQueue.Count()

    End Function



    Public Function ImportAll()

        Dim Item

        ImportAll = 0

        mLogger.SessionStart

        mStarted = Now()

        mQueue.Reset

        Do While mQueue.HasNext()

            Set Item = mQueue.MoveNext()

            If Not Item Is Nothing Then

                If ImportProduct(Item) Then
                    ImportAll = ImportAll + 1
                End If

            End If

        Loop

        mFinished = Now()

        mLogger.SessionEnd

    End Function



    Private Function ImportProduct(ByRef Item)

        Dim Product

        Set Product = New Product

        Product.Id = CLng("" & Item("id"))

        Product.Title = Item("title")
        Product.Url = Item("url")
        Product.Image = Item("image")
        Product.Model = Item("model")
        Product.Power = Item("power")
        Product.Engine = Item("engine")
        Product.Alternator = Item("alternator")
        Product.Frequency = Item("frequency")
        Product.RPM = Item("rpm")

        ImportProduct = SaveProduct(Product)

    End Function



    Private Function SaveProduct(ByRef Product)

        Dim SQL

        SQL = ""
        SQL = SQL & "INSERT INTO products("
        SQL = SQL & "title,url,image,model,power,"
        SQL = SQL & "engine,alternator,frequency,rpm"
        SQL = SQL & ") VALUES ("
        SQL = SQL & mDatabase.Escape(Product.Title) & ","
        SQL = SQL & mDatabase.Escape(Product.Url) & ","
        SQL = SQL & mDatabase.Escape(Product.Image) & ","
        SQL = SQL & mDatabase.Escape(Product.Model) & ","
        SQL = SQL & mDatabase.Escape(Product.Power) & ","
        SQL = SQL & mDatabase.Escape(Product.Engine) & ","
        SQL = SQL & mDatabase.Escape(Product.Alternator) & ","
        SQL = SQL & mDatabase.Escape(Product.Frequency) & ","
        SQL = SQL & mDatabase.Escape(Product.RPM)
        SQL = SQL & ")"

        If mDatabase.Execute(SQL) Then

            mHistory.Success Product.Id, Product.Title

            mLogger.Info Product.Title

            SaveProduct = True

        Else

            mHistory.Failed Product.Id, mDatabase.LastError

            mLogger.Error mDatabase.LastError

            SaveProduct = False

        End If

    End Function
 
     Public Function DownloadImages()

        Dim i
        Dim Item

        DownloadImages = 0

        For i = 0 To mParser.ProductCount() - 1

            Set Item = mParser.Product(i)

            If Len(Trim("" & Item("image"))) > 0 Then

                If mImage.Save(Item("image")) <> "" Then
                    DownloadImages = DownloadImages + 1
                End If

            End If

        Next

    End Function



    Public Function DownloadPdfFiles()

        Dim i
        Dim Item

        DownloadPdfFiles = 0

        For i = 0 To mParser.ProductCount() - 1

            Set Item = mParser.Product(i)

            If Item.Exists("pdf") Then

                If Len(Trim("" & Item("pdf"))) > 0 Then

                    If mPdf.Save(Item("pdf")) <> "" Then
                        DownloadPdfFiles = DownloadPdfFiles + 1
                    End If

                End If

            End If

        Next

    End Function



    Public Function StartTime()

        StartTime = mStarted

    End Function



    Public Function FinishTime()

        FinishTime = mFinished

    End Function



    Public Function ElapsedSeconds()

        If IsDate(mStarted) And IsDate(mFinished) Then
            ElapsedSeconds = DateDiff("s", mStarted, mFinished)
        Else
            ElapsedSeconds = 0
        End If

    End Function



    Public Sub Reset()

        mQueue.Clear

        ReDim mProducts(-1)

        mStarted = Empty
        mFinished = Empty

    End Sub



    Public Sub Close()

        mDatabase.Close

    End Sub

End Class
%>