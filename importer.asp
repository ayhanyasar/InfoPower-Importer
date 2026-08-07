<%
Option Explicit

Response.Buffer = True
Response.Charset = "utf-8"
Server.ScriptTimeout = 3600

<!--#include file="classes/class.config.asp"-->
<!--#include file="classes/class.utility.asp"-->
<!--#include file="classes/class.http.asp"-->
<!--#include file="classes/class.database.asp"-->
<!--#include file="classes/class.logger.asp"-->
<!--#include file="classes/class.cache.asp"-->
<!--#include file="classes/class.dom.asp"-->
<!--#include file="classes/class.parser.asp"-->
<!--#include file="classes/class.product.asp"-->
<!--#include file="classes/class.spec.asp"-->
<!--#include file="classes/class.image.asp"-->
<!--#include file="classes/class.pdf.asp"-->
<!--#include file="classes/class.queue.asp"-->
<!--#include file="classes/class.history.asp"-->
<!--#include file="classes/class.main.asp"-->

Dim Importer
Dim SourceUrl
Dim ImportedCount
Dim ImageCount
Dim PdfCount

Set Importer = New Importer

Importer.Logger.SessionStart

Importer.Database.ConnectionString = Importer.Config.ConnectionString

SourceUrl = Importer.Config.Get("SOURCE_URL")

If Len(SourceUrl) = 0 Then
    SourceUrl = "https://www.genpower.com.tr/tr/urunler/dizel-jeneratorleri"
End If

Response.Write "<h2>InfoPower Importer PRO v3.0</h2>"
Response.Write "<hr>"

Response.Write "Source : " & Server.HTMLEncode(SourceUrl) & "<br>"

If Not Importer.LoadUrl(SourceUrl) Then

    Response.Write "<b>ERROR :</b> Kaynak sayfa okunamadı.<br>"

    Importer.Logger.Error "Source page could not be loaded."

    Importer.Logger.SessionEnd

    Set Importer = Nothing

    Response.End

End If

Response.Write "HTML Loaded.<br>"

Response.Flush

Importer.Parse

Response.Write "Products Found : " & Importer.Parser.ProductCount() & "<br>"

Response.Flush

Importer.BuildQueue

ImportedCount = Importer.ImportAll()

Response.Write "Imported : " & ImportedCount & "<br>"

Response.Flush

ImageCount = Importer.DownloadImages()

Response.Write "Images : " & ImageCount & "<br>"

Response.Flush

PdfCount = Importer.DownloadPdfFiles()

Response.Write "PDF Files : " & PdfCount & "<br>"

Response.Write "Elapsed : " & Importer.ElapsedSeconds() & " sec<br>"

Importer.Logger.SessionEnd

Importer.Close

Set Importer = Nothing

Response.Write "<hr>"
Response.Write "<b>IMPORT COMPLETED</b>"
%>