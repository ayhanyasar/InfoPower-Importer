<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.http.asp
' Version : 3.0.0
'====================================================================
Option Explicit

Class HttpResponse

    Public Success
    Public Status
    Public StatusText
    Public Headers
    Public Body
    Public Bytes
    Public Url
    Public ErrorMessage

    Private Sub Class_Initialize()
        Success=False
        Status=0
        StatusText=""
        Url=""
        Body=""
        ErrorMessage=""
        Set Headers=Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
        If IsObject(Headers) Then
            Headers.RemoveAll
            Set Headers=Nothing
        End If
    End Sub

End Class



Class HttpClient

    Private mTimeoutResolve
    Private mTimeoutConnect
    Private mTimeoutSend
    Private mTimeoutReceive

    Private mUserAgent
    Private mAccept
    Private mAcceptLanguage
    Private mAcceptEncoding

    Private mProxy
    Private mIgnoreSSL
    Private mFollowRedirect

    Private mHeaders
    Private mCookies

    Private Sub Class_Initialize()

        mTimeoutResolve = 30000
        mTimeoutConnect = 30000
        mTimeoutSend = 60000
        mTimeoutReceive = 60000

        mUserAgent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0 Safari/537.36"

        mAccept="*/*"
        mAcceptLanguage="tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7"
        mAcceptEncoding="gzip, deflate"

        mIgnoreSSL=True
        mFollowRedirect=True

        Set mHeaders=Server.CreateObject("Scripting.Dictionary")
        Set mCookies=Server.CreateObject("Scripting.Dictionary")

    End Sub

    Private Sub Class_Terminate()

        If IsObject(mHeaders) Then
            mHeaders.RemoveAll
            Set mHeaders=Nothing
        End If

        If IsObject(mCookies) Then
            mCookies.RemoveAll
            Set mCookies=Nothing
        End If

    End Sub


    Public Property Let UserAgent(v)
        mUserAgent=CStr(v)
    End Property

    Public Property Get UserAgent()
        UserAgent=mUserAgent
    End Property


    Public Property Let Accept(v)
        mAccept=CStr(v)
    End Property

    Public Property Get Accept()
        Accept=mAccept
    End Property


    Public Property Let AcceptLanguage(v)
        mAcceptLanguage=CStr(v)
    End Property

    Public Property Get AcceptLanguage()
        AcceptLanguage=mAcceptLanguage
    End Property


    Public Property Let AcceptEncoding(v)
        mAcceptEncoding=CStr(v)
    End Property

    Public Property Get AcceptEncoding()
        AcceptEncoding=mAcceptEncoding
    End Property


    Public Property Let IgnoreSSL(v)
        mIgnoreSSL=CBool(v)
    End Property

    Public Property Get IgnoreSSL()
        IgnoreSSL=mIgnoreSSL
    End Property


    Public Property Let FollowRedirect(v)
        mFollowRedirect=CBool(v)
    End Property

    Public Property Get FollowRedirect()
        FollowRedirect=mFollowRedirect
    End Property


    Public Sub SetTimeouts(resolveMs,connectMs,sendMs,receiveMs)

        mTimeoutResolve=CLng(resolveMs)
        mTimeoutConnect=CLng(connectMs)
        mTimeoutSend=CLng(sendMs)
        mTimeoutReceive=CLng(receiveMs)

    End Sub


    Public Sub SetProxy(proxy)

        mProxy=Trim(CStr(proxy))

    End Sub


    Public Sub ClearProxy()

        mProxy=""

    End Sub


    Public Sub AddHeader(name,value)

        If mHeaders.Exists(name) Then
            mHeaders(name)=value
        Else
            mHeaders.Add name,value
        End If

    End Sub


    Public Sub RemoveHeader(name)

        If mHeaders.Exists(name) Then
            mHeaders.Remove name
        End If

    End Sub


    Public Sub ClearHeaders()

        mHeaders.RemoveAll

    End Sub


    Public Sub AddCookie(name,value)

        If mCookies.Exists(name) Then
            mCookies(name)=value
        Else
            mCookies.Add name,value
        End If

    End Sub


    Public Sub ClearCookies()

        mCookies.RemoveAll

    End Sub


    Public Function Get(url)

        Set Get=Request("GET",url,"")

    End Function


    Public Function Delete(url)

        Set Delete=Request("DELETE",url,"")

    End Function


    Public Function Post(url,data)

        Set Post=Request("POST",url,data)

    End Function


    Public Function Put(url,data)

        Set Put=Request("PUT",url,data)

    End Function


    Public Function Patch(url,data)

        Set Patch=Request("PATCH",url,data)

    End Function


    Public Function Head(url)

        Set Head=Request("HEAD",url,"")

    End Function


    Public Function Request(method,url,body)

        Dim oHttp
        Dim oResp
        Dim k
        Dim arr
        Dim headerText
        Dim lines
        Dim line
        Dim pos

        Set oResp=New HttpResponse

        On Error Resume Next

        Set oHttp=Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

        If Err.Number<>0 Then

            oResp.ErrorMessage=Err.Description
            Set Request=oResp
            Exit Function

        End If

        oHttp.setTimeouts _
            mTimeoutResolve, _
            mTimeoutConnect, _
            mTimeoutSend, _
            mTimeoutReceive

        oHttp.Open method,url,False

        If mIgnoreSSL Then

            On Error Resume Next
            oHttp.setOption 2,13056

        End If

        If Len(mProxy)>0 Then

            On Error Resume Next
            oHttp.setProxy 2,mProxy

        End If

        oHttp.setRequestHeader "User-Agent",mUserAgent
        oHttp.setRequestHeader "Accept",mAccept
        oHttp.setRequestHeader "Accept-Language",mAcceptLanguage
        oHttp.setRequestHeader "Accept-Encoding",mAcceptEncoding
        oHttp.setRequestHeader "Connection","Keep-Alive"

        For Each k In mHeaders.Keys

            oHttp.setRequestHeader CStr(k),CStr(mHeaders(k))

        Next

        If mCookies.Count>0 Then

            oHttp.setRequestHeader "Cookie",CookieString()

        End If

        Err.Clear

        oHttp.send body

        If Err.Number<>0 Then

            oResp.ErrorMessage=Err.Description

            Set Request=oResp

            Exit Function

        End If

        oResp.Status=oHttp.status
        oResp.StatusText=oHttp.statusText
        oResp.Url=url

        If oResp.Status>=200 And oResp.Status<300 Then
            oResp.Success=True
        End If

        On Error Resume Next

        oResp.Body=oHttp.responseText

        If Err.Number=0 Then
            oResp.Bytes=oHttp.responseBody
        End If

        headerText=oHttp.getAllResponseHeaders

        lines=Split(headerText,vbCrLf)

        For Each line In lines

            pos=InStr(line,":")

            If pos>0 Then

                oResp.Headers(Trim(Left(line,pos-1)))=Trim(Mid(line,pos+1))

            End If

        Next

        ParseCookies oResp

        Set oHttp=Nothing

        Set Request=oResp

    End Function


    Private Function CookieString()

        Dim s
        Dim k

        s=""

        For Each k In mCookies.Keys

            If Len(s)>0 Then s=s & "; "

            s=s & k & "=" & mCookies(k)

        Next

        CookieString=s

    End Function


    Private Sub ParseCookies(oResp)

        Dim h

        For Each h In oResp.Headers.Keys

            If LCase(h)="set-cookie" Then

                Dim p
                Dim c

                c=Split(oResp.Headers(h),";")(0)

                p=InStr(c,"=")

                If p>0 Then

                    AddCookie _
                        Trim(Left(c,p-1)), _
                        Trim(Mid(c,p+1))

                End If

            End If

        Next

    End Sub

End Class
%>