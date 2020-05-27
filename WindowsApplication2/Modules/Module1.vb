Imports System.Net
Imports System.IO



Module Module1
    Dim CookieJar As New CookieContainer

    Sub Main()
        Dim arrArg() As String
        arrArg = System.Environment.GetCommandLineArgs()

        Call Form1.Show()

        If arrArg.Length >= 2 Then
            Call Routine(arrArg(1))
        Else
            MsgBox("WTF?")
        End If
    End Sub


    Sub Routine(ByVal strFileName As String)
        Dim sURL As String, strToken As String, strOriLine As String
        Dim intCount As Integer, intPos As Integer
        Dim objStream As Stream

        strFileName = Replace(strFileName, "=", "")

        Try
            'get file extension  
            strToken = GetCurrentToken(strFileName)
            sURL = "https://www.lazada.com.my/ajax/lottery/spinTheWheel/?lang=en&platform=desktop&wheelToken=" & strToken & "&dpr=" & Trim(strFileName)
            'sURL = "http://www.lazada.com.my/ajax/campaign/play/?lang=en&platform=desktop&dpr=" & Trim(strFileName)

            intCount = 0
            While True
                intCount = intCount + 1
                Call Form1.SetMessage("Voucher Miner Processing Count - " & intCount)

                If intCount > 10000 Then
                    strToken = GetCurrentToken(strFileName)
                    sURL = "https://www.lazada.com.my/ajax/lottery/spinTheWheel/?lang=en&platform=desktop&wheelToken=" & strToken & "&dpr=" & Trim(strFileName)
                    'sURL = "http://www.lazada.com.my/ajax/campaign/play/?lang=en&platform=desktop&dpr=" & Trim(strFileName)
                    intCount = 0
                End If


                Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(sURL), HttpWebRequest)
                'Threading.Thread.Sleep(1000)
                Dim res As HttpWebResponse
                myHttpWebRequest.CookieContainer = CookieJar
                'Dim myProxy As New WebProxy("128.199.190.130:8080", True)
                myHttpWebRequest.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; rv:42.0) Gecko/20100101 Firefox/42.0"
                myHttpWebRequest.Referer = "https://www.lazada.com.my/online-revolution-spin-the-wheel/"
                myHttpWebRequest.Host = "www.lazada.com.my"
                myHttpWebRequest.Accept = "*/*"
                myHttpWebRequest.ContentType = "application/x-www-form-urlencoded; charset=UTF-8"
                'myHttpWebRequest.Proxy = vbNullString

                objStream = myHttpWebRequest.GetResponse.GetResponseStream
                res = myHttpWebRequest.GetResponse
                SaveIncomingCookies(res, Trim(strFileName))


                Dim objReader As New StreamReader(objStream)
                Dim sLine As String = ""

                Dim block(300) As Char
                objReader.ReadBlock(block, 0, 300)
                sLine = New String(block)

                If Not sLine Is Nothing Or Trim(sLine) <> "" Then
                    'extract sLine
                    strOriLine = sLine
                    intPos = InStr(sLine, """productsBlockTitle"":", CompareMethod.Text)
                    sLine = sLine.Substring(intPos + Len("""productsBlockTitle"":"))
                    intPos = InStr(sLine, """", CompareMethod.Text)
                    sLine = sLine.Remove(IIf(intPos - 1 >= 0, intPos - 1, 0))
                    sLine = Replace(sLine, """", "")

                    If sLine <> "Exclusive Early Deals! Don't miss out!" Then
                        'intPos = 0
                        'intPos = 1 / intPos
                        System.IO.File.AppendAllText("C:\Lazada\" & strFileName & ".txt", strOriLine & vbCrLf)
                        res.Cookies(0).Expires = DateTime.Now
                        res.Cookies(1).Expires = DateTime.Now
                        res.Cookies(2).Expires = DateTime.Now
                        Exit Sub
                    End If
                End If

                'objReader = Nothing
            End While
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Private Function GetCurrentToken(strFileName As String) As String
        Dim sURL As String, sToken As String
        'Dim wrGETURL As WebRequest
        Dim objStream As Stream
        Dim intPos As Integer

        sURL = "https://www.lazada.com.my/ajax/lottery/settings/?lang=en&platform=desktop&dpr=" & Trim(strFileName)
        'sURL = "http://www.lazada.com.my/ajax/campaign/play/?lang=en&platform=desktop&dpr=" & Trim(strFileName)
        Dim res As HttpWebResponse
        'Dim myProxy As New WebProxy(IProxy, True)
        Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(sURL), HttpWebRequest)
        myHttpWebRequest.CookieContainer = CookieJar
        myHttpWebRequest.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; rv:42.0) Gecko/20100101 Firefox/42.0"
        myHttpWebRequest.Referer = "https://www.lazada.com.my/shopping-kaw-kaw"
        myHttpWebRequest.Host = "www.lazada.com.my"
        myHttpWebRequest.Accept = "*/*"
        myHttpWebRequest.ContentType = "application/x-www-form-urlencoded; charset=UTF-8"
        'myHttpWebRequest.Proxy = myProxy

        objStream = myHttpWebRequest.GetResponse.GetResponseStream()
        res = myHttpWebRequest.GetResponse
        SaveIncomingCookies(res, Trim(strFileName))


        Dim objReader As New StreamReader(objStream)
        Dim sLine As String = ""
        sLine = objReader.ReadLine
        'extract wheeltoken
        intPos = InStr(sLine, "{""wheelToken"":", CompareMethod.Text)
        sLine = sLine.Substring(intPos + Len("{""wheelToken"":"))
        intPos = InStr(sLine, """", CompareMethod.Text)
        sLine = sLine.Remove(intPos - 1)
        sToken = sLine

        'clear all cookies
        res.Cookies(0).Expires = DateTime.Now
        res.Cookies(1).Expires = DateTime.Now
        res.Cookies(2).Expires = DateTime.Now
        res.Cookies(3).Expires = DateTime.Now
        res.Cookies(4).Expires = DateTime.Now
        res.Cookies(5).Expires = DateTime.Now

        Return sToken
    End Function

    Private Sub SaveIncomingCookies(ByRef response As HttpWebResponse, ByVal dpr As Integer)
        If response.Headers("Set-Cookie") <> Nothing Then
            'CookieJar.SetCookies(New Uri("http://www.lazada.com.my/ajax/campaign/play/?lang=en&platform=desktop&dpr=" & dpr), response.Headers("Set-Cookie"))
            CookieJar.SetCookies(New Uri("https://www.lazada.com.my/ajax/lottery/play/?lang=en&platform=desktop&dpr=" & dpr), response.Headers("Set-Cookie"))
        End If
    End Sub
End Module
