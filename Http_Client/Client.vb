Imports System
Imports SeasideResearch.LibCurlNet
Imports System.Net
Imports System.IO

Public Structure InfoSend
    Private _CURLINFO_CONNECT_TIME As Double
    Public Property CURLINFO_CONNECT_TIME() As Double
        Get
            Return _CURLINFO_CONNECT_TIME
        End Get
        Set(ByVal value As Double)
            _CURLINFO_CONNECT_TIME = value
        End Set
    End Property

    Private _CURLINFO_CONTENT_LENGTH_DOWNLOAD As Double
    Public Property CURLINFO_CONTENT_LENGTH_DOWNLOAD() As Double
        Get
            Return _CURLINFO_CONTENT_LENGTH_DOWNLOAD
        End Get
        Set(ByVal value As Double)
            _CURLINFO_CONTENT_LENGTH_DOWNLOAD = value
        End Set
    End Property

    Private _CURLINFO_CONTENT_LENGTH_UPLOAD As Double
    Public Property CURLINFO_CONTENT_LENGTH_UPLOAD() As Double
        Get
            Return _CURLINFO_CONTENT_LENGTH_UPLOAD
        End Get
        Set(ByVal value As Double)
            _CURLINFO_CONTENT_LENGTH_UPLOAD = value
        End Set
    End Property

    Private _CURLINFO_CONTENT_TYPE As String
    Public Property CURLINFO_CONTENT_TYPE() As String
        Get
            Return _CURLINFO_CONTENT_TYPE
        End Get
        Set(ByVal value As String)
            _CURLINFO_CONTENT_TYPE = value
        End Set
    End Property

    Private _CURLINFO_FILETIME As Date
    Public Property CURLINFO_FILETIME() As Date
        Get
            Return _CURLINFO_FILETIME
        End Get
        Set(ByVal value As Date)
            _CURLINFO_FILETIME = value
        End Set
    End Property

    Private _CURLINFO_HEADER_SIZE As Integer
    Public Property CURLINFO_HEADER_SIZE() As Integer
        Get
            Return _CURLINFO_HEADER_SIZE
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_HEADER_SIZE = value
        End Set
    End Property

    Private _CURLINFO_HTTPAUTH_AVAIL As Integer
    Public Property CURLINFO_HTTPAUTH_AVAIL() As Integer
        Get
            Return _CURLINFO_HTTPAUTH_AVAIL
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_HTTPAUTH_AVAIL = value
        End Set
    End Property

    Private _CURLINFO_HTTP_CONNECTCODE As Integer
    Public Property CURLINFO_HTTP_CONNECTCODE() As Integer
        Get
            Return _CURLINFO_HTTP_CONNECTCODE
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_HTTP_CONNECTCODE = value
        End Set
    End Property

    Private _CURLINFO_NAMELOOKUP_TIME As Double
    Public Property CURLINFO_NAMELOOKUP_TIME() As Double
        Get
            Return _CURLINFO_NAMELOOKUP_TIME
        End Get
        Set(ByVal value As Double)
            _CURLINFO_NAMELOOKUP_TIME = value
        End Set
    End Property

    Private _CURLINFO_OS_ERRNO As Integer
    Public Property CURLINFO_OS_ERRNO() As Integer
        Get
            Return _CURLINFO_OS_ERRNO
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_OS_ERRNO = value
        End Set
    End Property

    Private _CURLINFO_PRETRANSFER_TIME As Double
    Public Property CURLINFO_PRETRANSFER_TIME() As Double
        Get
            Return _CURLINFO_PRETRANSFER_TIME
        End Get
        Set(ByVal value As Double)
            _CURLINFO_PRETRANSFER_TIME = value
        End Set
    End Property

    Private _CURLINFO_PRIVATE As Object
    Public Property CURLINFO_PRIVATE() As Object
        Get
            Return _CURLINFO_PRIVATE
        End Get
        Set(ByVal value As Object)
            _CURLINFO_PRIVATE = value
        End Set
    End Property

    Private _CURLINFO_PROXYAUTH_AVAIL As Integer
    Public Property CURLINFO_PROXYAUTH_AVAIL() As Integer
        Get
            Return _CURLINFO_PROXYAUTH_AVAIL
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_PROXYAUTH_AVAIL = value
        End Set
    End Property

    Private _CURLINFO_REDIRECT_COUNT As Integer
    Public Property CURLINFO_REDIRECT_COUNT() As Integer
        Get
            Return _CURLINFO_REDIRECT_COUNT
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_REDIRECT_COUNT = value
        End Set
    End Property

    Private _CURLINFO_REDIRECT_TIME As Double
    Public Property CURLINFO_REDIRECT_TIME() As Double
        Get
            Return _CURLINFO_REDIRECT_TIME
        End Get
        Set(ByVal value As Double)
            _CURLINFO_REDIRECT_TIME = value
        End Set
    End Property

    Private _CURLINFO_REQUEST_SIZE As Integer
    Public Property CURLINFO_REQUEST_SIZE() As Integer
        Get
            Return _CURLINFO_REQUEST_SIZE
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_REQUEST_SIZE = value
        End Set
    End Property

    Private _CURLINFO_RESPONSE_CODE As Integer
    Public Property CURLINFO_RESPONSE_CODE() As Integer
        Get
            Return _CURLINFO_RESPONSE_CODE
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_RESPONSE_CODE = value
        End Set
    End Property

    Private _CURLINFO_SIZE_DOWNLOAD As Double
    Public Property CURLINFO_SIZE_DOWNLOAD() As Double
        Get
            Return _CURLINFO_SIZE_DOWNLOAD
        End Get
        Set(ByVal value As Double)
            _CURLINFO_SIZE_DOWNLOAD = value
        End Set
    End Property

    Private _CURLINFO_SIZE_UPLOAD As Double
    Public Property CURLINFO_SIZE_UPLOAD() As Double
        Get
            Return _CURLINFO_SIZE_UPLOAD
        End Get
        Set(ByVal value As Double)
            _CURLINFO_SIZE_UPLOAD = value
        End Set
    End Property

    Private _CURLINFO_SPEED_DOWNLOAD As Double
    Public Property CURLINFO_SPEED_DOWNLOAD() As Double
        Get
            Return _CURLINFO_SPEED_DOWNLOAD
        End Get
        Set(ByVal value As Double)
            _CURLINFO_SPEED_DOWNLOAD = value
        End Set
    End Property

    Private _CURLINFO_SPEED_UPLOAD As Double
    Public Property CURLINFO_SPEED_UPLOAD() As Double
        Get
            Return _CURLINFO_SPEED_UPLOAD
        End Get
        Set(ByVal value As Double)
            _CURLINFO_SPEED_UPLOAD = value
        End Set
    End Property

    Private _CURLINFO_SSL_VERIFYRESULT As Integer
    Public Property CURLINFO_SSL_VERIFYRESULT() As Integer
        Get
            Return _CURLINFO_SSL_VERIFYRESULT
        End Get
        Set(ByVal value As Integer)
            _CURLINFO_SSL_VERIFYRESULT = value
        End Set
    End Property

    Private _CURLINFO_STARTTRANSFER_TIME As Double
    Public Property CURLINFO_STARTTRANSFER_TIME() As Double
        Get
            Return _CURLINFO_STARTTRANSFER_TIME
        End Get
        Set(ByVal value As Double)
            _CURLINFO_STARTTRANSFER_TIME = value
        End Set
    End Property

    Private _CURLINFO_TOTAL_TIME As Double
    Public Property CURLINFO_TOTAL_TIME() As Double
        Get
            Return _CURLINFO_TOTAL_TIME
        End Get
        Set(ByVal value As Double)
            _CURLINFO_TOTAL_TIME = value
        End Set
    End Property

End Structure

Public Class Client

    Public Event Progress(ByVal _total_bajar As Double, ByVal _total_bajado As Double, ByVal _total_subir As Double, ByVal _total_subido As Double)

    Private _post_fields_claves As Collection = New Collection
    Private _post_fields_valores As Collection = New Collection

    'Private _get As WebRequest

    Private _post_field_file_claves As Collection = New Collection
    Private _post_field_file_valores As Collection = New Collection

    Private _infcurl As InfoSend
    ''' <summary>
    ''' Retorna o establece informacion de la petición recientemente realizada
    ''' </summary>
    ''' <value>Estructura InfoSend</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Info() As InfoSend
        Get
            Return _infcurl
        End Get
        Set(ByVal value As InfoSend)
            _infcurl = value
        End Set
    End Property

    Private _respuesta As String
    ''' <summary>
    ''' Contiene el resultado del la petición http
    ''' </summary>
    ''' <value></value>
    ''' <returns>String de la respuesta</returns>
    ''' <remarks></remarks>
    Public Property Http_Response() As String
        Get
            Return _respuesta
        End Get
        Set(ByVal value As String)
            _respuesta = value
        End Set
    End Property

    Public Sub New()
        _respuesta = Nothing
    End Sub



    ''' <summary>
    ''' Agrega campos para peticiones GET o POST
    ''' </summary>
    ''' <param name="clave">Clave del campo</param>
    ''' <param name="valor">Valor del campo</param>
    ''' <remarks></remarks>
    Public Sub Add_Field(ByVal clave, ByVal valor)

        Try

            _post_fields_claves.Add(clave)
            _post_fields_valores.Add(valor)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' Agrega campos de tipo Multipart/Data para el envio de archivos a travez de HTTP POST
    ''' </summary>
    ''' <param name="clave">Clave del campo</param>
    ''' <param name="valor">Valor del campo</param>
    ''' <remarks></remarks>
    Public Sub Add_File_Field(ByVal clave, ByVal valor)

        Try

            _post_field_file_claves.Add(clave)
            _post_field_file_valores.Add(valor)

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    Private Sub _clear_fields()

        Try

            _post_fields_claves.Clear()
            _post_fields_valores.Clear()

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    Private Function _build_get() As String
        Dim post As String
        post = ""

        Try

            If _post_fields_claves.Count < 1 Then
                Return Nothing
                Exit Function
            End If

            For i As Integer = 1 To CInt(_post_fields_valores.Count)
                post += _post_fields_claves.Item(i) + "=" + _post_fields_valores.Item(i)
                If (Not (post = Nothing)) Then
                    post += "&"
                End If
            Next i

            Return "?" + post.Trim("&")

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    Private Function _build_post() As String
        Dim post As String
        post = ""

        Try

            If _post_fields_claves.Count < 1 Then
                Return Nothing
                Exit Function
            End If

            For i As Integer = 1 To CInt(_post_fields_valores.Count)
                post += _post_fields_claves.Item(i) + "=" + _post_fields_valores.Item(i)
                If (Not (post = Nothing)) Then
                    post += "&"
                End If
            Next i

            Return post.Trim("&")

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    ''' <summary>
    ''' Realiza la petición via HTTP GET
    ''' </summary>
    ''' <param name="url">URL de la petición</param>
    ''' <returns>Respuesta del servidor</returns>
    ''' <remarks></remarks>
    Public Function HTTP_GET(ByVal url As String) As String
        'Dim enc As New System.Text.ASCIIEncoding
        'Dim _res As WebResponse
        Dim u As New Uri(url & "?" & _build_post())

        _respuesta = Nothing

        Curl.GlobalInit(CURLinitFlag.CURL_GLOBAL_ALL)

        Dim easy As Easy
        easy = New Easy

        ' Set up write delegate
        Dim wf As Easy.WriteFunction
        wf = New Easy.WriteFunction(AddressOf _get_respuesta)

        ' and the rest of the cURL options
        easy.SetOpt(CURLoption.CURLOPT_URL, u.AbsoluteUri)
        easy.SetOpt(CURLoption.CURLOPT_FOLLOWLOCATION, True)
        easy.SetOpt(CURLoption.CURLOPT_AUTOREFERER, True)
        easy.SetOpt(CURLoption.CURLOPT_WRITEFUNCTION, wf)

        easy.Perform()

        easy.GetInfo(CURLINFO.CURLINFO_CONNECT_TIME, _infcurl.CURLINFO_CONNECT_TIME)
        easy.GetInfo(CURLINFO.CURLINFO_CONTENT_LENGTH_DOWNLOAD, _infcurl.CURLINFO_CONTENT_LENGTH_DOWNLOAD)
        easy.GetInfo(CURLINFO.CURLINFO_CONTENT_LENGTH_UPLOAD, _infcurl.CURLINFO_CONTENT_LENGTH_UPLOAD)
        easy.GetInfo(CURLINFO.CURLINFO_CONTENT_TYPE, _infcurl.CURLINFO_CONTENT_TYPE)
        easy.GetInfo(CURLINFO.CURLINFO_FILETIME, _infcurl.CURLINFO_FILETIME)
        easy.GetInfo(CURLINFO.CURLINFO_HEADER_SIZE, _infcurl.CURLINFO_HEADER_SIZE)
        easy.GetInfo(CURLINFO.CURLINFO_HTTPAUTH_AVAIL, _infcurl.CURLINFO_HTTPAUTH_AVAIL)
        easy.GetInfo(CURLINFO.CURLINFO_HTTP_CONNECTCODE, _infcurl.CURLINFO_HTTP_CONNECTCODE)
        easy.GetInfo(CURLINFO.CURLINFO_NAMELOOKUP_TIME, _infcurl.CURLINFO_NAMELOOKUP_TIME)
        easy.GetInfo(CURLINFO.CURLINFO_OS_ERRNO, _infcurl.CURLINFO_OS_ERRNO)
        easy.GetInfo(CURLINFO.CURLINFO_PRETRANSFER_TIME, _infcurl.CURLINFO_PRETRANSFER_TIME)
        easy.GetInfo(CURLINFO.CURLINFO_PRIVATE, _infcurl.CURLINFO_PRIVATE)
        easy.GetInfo(CURLINFO.CURLINFO_PROXYAUTH_AVAIL, _infcurl.CURLINFO_PROXYAUTH_AVAIL)
        easy.GetInfo(CURLINFO.CURLINFO_REDIRECT_COUNT, _infcurl.CURLINFO_REDIRECT_COUNT)
        easy.GetInfo(CURLINFO.CURLINFO_REDIRECT_TIME, _infcurl.CURLINFO_REDIRECT_TIME)
        easy.GetInfo(CURLINFO.CURLINFO_REQUEST_SIZE, _infcurl.CURLINFO_REQUEST_SIZE)
        easy.GetInfo(CURLINFO.CURLINFO_RESPONSE_CODE, _infcurl.CURLINFO_RESPONSE_CODE)
        easy.GetInfo(CURLINFO.CURLINFO_SIZE_DOWNLOAD, _infcurl.CURLINFO_SIZE_DOWNLOAD)
        easy.GetInfo(CURLINFO.CURLINFO_SIZE_UPLOAD, _infcurl.CURLINFO_SIZE_UPLOAD)
        easy.GetInfo(CURLINFO.CURLINFO_SPEED_DOWNLOAD, _infcurl.CURLINFO_SPEED_DOWNLOAD)
        easy.GetInfo(CURLINFO.CURLINFO_SPEED_UPLOAD, _infcurl.CURLINFO_SPEED_UPLOAD)
        easy.GetInfo(CURLINFO.CURLINFO_SSL_VERIFYRESULT, _infcurl.CURLINFO_SSL_VERIFYRESULT)
        easy.GetInfo(CURLINFO.CURLINFO_STARTTRANSFER_TIME, _infcurl.CURLINFO_STARTTRANSFER_TIME)
        easy.GetInfo(CURLINFO.CURLINFO_TOTAL_TIME, _infcurl.CURLINFO_TOTAL_TIME)

        easy.Cleanup()

        Curl.GlobalCleanup()

        '_get = WebRequest.Create(url & _build_post())
        'CType(_get, HttpWebRequest).UserAgent = "Mozilla 4.0 (compatible; MSIE 6.0; Win32)"
        '_res = _get.GetResponse()
        '_respuesta = New StreamReader(_res.GetResponseStream()).ReadToEnd()
        '_res.Close()


        _clear_fields()

        Return _respuesta

    End Function

    Public Function HTTP_THREAD_GET(ByVal url As String) As String
        'Dim enc As New System.Text.ASCIIEncoding
        'Dim _res As WebResponse
        Dim u As New Uri(url & "?" & _build_post())

        _respuesta = Nothing

        Curl.GlobalInit(CURLinitFlag.CURL_GLOBAL_ALL)

        Dim easy As Easy
        easy = New Easy

        ' Set up write delegate
        Dim wf As Easy.WriteFunction
        wf = New Easy.WriteFunction(AddressOf _get_respuesta)

        ' and the rest of the cURL options
        easy.SetOpt(CURLoption.CURLOPT_URL, u.AbsoluteUri)
        easy.SetOpt(CURLoption.CURLOPT_FOLLOWLOCATION, True)
        easy.SetOpt(CURLoption.CURLOPT_AUTOREFERER, True)
        easy.SetOpt(CURLoption.CURLOPT_WRITEFUNCTION, wf)

        Dim multi As Multi
        multi = New Multi()
        multi.AddHandle(easy)

        Dim stillRunning As Int32
        stillRunning = 1
        ' call Multi.Perform right away
        While (multi.Perform(stillRunning) = _
            CURLMcode.CURLM_CALL_MULTI_PERFORM)
        End While

        Dim rc As Int32
        While (stillRunning <> 0)
            multi.FDSet()
            rc = multi.Select(1000)
            If (rc = -1) Then
                stillRunning = 0
                Exit While
            End If
            ' call multi.Perform
            While (multi.Perform(stillRunning) = _
                CURLMcode.CURLM_CALL_MULTI_PERFORM)
            End While
        End While

        multi.Cleanup()
        easy.Cleanup()
        Curl.GlobalCleanup()

        '_get = WebRequest.Create(url & _build_post())
        'CType(_get, HttpWebRequest).UserAgent = "Mozilla 4.0 (compatible; MSIE 6.0; Win32)"
        '_res = _get.GetResponse()
        '_respuesta = New StreamReader(_res.GetResponseStream()).ReadToEnd()
        '_res.Close()


        _clear_fields()

        Return _respuesta

    End Function

    ''' <summary>
    ''' Envio petición HTTP POST
    ''' </summary>
    ''' <param name="url">URL de la petición</param>
    ''' <returns>Respuesta del servidor</returns>
    ''' <remarks></remarks>
    Public Function HTTP_POST(ByVal url As String) As String
        'Dim dt As Byte()
        'Dim strm As Stream
        'Dim _res As WebResponse
        Dim u As New Uri("http://test/test.php?" + _build_post())
        Dim query As String = ""

        query = u.Query

        _respuesta = Nothing

        Curl.GlobalInit(CURLinitFlag.CURL_GLOBAL_ALL)

        Dim easy As Easy
        easy = New Easy

        ' Set up write delegate
        Dim wf As Easy.WriteFunction
        wf = New Easy.WriteFunction(AddressOf _get_respuesta)
        easy.SetOpt(CURLoption.CURLOPT_WRITEFUNCTION, wf)

        easy.SetOpt(CURLoption.CURLOPT_AUTOREFERER, True)
        easy.SetOpt(CURLoption.CURLOPT_URL, url)
        easy.SetOpt(CURLoption.CURLOPT_FOLLOWLOCATION, True)
        easy.SetOpt(CURLoption.CURLOPT_USERAGENT, "Mozilla/5.0 (Windows; U; Windows NT 5.1; it; rv:1.9.0.6; .NET CLR 3.0; ffco7) Gecko/2009011913 Firefox/3.0.6")
        easy.SetOpt(CURLoption.CURLOPT_POSTFIELDS, query.Replace("?", ""))
        easy.SetOpt(CURLoption.CURLOPT_POST, True)

        easy.Perform()

        easy.Cleanup()
        Curl.GlobalCleanup()

        '_get = WebRequest.Create(url)
        '_get.Credentials = CredentialCache.DefaultCredentials
        'CType(_get, HttpWebRequest).UserAgent = "Mozilla 4.0 (compatible; MSIE 6.0; Win32)"
        '_get.Method = "POST"

        'dt = Text.Encoding.UTF8.GetBytes(_build_post())
        '_get.ContentLength = dt.Length
        '_get.ContentType = "application/x-www-form-urlencoded"

        'strm = _get.GetRequestStream()

        'strm.Write(dt, 0, dt.Length)
        'strm.Close()

        'Dim a As Object = Text.Encoding.UTF8.GetString(dt)

        '_get.Method = "POST"
        '_res = _get.GetResponse()
        '_respuesta = New StreamReader(_res.GetResponseStream()).ReadToEnd()
        '_res.Close()

        _clear_fields()

        Return _respuesta

    End Function

    ''' <summary>
    ''' Envia petición HTTP POST para envio de archivos
    ''' </summary>
    ''' <param name="url">URL de la petición</param>
    ''' <returns>Respuesta del servidor</returns>
    ''' <remarks></remarks>
    Public Function POST_UPLOAD(ByVal url As String) As String
        Try

            _respuesta = Nothing

            Curl.GlobalInit(CURLinitFlag.CURL_GLOBAL_ALL)

            Dim mf As MultiPartForm
            mf = New MultiPartForm()

            For i As Integer = 1 To _post_fields_claves.Count

                mf.AddSection(CURLformoption.CURLFORM_COPYNAME, _post_fields_claves.Item(i), _
                   CURLformoption.CURLFORM_COPYCONTENTS, _post_fields_valores.Item(i), _
                   CURLformoption.CURLFORM_END)
            Next i

            For i As Integer = 1 To _post_field_file_claves.Count

                mf.AddSection(CURLformoption.CURLFORM_COPYNAME, _post_field_file_claves.Item(i), _
                    CURLformoption.CURLFORM_FILE, _post_field_file_valores.Item(i), _
                    CURLformoption.CURLFORM_CONTENTTYPE, "application/binary", _
                    CURLformoption.CURLFORM_END)
            Next i

            Dim easy As Easy
            easy = New Easy

            Dim wf As Easy.WriteFunction
            wf = New Easy.WriteFunction(AddressOf _get_respuesta)
            easy.SetOpt(CURLoption.CURLOPT_WRITEFUNCTION, wf)

            Dim pf As Easy.ProgressFunction
            pf = New Easy.ProgressFunction(AddressOf _get_progress)
            easy.SetOpt(CURLoption.CURLOPT_PROGRESSFUNCTION, pf)

            easy.SetOpt(CURLoption.CURLOPT_URL, url)
            easy.SetOpt(CURLoption.CURLOPT_HTTPPOST, mf)

            easy.Perform()

            easy.GetInfo(CURLINFO.CURLINFO_CONNECT_TIME, _infcurl.CURLINFO_CONNECT_TIME)
            easy.GetInfo(CURLINFO.CURLINFO_CONTENT_LENGTH_DOWNLOAD, _infcurl.CURLINFO_CONTENT_LENGTH_DOWNLOAD)
            easy.GetInfo(CURLINFO.CURLINFO_CONTENT_LENGTH_UPLOAD, _infcurl.CURLINFO_CONTENT_LENGTH_UPLOAD)
            easy.GetInfo(CURLINFO.CURLINFO_CONTENT_TYPE, _infcurl.CURLINFO_CONTENT_TYPE)
            easy.GetInfo(CURLINFO.CURLINFO_FILETIME, _infcurl.CURLINFO_FILETIME)
            easy.GetInfo(CURLINFO.CURLINFO_HEADER_SIZE, _infcurl.CURLINFO_HEADER_SIZE)
            easy.GetInfo(CURLINFO.CURLINFO_HTTPAUTH_AVAIL, _infcurl.CURLINFO_HTTPAUTH_AVAIL)
            easy.GetInfo(CURLINFO.CURLINFO_HTTP_CONNECTCODE, _infcurl.CURLINFO_HTTP_CONNECTCODE)
            easy.GetInfo(CURLINFO.CURLINFO_NAMELOOKUP_TIME, _infcurl.CURLINFO_NAMELOOKUP_TIME)
            easy.GetInfo(CURLINFO.CURLINFO_OS_ERRNO, _infcurl.CURLINFO_OS_ERRNO)
            easy.GetInfo(CURLINFO.CURLINFO_PRETRANSFER_TIME, _infcurl.CURLINFO_PRETRANSFER_TIME)
            easy.GetInfo(CURLINFO.CURLINFO_PRIVATE, _infcurl.CURLINFO_PRIVATE)
            easy.GetInfo(CURLINFO.CURLINFO_PROXYAUTH_AVAIL, _infcurl.CURLINFO_PROXYAUTH_AVAIL)
            easy.GetInfo(CURLINFO.CURLINFO_REDIRECT_COUNT, _infcurl.CURLINFO_REDIRECT_COUNT)
            easy.GetInfo(CURLINFO.CURLINFO_REDIRECT_TIME, _infcurl.CURLINFO_REDIRECT_TIME)
            easy.GetInfo(CURLINFO.CURLINFO_REQUEST_SIZE, _infcurl.CURLINFO_REQUEST_SIZE)
            easy.GetInfo(CURLINFO.CURLINFO_RESPONSE_CODE, _infcurl.CURLINFO_RESPONSE_CODE)
            easy.GetInfo(CURLINFO.CURLINFO_SIZE_DOWNLOAD, _infcurl.CURLINFO_SIZE_DOWNLOAD)
            easy.GetInfo(CURLINFO.CURLINFO_SIZE_UPLOAD, _infcurl.CURLINFO_SIZE_UPLOAD)
            easy.GetInfo(CURLINFO.CURLINFO_SPEED_DOWNLOAD, _infcurl.CURLINFO_SPEED_DOWNLOAD)
            easy.GetInfo(CURLINFO.CURLINFO_SPEED_UPLOAD, _infcurl.CURLINFO_SPEED_UPLOAD)
            easy.GetInfo(CURLINFO.CURLINFO_SSL_VERIFYRESULT, _infcurl.CURLINFO_SSL_VERIFYRESULT)
            easy.GetInfo(CURLINFO.CURLINFO_STARTTRANSFER_TIME, _infcurl.CURLINFO_STARTTRANSFER_TIME)
            easy.GetInfo(CURLINFO.CURLINFO_TOTAL_TIME, _infcurl.CURLINFO_TOTAL_TIME)

            easy.Cleanup()
            mf.Free()
            Curl.GlobalCleanup()

            _clear_fields()

            Return _respuesta

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Private Function _get_respuesta(ByVal buf() As Byte, ByVal size As Integer, ByVal nmemb As Integer, ByVal extraData As Object) As Integer
        _respuesta += System.Text.Encoding.UTF8.GetString(buf)
        Return size * nmemb
    End Function

    Public Function _get_progress(ByVal extraData As Object, ByVal dlTotal As Double, ByVal dlNow As Double, ByVal ulTotal As Double, ByVal ulNow As Double) As Int32
        RaiseEvent Progress(dlTotal, dlNow, ulTotal, ulNow)
        Return 0
    End Function

End Class