Imports System.IO
Imports System.Net
Imports System.Threading

Namespace OData

    Friend MustInherit Class oDataConnection
        Public Shared allDone As New ManualResetEvent(False)

#Region "Overrides"

        MustOverride ReadOnly Property Method As String

        Overridable ReadOnly Property RequestStream As MemoryStream
            Get
                Return Nothing
            End Get
        End Property

#End Region

#Region "Properties"

        Private _uploadRequest As Net.HttpWebRequest
        Public ReadOnly Property uploadRequest As HttpWebRequest
            Get
                Return _uploadRequest
            End Get
        End Property

        Private _Responsehandler As ResponseEventHandler
        Public Property Responsehandler As ResponseEventHandler
            Get
                Return _Responsehandler
            End Get
            Set(value As ResponseEventHandler)
                _Responsehandler = value
            End Set
        End Property

#End Region

#Region "Constructor"

        Public Sub SharedConstructor(ByRef url As String)

            Connection.LastError = Nothing
            _uploadRequest = CType(HttpWebRequest.Create(url), HttpWebRequest)
            With _uploadRequest
                .UserAgent = "Medatech REST Client 1.0"
                .Method = Method
                .ContentType = "application/json;odata.metadata=minimal"
                .Accept = "application/json"
                .Headers.Add("OData-Version: 4.0")
                .Credentials = Connection.Settings.Credential
                .Proxy = Connection.Settings.Proxy
                .AllowWriteStreamBuffering = True

                If Connection.Settings.Debug.ShowMethod Then
                    Connection.RaiseDebug( _
                    "-->{0}: {1}", _
                    .Method.ToString, _
                    .RequestUri.AbsoluteUri _
                )
                End If

                If Connection.Settings.Debug.ShowHeaders Then
                    For Each k As String In .Headers
                        Connection.RaiseDebug( _
                            "{0}: {1}", _
                            k, .Headers(k) _
                        )
                    Next
                End If

            End With

            Dim result As IAsyncResult
            If RequestStream Is Nothing Then
                result = _
                    CType(_uploadRequest.BeginGetResponse(AddressOf GetResponseCallback, _uploadRequest),  _
                    IAsyncResult _
                )
            Else
                result = _
                    CType(_uploadRequest.BeginGetRequestStream(AddressOf GetRequestStreamCallback, _uploadRequest),  _
                    IAsyncResult _
                )
            End If

            With allDone
                OData.Connection.RaiseStartData()
                .WaitOne()
                .Reset()
            End With

        End Sub

        'Private Function hssl(sender As Object,
        '    certificate As System.Security.Cryptography.X509Certificates.X509Certificate,
        '    chain As System.Security.Cryptography.X509Certificates.X509Chain,
        '    sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean

        '    Return True

        'End Function

#End Region

#Region "Async Request"

        Private Sub GetRequestStreamCallback(ByVal asynchronousResult As IAsyncResult)

            Dim request As HttpWebRequest = CType(asynchronousResult.AsyncState, HttpWebRequest)

            ' End the operation
            Using postStream As Stream = request.EndGetRequestStream(asynchronousResult)

                ' Upload
                Dim buffer(1024) As Byte
                Dim bytesRead As Integer
                Dim ms As MemoryStream = RequestStream

                With ms

                    If OData.Connection.Settings.Debug.ShowInput Then
                        Dim sr As New StreamReader(ms)
                        .Position = 0
                        Connection.RaiseDebug(sr.ReadToEnd)
                        .Position = 0
                    End If

                    While True
                        bytesRead = .Read(buffer, 0, buffer.Length)
                        If bytesRead = 0 Then
                            Exit While
                        End If
                        postStream.Write(buffer, 0, bytesRead)

                    End While

                End With

            End Using

            ' Start the asynchronous operation to get the response
            Dim result As IAsyncResult = _
                CType(request.BeginGetResponse(AddressOf GetResponseCallback, request),  _
                IAsyncResult _
            )

        End Sub ' ReadRequestStreamCallback

        Private Sub GetResponseCallback(ByVal asynchronousResult As IAsyncResult)

            Dim request As HttpWebRequest = CType(asynchronousResult.AsyncState, HttpWebRequest)
            Try
                '  Get the response.
                Dim response As HttpWebResponse = CType(request.EndGetResponse(asynchronousResult),  _
                   HttpWebResponse _
                )
                _Responsehandler.Invoke(Me, New ResponseEventArgs(response))

            Catch e As WebException
                If e.Response Is Nothing Then
                    Throw (e)
                Else
                    _Responsehandler.Invoke(Me, New ResponseEventArgs(e.Response, e))
                End If

            Catch e As Exception
                Throw (e)

            Finally                
                allDone.Set()

            End Try

        End Sub ' ReadResponseCallback

#End Region

    End Class

End Namespace