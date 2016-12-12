Imports System.IO
Imports System.Net
Imports System.Xml

Public Class ResponseEventArgs
    Inherits System.EventArgs

    Private _MemoryStream As New MemoryStream

    Private _WebException As WebException
    Public ReadOnly Property WebException As WebException
        Get
            Return _WebException
        End Get
    End Property

    Private _HttpWebResponse As HttpWebResponse
    Public ReadOnly Property HttpWebResponse As HttpWebResponse
        Get
            Return _HttpWebResponse
        End Get
    End Property

    Public ReadOnly Property StreamReader As StreamReader
        Get
            Dim ret As New StreamReader(_MemoryStream)
            ret.BaseStream.Position = 0
            Return ret
        End Get
    End Property

    Public ReadOnly Property InterfaceException As Exception
        Get
            Dim doc As New XmlDocument
            Try
                Dim xdoc As XDocument = XDocument.Parse(StreamReader.ReadToEnd)
                Return New Exception( _
                    xdoc.Descendants( _
                        "FORM" _
                    ).FirstOrDefault.Descendants( _
                        "InterfaceErrors" _
                    ).FirstOrDefault.Descendants( _
                        "text" _
                    ).FirstOrDefault.Value _
                )

            Catch ex As Exception
                Return ex
            End Try
        End Get
    End Property

    Public Sub New(ByRef HttpWebResponse As HttpWebResponse, Optional WebException As WebException = Nothing)

        _WebException = WebException
        _HttpWebResponse = HttpWebResponse

        Dim buffer() As Byte
        Dim BytesRead As Integer

        Using HttpWebResponse.GetResponseStream            
            Do
                ReDim buffer(1024)
                BytesRead = HttpWebResponse.GetResponseStream.Read(buffer, 0, buffer.Length)
                _MemoryStream.Write(buffer, 0, BytesRead)
            Loop While BytesRead > 0
        End Using
        HttpWebResponse.GetResponseStream.Close()

        Dim sr As New StreamReader(_MemoryStream)
        sr.BaseStream.Position = 0

        If OData.Connection.Settings.Debug.ShowOutput Then
            OData.Connection.RaiseDebug( _
                "<--{0}: {1}{2}{3}", _
                CInt(_HttpWebResponse.StatusCode), _
                _HttpWebResponse.StatusDescription, _
                vbCrLf, _
                sr.ReadToEnd _
            )
        End If

        _MemoryStream.Position = 0


    End Sub

    Public Sub toMsg()
        With _HttpWebResponse
            Dim ret As New StreamReader(.GetResponseStream)
            OData.Connection.RaiseError( _
                String.Format( _
                    "{0}: {1}", _
                    CInt(.StatusCode), _
                    .StatusDescription _
                ), _
                ret.ReadToEnd _
            )

        End With
    End Sub

End Class

Public Delegate Sub ResponseEventHandler(sender As Object, e As ResponseEventArgs)

