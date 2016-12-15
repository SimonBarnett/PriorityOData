Imports System.Net

Namespace OData

    Public Module Connection

        Public Settings As New Settings

        Private _LastError As Exception
        Public Property LastError() As Exception
            Get
                Return _LastError
            End Get
            Set(ByVal value As Exception)
                _LastError = value
            End Set
        End Property


#Region "Events"

#Region "Debug Events"

        Public Event DebugOutput(ByVal data As String)
        Friend Sub RaiseDebug(ByVal data As String)
            RaiseEvent DebugOutput(data)

        End Sub
        Friend Sub RaiseDebug(ByVal data As String, ByVal ParamArray args() As String)
            RaiseEvent DebugOutput(String.Format(data, args))

        End Sub

        Public Event ErrorOutput(ByVal Title As String, ByVal data As String)
        Friend Sub RaiseError(ByVal Title As String, ByVal data As String)
            RaiseEvent ErrorOutput(Title, data)

        End Sub
        Friend Sub RaiseError(ByVal Title As String, ByVal data As String, ByVal ParamArray args() As String)
            RaiseEvent ErrorOutput(Title, String.Format(data, args))

        End Sub

#End Region

#Region "Connection Events"

        Public Event StartData()
        Friend Sub RaiseStartData()
            RaiseEvent StartData()

        End Sub

        Public Event EndData()
        Public Sub RaiseEndData()
            RaiseEvent EndData()

        End Sub

#End Region

#End Region

    End Module

    Public Class Settings

        Public Credential As NetworkCredential
        Public Proxy As IWebProxy = Nothing

        Private _CheckSSL As Boolean = True
        Public Property CheckSSL() As Boolean
            Get
                Return _CheckSSL
            End Get
            Set(ByVal value As Boolean)
                If Not value Then
                    System.Net.ServicePointManager.CertificatePolicy = New QueryTrustCertificatePolicy
                End If
            End Set
        End Property

        Public Debug As New Debug

        Private _Server As String
        Public Property Server() As String
            Get
                Return _Server
            End Get
            Set(ByVal value As String)
                _Server = value
            End Set
        End Property

        Private _Environment As String
        Public Property Environment() As String
            Get
                Return _Environment
            End Get
            Set(ByVal value As String)
                _Environment = value
            End Set
        End Property

    End Class

    Public Class Debug

        Public Sub ShowAll()
            _ShowMethod = True
            _ShowHeaders = True
            _ShowInput = True
            _ShowOutput = True
        End Sub

        Public Sub ShowNone()
            _ShowMethod = False
            _ShowHeaders = False
            _ShowInput = False
            _ShowOutput = False
        End Sub

        Private _ShowMethod As Boolean = False
        Public Property ShowMethod As Boolean
            Get
                Return _ShowMethod
            End Get
            Set(value As Boolean)
                _ShowMethod = value
            End Set
        End Property

        Private _ShowHeaders As Boolean = False
        Public Property ShowHeaders As Boolean
            Get
                Return _ShowHeaders
            End Get
            Set(value As Boolean)
                _ShowHeaders = value
            End Set
        End Property

        Private _ShowInput As Boolean = False
        Public Property ShowInput As Boolean
            Get
                Return _ShowInput
            End Get
            Set(value As Boolean)
                _ShowInput = value
            End Set
        End Property

        Private _ShowOutput As Boolean = False
        Public Property ShowOutput As Boolean
            Get
                Return _ShowOutput
            End Get
            Set(value As Boolean)
                _ShowOutput = value
            End Set
        End Property

    End Class

End Namespace

