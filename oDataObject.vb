Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Namespace OData

    Public MustInherit Class oDataObject : Implements IDisposable

#Region "regex"

        'Dim numeric As New Regex("^[0-9\.\-]+$")
        'Dim INT64 As New Regex("^[0-9\-]+$")

        Public Function Validate(ByVal Name As String, ByVal Value As Object, ByVal Regex As String) As Boolean
            Dim strReg As New Regex(Regex)
            If Not strReg.Match(Value).Success Then
                Connection.LastError = New Exception( _
                    String.Format("Invalid data for {0}.", Name) _
                )
                Return False
            Else
                Return True

            End If

        End Function

        'Public Function Validate(ByVal Name As String, ByVal Value As Object, ByVal Precision As Integer, ByVal Scale As Integer) As Boolean
        '    If Not numeric.Match(Value).Success Then
        '        Connection.LastError = New Exception( _
        '            String.Format("{0} must be a number.", Name) _
        '        )
        '        Return False
        '    Else
        '        Return True
        '    End If
        'End Function

        'Public Function Validate(ByVal Name As String, ByVal Value As Object, ByVal isDate As Boolean) As Boolean
        '    Return True
        'End Function

        'Public Function Validate(ByVal Name As String, ByVal Value As Object) As Boolean
        '    If Not INT64.Match(Value).Success Then
        '        Connection.LastError = New Exception( _
        '            String.Format("{0} must be an integer value.", Name) _
        '        )
        '        Return False
        '    Else
        '        Return True
        '    End If
        'End Function

#End Region

#Region "Friend Properties"

        <JsonIgnore()> _
        Friend Loading As Boolean = True

        <JsonIgnore()> _
        Protected Friend ReadOnly Property url() As String
            Get
                Select Case Parent Is Nothing
                    Case True
                        Return String.Format( _
                            "https://{0}/odata/Priority/tabula.ini/{1}/{2}", _
                            Connection.Settings.Server, _
                            Connection.Settings.Environment, _
                        Me.EntityName _
                        )
                    Case Else
                        Return String.Format( _
                        "{0}({1})/{2}", _
                        Parent.url, _
                        Parent.KeyString, _
                        Me.EntityName _
                    )

                End Select
            End Get
        End Property

#End Region

#Region "Must Override Methods"

        Protected Friend MustOverride Sub toJSON(ByRef jw As JsonTextWriter)

        Protected Friend MustOverride Sub toXML(ByRef xw As XmlWriter, ByVal name As String)

        Protected Friend MustOverride Sub HandlesEdit(ByVal sender As Object, ByVal e As ResponseEventArgs)

#End Region

#Region "Must Override Properties"

        <JsonIgnore(), Browsable(False)> _
        Protected Friend MustOverride ReadOnly Property EntityName() As String

        <JsonIgnore(), Browsable(False)> _
        Public MustOverride ReadOnly Property KeyString() As String

        <JsonIgnore(), Browsable(False)> _
        Public MustOverride Property Parent() As oDataObject

#End Region

#Region "Friend Methods"

        Friend Function PropertyStream(ByVal Name As String, ByVal Value As Object) As MemoryStream
            Dim ms As New MemoryStream
            Dim patch As New JsonTextWriter(New StreamWriter(ms))
            With patch
                .WriteRaw("{ ")
                .WriteRaw(String.Format("""{0}"": ", Name))
                .WriteValue(Value)
                .WriteRaw(" }")
                .Flush()
            End With
            Return ms
        End Function

        Friend Function toStream() As MemoryStream
            Dim ms As New MemoryStream
            Dim sw As TextWriter = New StreamWriter(ms)
            Dim jw As New Newtonsoft.Json.JsonTextWriter(sw)
            With jw
                .WriteRaw("{ ")
                toJSON(jw)
                .WriteRaw(" }")
                .Flush()
            End With
            Return ms
        End Function

        Friend Sub HandlesPost(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException Is Nothing Then
                e.toMsg()
            Else
                Dim resp As New StreamReader(e.HttpWebResponse.GetResponseStream)
                Connection.RaiseDebug(resp.ReadToEnd)
            End If
        End Sub

#End Region

#Region "Public Methods"

        ''' <summary>
        ''' Post this oDataObject to the server, including any sub queries.
        ''' </summary>
        Public Sub Post()
            Dim cn As New oDataPOST(Me, AddressOf HandlesPost)
        End Sub

        ''' <summary>
        ''' Write an XML representation of the oDataObject
        ''' </summary>
        ''' <param name="xw">An XML stream to write to.</param>
        Public Sub writeXML(ByRef xw As XmlWriter)
            Me.toXML(xw, Nothing)
        End Sub

        <JsonIgnore()> _
        Public ChildQuery As New Dictionary(Of Integer, oNavigation)

#End Region

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

End Namespace