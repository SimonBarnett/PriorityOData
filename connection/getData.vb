﻿Imports System.Xml
Imports System.IO
Imports Priority.OData

Public Class GetData
    Implements IDisposable

    Public Event StartRead()
    Public Event EndRead()

#Region "private variables"

    Private Enum eRequestType
        Tabular = 1
        Scalar = 2
        NonQuery = 3
    End Enum

    Private disposedValue As Boolean = False        ' To detect redundant calls
    Private _SQL As String
    Private _RequestType As eRequestType

#End Region

#Region "initialisation and finalisation"

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

#Region "Connection Properties"

    Private ReadOnly Property URL() As String
        Get
            Return _
                String.Format( _
                    "https://{0}/api/sqlHandler.ashx", _
                    Connection.Settings.Server _
                )
        End Get
    End Property


    Private ReadOnly Property Environment() As String
        Get
            Return Connection.Settings.Environment
        End Get
    End Property

#End Region

#Region "Public Methods"

    Public Function ExecuteReader(ByVal SQL As String, Optional ByVal Retry As Boolean = True) As Data.DataTable

        _SQL = SQL
        _RequestType = eRequestType.Tabular

        Dim result As New Data.DataTable()
        Dim ex As Exception = Nothing
        Dim response As XmlDocument

        Do
            If Not IsNothing(ex) Then
                If Retry Then
                    If MsgBox(ex.Message, MsgBoxStyle.RetryCancel, "Connection error.") = MsgBoxResult.Cancel Then
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End If

            RaiseEvent StartRead()
            ex = Nothing
            response = Post(ex)
            RaiseEvent EndRead()

        Loop Until IsNothing(ex)

        If Not IsNothing(ex) Then
            Throw New Exception(ex.Message)

        Else
            RaiseEvent StartRead()
            Try
                For Each col As XmlNode In response.SelectNodes("//column")
                    result.Columns.Add( _
                        col.Attributes("name").InnerText, _
                        Type.GetType(col.Attributes("type").InnerText) _
                    )
                Next
                For Each row As XmlNode In response.SelectNodes("//row")
                    Dim values() As Object = Nothing
                    Dim i As Integer = 0
                    For Each col As System.Data.DataColumn In result.Columns
                        Try
                            ReDim Preserve values(UBound(values) + 1)
                        Catch
                            ReDim values(0)
                        Finally
                            values(UBound(values)) = row.Attributes(i).InnerText
                            i += 1
                        End Try
                    Next
                    result.Rows.Add(values)
                Next

            Catch exm As Exception
                RaiseEvent EndRead()
                Throw New Exception(exm.Message)

            Finally
                RaiseEvent EndRead()

            End Try

        End If

        Return result


    End Function

    Public Function ExecuteScalar(ByVal SQL As String, Optional ByVal Retry As Boolean = True) As Object

        _SQL = SQL
        _RequestType = eRequestType.Scalar

        Dim ex As Exception = Nothing
        Dim response As XmlDocument
        Dim result As Object

        Do
            If Not IsNothing(ex) Then
                If Retry Then
                    If MsgBox(ex.Message, MsgBoxStyle.RetryCancel, "Connection error.") = MsgBoxResult.Cancel Then
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End If

            RaiseEvent StartRead()
            ex = Nothing
            response = Post(ex)
            RaiseEvent EndRead()

        Loop Until IsNothing(ex)

        If Not IsNothing(ex) Then
            Throw New Exception(ex.Message)

        Else
            Try
                result = response.SelectSingleNode("//result").InnerText

            Catch exm As Exception
                RaiseEvent EndRead()
                Throw New Exception(exm.Message)

            Finally
                RaiseEvent EndRead()

            End Try

        End If

        Return result

    End Function

    Public Sub ExecuteNonQuery(ByVal SQL As String, Optional ByVal Retry As Boolean = True)

        RaiseEvent StartRead()

        _SQL = SQL
        _RequestType = eRequestType.NonQuery

        Dim ex As Exception = Nothing
        Dim response As XmlDocument

        Do
            If Not IsNothing(ex) Then
                If Retry Then
                    If MsgBox(ex.Message, MsgBoxStyle.RetryCancel, "Connection error.") = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            End If

            RaiseEvent StartRead()
            ex = Nothing
            response = Post(ex)
            RaiseEvent EndRead()

        Loop Until IsNothing(ex)

        If Not IsNothing(ex) Then
            Throw New Exception(ex.Message)

        End If

    End Sub

#End Region

#Region "Post request"

    Private Function Post(ByRef Ex As Exception) As XmlDocument

        Dim uploadResponse As Net.HttpWebResponse = Nothing
        Dim requestStream As Stream = Nothing
        Dim posted As Boolean = False
        Ex = Nothing

        Try

            Dim ms As MemoryStream = New MemoryStream(toByte)
            Dim uploadRequest As Net.HttpWebRequest = CType(Net.HttpWebRequest.Create(URL), Net.HttpWebRequest)

            uploadRequest.Method = "POST"
            uploadRequest.Proxy = Nothing
            uploadRequest.SendChunked = True
            requestStream = uploadRequest.GetRequestStream()

            ' Upload the XML
            Dim buffer(1024) As Byte
            Dim bytesRead As Integer
            While True
                bytesRead = ms.Read(buffer, 0, buffer.Length)
                If bytesRead = 0 Then
                    Exit While
                End If
                requestStream.Write(buffer, 0, bytesRead)
            End While

            ' The request stream must be closed before getting the response.
            requestStream.Close()
            uploadResponse = uploadRequest.GetResponse()

            Dim thisRequest As New XmlDocument
            Dim reader As New StreamReader(uploadResponse.GetResponseStream)
            thisRequest.LoadXml(reader.ReadToEnd)

            With thisRequest.SelectSingleNode("response")
                Select Case CInt(.Attributes("status").Value)
                    Case 200
                        Return thisRequest

                    Case 400
                        Throw New Exception(String.Format("Request error: {0}", .Attributes("message").Value))

                    Case 500
                        Throw New Exception(String.Format("Response error: {0}", .Attributes("message").Value))

                End Select

            End With

        Catch exep As UriFormatException
            Ex = New Exception(String.Format("Invalid URL: {0}", exep.Message))
        Catch exep As Net.WebException
            Ex = New Exception(String.Format("Connection Error: {0}", exep.Message))
        Catch exep As Exception
            Ex = New Exception(String.Format("Request failed: {0}", exep.Message))
        Finally
            ' Clean up the streams
            If Not IsNothing(uploadResponse) Then
                uploadResponse.Close()
            End If
            If Not IsNothing(requestStream) Then
                requestStream.Close()
            End If
        End Try

        Return Nothing

    End Function

    Private ReadOnly Property toByte() As Byte()
        Get
            Dim myEncoder As New System.Text.ASCIIEncoding
            Dim str As New System.Text.StringBuilder
            Dim xw As XmlWriter = XmlWriter.Create(str)
            With xw

                .WriteStartDocument()
                .WriteStartElement("GetRequest")
                .WriteElementString("Environment", Environment)

                Select Case _RequestType
                    Case eRequestType.Tabular
                        .WriteElementString("RequestType", "tabular")
                    Case eRequestType.Scalar
                        .WriteElementString("RequestType", "scalar")
                    Case eRequestType.NonQuery
                        .WriteElementString("RequestType", "nonquery")
                    Case Else

                End Select

                .WriteElementString("SQL", _SQL)

                .WriteEndDocument()
                .Flush()

            End With

            Return myEncoder.GetBytes(str.ToString)
        End Get
    End Property

#End Region

End Class
