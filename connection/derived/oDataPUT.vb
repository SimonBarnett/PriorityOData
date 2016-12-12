Imports System.IO

Namespace OData

    Friend Class oDataPUT : Inherits oDataConnection

#Region "Constructor"

        Private this As MemoryStream
        Public Sub New(ByRef Obj As oDataObject, ByRef ms As MemoryStream, ByRef Handler As ResponseEventHandler)
            Responsehandler = Handler
            this = ms
            SharedConstructor(String.Format("{0}({1})", Obj.url, Obj.KeyString))

        End Sub

#End Region

#Region "Overriden properties"

        Public Overrides ReadOnly Property Method As String
            Get
                Return "PATCH"
            End Get
        End Property

        Public Overrides ReadOnly Property RequestStream As MemoryStream
            Get
                Return this
            End Get
        End Property

#End Region

    End Class

End Namespace