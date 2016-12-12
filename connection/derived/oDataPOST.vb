Imports System.IO
Imports System.Net

Namespace OData

    Friend Class oDataPOST : Inherits oDataConnection

#Region "Constructor"

        Private _thisObject As oDataObject
        Public Property thisObject() As oDataObject
            Get
                Return _thisObject
            End Get
            Set(ByVal value As oDataObject)
                _thisObject = value
            End Set
        End Property

        Public Sub New(ByRef Obj As oDataObject, ByRef Handler As ResponseEventHandler)
            Responsehandler = Handler
            _thisObject = Obj
            SharedConstructor(Obj.url)

        End Sub

#End Region

#Region "Overriden properties"

        Public Overrides ReadOnly Property Method As String
            Get
                Return "POST"
            End Get
        End Property

        Public Overrides ReadOnly Property RequestStream As MemoryStream
            Get
                Return thisObject.toStream
            End Get
        End Property

#End Region

    End Class

End Namespace