Imports System.IO
Imports System.Net

Namespace OData

    Friend Class oDataDELETE : Inherits oDataConnection

#Region "Constructor"

        Public thisObject As oDataObject
        Public Sub New(ByRef Obj As oDataObject, ByRef Handler As ResponseEventHandler)
            Responsehandler = Handler
            thisObject = Obj
            SharedConstructor(String.Format("{0}({1})", Obj.url, Obj.KeyString))

        End Sub

#End Region

#Region "Overriden properties"

        Public Overrides ReadOnly Property Method As String
            Get
                Return "DELETE"
            End Get
        End Property

#End Region

    End Class

End Namespace