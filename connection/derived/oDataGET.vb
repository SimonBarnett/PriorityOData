Imports System.IO
Imports System.Net

Namespace OData

    Friend Class oDataGET : Inherits oDataConnection

#Region "Constructor"

        Private thisObject As oDataQuery
        Public Sub New(ByRef Obj As oDataQuery, ByRef Handler As ResponseEventHandler, Optional Filter As String = Nothing, Optional Sort As String = Nothing, Optional Direction As String = Nothing)

            Responsehandler = Handler
            thisObject = Obj

            If Filter Is Nothing And Sort Is Nothing Then
                SharedConstructor(Obj.url)

            ElseIf Not Filter Is Nothing And Sort Is Nothing Then
                SharedConstructor( _
                    String.Format( _
                        "{0}?$filter={1}", _
                        Obj.url, _
                        Filter _
                    ) _
                )

            ElseIf Filter Is Nothing And Not Sort Is Nothing Then
                SharedConstructor( _
                    String.Format( _
                        "{0}?orderby={1} {2}", _
                        Obj.url, _
                        Sort, _
                        Direction.ToString _
                    ) _
                )

            Else
                SharedConstructor( _
                    String.Format( _
                        "{0}?$filter={1}&orderby={2} {3}", _
                        Obj.url, _
                        Filter, _
                        Sort, _
                        Direction.ToString _
                    ) _
                )

            End If

        End Sub

#End Region

#Region "Overriden properties"

        Public Overrides ReadOnly Property Method As String
            Get
                Return "GET"
            End Get
        End Property

#End Region

    End Class

End Namespace