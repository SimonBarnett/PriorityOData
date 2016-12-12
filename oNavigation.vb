Imports Priority.OData

Namespace OData

    Public Class oNavigation

        Public Sub New(ByVal Name As String, ByRef oDataQuery As oDataQuery)
            _Name = Name
            _oDataQuery = oDataQuery
        End Sub

        Public Sub New(ByVal Name As String)
            _Name = Name        
        End Sub

        Private _Name As String
        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                _Name = value
            End Set
        End Property

        Private _oDataQuery As oDataQuery
        Public ReadOnly Property oDataQuery() As oDataQuery
            Get
                Return _oDataQuery
            End Get
        End Property

        Public Sub SetoDataQuery(ByRef oDataQuery As oDataQuery)
            _oDataQuery = oDataQuery
        End Sub

    End Class

End Namespace