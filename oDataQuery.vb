Imports System.IO
Imports System.Net
Imports Newtonsoft.Json

Namespace OData

    Public Enum SortDirection
        asc
        desc
    End Enum

    Public MustInherit Class oDataQuery

#Region "Private Properties"

        <JsonIgnore()> _
        Friend Adding As Boolean = False

        <JsonIgnore()> _
        Private Deleting As Boolean = False

        <JsonIgnore()> _
        Private _LastResult As Exception

#End Region

#Region "Friend Properties"

        <JsonIgnore()> _
        Public MustOverride Property Parent() As oDataObject

        <JsonIgnore()> _
        Protected Friend MustOverride ReadOnly Property EntityName() As String

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

#Region "Friend Methods"

        Protected Friend MustOverride Sub Deserialise(ByRef stream As StreamReader)

        Protected Friend MustOverride Sub Remove(ByRef obj As oDataObject)

        Protected Friend MustOverride Sub HandlesAdd(ByVal sender As Object, ByVal e As ResponseEventArgs)

        Public MustOverride ReadOnly Property ObjectType As Type

        Public MustOverride ReadOnly Property BindingSource() As System.Windows.Forms.BindingSource

#End Region

#Region "Public Methods"

        ''' <summary>
        ''' Load the query without a filter ot sort.
        ''' </summary>
        Public Sub Load()
            Dim cn As New oDataGET(Me, AddressOf HandlesGet)
            OData.Connection.RaiseEndData()

        End Sub

        ''' <summary>
        ''' Load the query with a filter.
        ''' </summary>
        ''' <param name="Filter">A string containing the filter expression.</param>
        Public Sub Load(Filter As String)
            Dim cn As New oDataGET(Me, AddressOf HandlesGet, Filter)
            OData.Connection.RaiseEndData()

        End Sub

        ''' <summary>
        ''' Load the query sorted by the SortField.
        ''' </summary>
        ''' <param name="SortField">The name of the field to sort by.</param>
        ''' <param name="Direction">The sort direction.</param>
        Public Sub Load(SortField As String, Direction As SortDirection)
            Dim cn As New oDataGET(Me, AddressOf HandlesGet, Nothing, SortField, Direction.ToString)
            OData.Connection.RaiseEndData()

        End Sub

        ''' <summary>
        ''' Load the query with both a filter and a sort.
        ''' </summary>
        ''' <param name="Filter">A string containing the filter expression.</param>
        ''' <param name="SortField">The name of the field to sort by.</param>
        ''' <param name="Direction">The sort direction.</param>
        Public Sub Load(Filter As String, SortField As String, Direction As SortDirection)
            Dim cn As New oDataGET(Me, AddressOf HandlesGet, Filter, SortField, Direction.ToString)
            OData.Connection.RaiseEndData()

        End Sub

        ''' <summary>
        ''' Add an oDataObject to the oDataQueryArray
        ''' </summary>
        ''' <param name="obj">The oDataObject to add</param>
        ''' <remarks></remarks>
        Public Sub Add(ByRef obj As oDataObject)
            Adding = True
            Dim cn As New oDataPOST(obj, AddressOf Me.HandlesAdd)
            obj = cn.thisObject
            OData.Connection.RaiseEndData()

        End Sub

        ''' <summary>
        ''' Delete an object from the array.
        ''' Throws an exception containing the Interface errors if fails.
        ''' </summary>
        ''' <param name="obj">The oDataObject to delete from the array.</param>
        Public Sub Delete(obj As oDataObject)
            Deleting = True
            Dim cn As New oDataDELETE(obj, AddressOf HandlesDelete)
            OData.Connection.RaiseEndData()
            If Not _LastResult Is Nothing Then
                Throw (_LastResult)
            End If

        End Sub

        ''' <summary>
        ''' A dictionary containing the titles of the queries child forms.
        ''' </summary>
        ''' <remarks></remarks>
        <JsonIgnore()> _
        Public ChildQuery As New Dictionary(Of Integer, String)

        ''' <summary>
        ''' A dictionary containing references to this forms siblig forms.
        ''' </summary>
        ''' <remarks></remarks>
        <JsonIgnore()> _
        Public SibligQuery As New Dictionary(Of Integer, oNavigation)

#End Region

#Region "Response Handlers"

        Friend Sub HandlesGet(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException Is Nothing Then
                e.toMsg()
            Else
                Deserialise(e.StreamReader)
            End If
        End Sub

        Friend Sub HandlesDelete(ByVal sender As Object, ByVal e As ResponseEventArgs)
            If Not e.WebException Is Nothing Then
                _LastResult = e.InterfaceException

            Else
                _LastResult = Nothing
                Remove(TryCast(sender, oDataDELETE).thisObject)

            End If
            Deleting = False
        End Sub

#End Region

    End Class

End Namespace
