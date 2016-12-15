Imports System
Imports System.Reflection

Namespace OData

    Public Module ClassAttr
        ' Visual Basic requires the AttributeUsage be specified.

        <AttributeUsage(AttributeTargets.Class)> _
        Public Class QueryTitle
            Inherits Attribute

            ' Keep a variable internally ...
            Private _QueryTitle As String

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal QueryTitle As String)
                _QueryTitle = QueryTitle
            End Sub

            ' .. and show a copy to the outside world.
            Public Property QueryTitle() As String
                Get
                    Return _QueryTitle
                End Get
                Set(ByVal Value As String)
                    _QueryTitle = Value
                End Set
            End Property

        End Class

    End Module

    Public Class ClassProperties

        Sub New(ByRef ClassInfo As Type)
            For Each attr As Attribute In Attribute.GetCustomAttributes(ClassInfo)
                If TypeOf attr Is QueryTitle Then
                    _QueryTitle = CType(attr, QueryTitle).QueryTitle
                End If
            Next
        End Sub

        Private _QueryTitle As String
        Public ReadOnly Property QueryTitle() As String
            Get
                Return _QueryTitle
            End Get
        End Property

    End Class

End Namespace