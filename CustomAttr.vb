Imports System
Imports System.Reflection

Namespace OData

    Public Module CustomAttr
        ' Visual Basic requires the AttributeUsage be specified.

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class DisplayName
            Inherits Attribute

            ' Keep a variable internally ...
            Private _DisplayName As String

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal DisplayName As String)
                _DisplayName = DisplayName
            End Sub

            ' .. and show a copy to the outside world.
            Public Property DisplayName() As String
                Get
                    Return _DisplayName
                End Get
                Set(ByVal Value As String)
                    _DisplayName = Value
                End Set
            End Property

        End Class

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class Pos
            Inherits Attribute

            ' Keep a variable internally ...
            Private _Pos As Integer = 0

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal Pos As Integer)
                _Pos = Pos
            End Sub

            ' .. and show a copy to the outside world.
            Public Property Pos() As Integer
                Get
                    Return _Pos
                End Get
                Set(ByVal Value As Integer)
                    _Pos = Value
                End Set
            End Property

        End Class

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class Tab
            Inherits Attribute

            ' Keep a variable internally ...
            Private _Tab As String

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal Tab As String)
                _Tab = Tab
            End Sub

            ' .. and show a copy to the outside world.
            Public Property Tab() As String
                Get
                    Return _Tab
                End Get
                Set(ByVal Value As String)
                    _Tab = Value
                End Set
            End Property

        End Class

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class Mandatory
            Inherits Attribute

            ' Keep a variable internally ...
            Private _isMandatory As Boolean = True

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal isMandatory As Boolean)
                _isMandatory = isMandatory
            End Sub

            ' .. and show a copy to the outside world.
            Public Property isMandatory() As Boolean
                Get
                    Return _isMandatory
                End Get
                Set(ByVal Value As Boolean)
                    _isMandatory = Value
                End Set
            End Property

        End Class

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class Browsable
            Inherits Attribute

            ' Keep a variable internally ...
            Private _Browsable As Boolean = True

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal Browsable As Boolean)
                _Browsable = Browsable
            End Sub

            ' .. and show a copy to the outside world.
            Public Property Browsable() As Boolean
                Get
                    Return _Browsable
                End Get
                Set(ByVal Value As Boolean)
                    _Browsable = Value
                End Set
            End Property

        End Class

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class [ReadOnly]
            Inherits Attribute

            ' Keep a variable internally ...
            Private _isReadOnly As Boolean

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal [ReadOnly] As Boolean)
                _isReadOnly = [ReadOnly]
            End Sub

            ' .. and show a copy to the outside world.
            Public Property [ReadOnly]() As Boolean
                Get
                    Return _isReadOnly
                End Get
                Set(ByVal Value As Boolean)
                    _isReadOnly = Value
                End Set
            End Property

        End Class

    End Module

    Public Class CustomProperties : Inherits PropertyInfo
        Private _me As PropertyInfo

        Sub New(ByRef Prop As PropertyInfo)
            _me = Prop
            For Each attr As Attribute In Attribute.GetCustomAttributes(Prop)
                If TypeOf attr Is [ReadOnly] Then
                    _readonly = CType(attr, [ReadOnly]).ReadOnly
                End If
                If TypeOf attr Is Mandatory Then
                    _Mandatory = CType(attr, Mandatory).isMandatory
                End If
                If TypeOf attr Is Browsable Then
                    _Browsable = CType(attr, Browsable).Browsable
                End If
                If TypeOf attr Is DisplayName Then
                    _DisplayName = CType(attr, DisplayName).DisplayName
                End If
                If TypeOf attr Is Tab Then
                    _Tab = CType(attr, Tab).Tab
                End If
                If TypeOf attr Is Pos Then
                    _pos = CType(attr, Pos).Pos
                End If
            Next
        End Sub

        Private _readonly As Boolean = False
        Public ReadOnly Property [Readonly]() As Boolean
            Get
                Return _readonly
            End Get
        End Property

        Private _Mandatory As Boolean = False
        Public ReadOnly Property Mandatory() As Boolean
            Get
                Return _Mandatory
            End Get
        End Property

        Private _Browsable As Boolean = True
        Public ReadOnly Property Browsable() As Boolean
            Get
                Return _Browsable
            End Get
        End Property

        Private _DisplayName As String
        Public ReadOnly Property DisplayName() As String
            Get
                Return _DisplayName
            End Get
        End Property

        Private _Tab As String = Nothing
        Public ReadOnly Property Tab() As String
            Get
                Return _Tab
            End Get
        End Property

        Private _pos As Integer = 0
        Public ReadOnly Property POS() As Integer
            Get
                Return _pos
            End Get
        End Property        

        Public Overrides ReadOnly Property Attributes() As System.Reflection.PropertyAttributes
            Get
                Return _me.Attributes
            End Get
        End Property

        Public Overrides ReadOnly Property CanRead() As Boolean
            Get
                Return _me.CanRead
            End Get
        End Property

        Public Overrides ReadOnly Property CanWrite() As Boolean
            Get
                Return _me.CanWrite
            End Get
        End Property

        Public Overrides ReadOnly Property DeclaringType() As System.Type
            Get
                Return _me.DeclaringType
            End Get
        End Property

        Public Overloads Overrides Function GetAccessors(ByVal nonPublic As Boolean) As System.Reflection.MethodInfo()
            Return _me.GetAccessors(nonPublic)
        End Function

        Public Overloads Overrides Function GetCustomAttributes(ByVal inherit As Boolean) As Object()
            Return _me.GetCustomAttributes(inherit)
        End Function

        Public Overloads Overrides Function GetCustomAttributes(ByVal attributeType As System.Type, ByVal inherit As Boolean) As Object()
            Return _me.GetCustomAttributes(attributeType, inherit)
        End Function

        Public Overloads Overrides Function GetGetMethod(ByVal nonPublic As Boolean) As System.Reflection.MethodInfo
            Return _me.GetGetMethod(nonPublic)
        End Function

        Public Overrides Function GetIndexParameters() As System.Reflection.ParameterInfo()
            Return _me.GetIndexParameters
        End Function

        Public Overloads Overrides Function GetSetMethod(ByVal nonPublic As Boolean) As System.Reflection.MethodInfo
            Return _me.GetSetMethod(nonPublic)
        End Function

        Public Overloads Overrides Function GetValue(ByVal obj As Object, ByVal invokeAttr As System.Reflection.BindingFlags, ByVal binder As System.Reflection.Binder, ByVal index() As Object, ByVal culture As System.Globalization.CultureInfo) As Object
            Return _me.GetValue(obj, invokeAttr, binder, index, culture)
        End Function

        Public Overrides Function IsDefined(ByVal attributeType As System.Type, ByVal inherit As Boolean) As Boolean
            Return _me.IsDefined(attributeType, inherit)
        End Function

        Public Overrides ReadOnly Property Name() As String
            Get
                Return _me.Name

            End Get
        End Property

        Public Overrides ReadOnly Property PropertyType() As System.Type
            Get
                Return _me.PropertyType
            End Get
        End Property

        Public Overrides ReadOnly Property ReflectedType() As System.Type
            Get
                Return _me.ReflectedType
            End Get
        End Property

        Public Overloads Overrides Sub SetValue(ByVal obj As Object, ByVal value As Object, ByVal invokeAttr As System.Reflection.BindingFlags, ByVal binder As System.Reflection.Binder, ByVal index() As Object, ByVal culture As System.Globalization.CultureInfo)
            _me.SetValue(obj, value, invokeAttr, binder, index, culture)
        End Sub
    End Class

End Namespace