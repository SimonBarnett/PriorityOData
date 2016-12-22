Imports System
Imports System.Reflection

Namespace OData

#Region "Attribute Classes"

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

#Region "Property Type"

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class nType
            Inherits Attribute

            ' Keep a variable internally ...
            Private _nType As String

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal nType As String)
                _nType = nType
            End Sub

            ' .. and show a copy to the outside world.
            Public Property nType() As String
                Get
                    Return _nType
                End Get
                Set(ByVal Value As String)
                    _nType = Value
                End Set
            End Property

        End Class

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class Scale
            Inherits Attribute

            ' Keep a variable internally ...
            Private _Scale As Integer = 0

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal Scale As Integer)
                _Scale = Scale
            End Sub

            ' .. and show a copy to the outside world.
            Public Property Scale() As Integer
                Get
                    Return _Scale
                End Get
                Set(ByVal Value As Integer)
                    _Scale = Value
                End Set
            End Property

        End Class

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class Precision
            Inherits Attribute

            ' Keep a variable internally ...
            Private _Precision As Integer = 0

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal Precision As Integer)
                _Precision = Precision
            End Sub

            ' .. and show a copy to the outside world.
            Public Property Precision() As Integer
                Get
                    Return _Precision
                End Get
                Set(ByVal Value As Integer)
                    _Precision = Value
                End Set
            End Property

        End Class

#End Region

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

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class twodBarcode
            Inherits Attribute

            ' Keep a variable internally ...
            Private _twodBarcode As String

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal twodBarcode As String)
                _twodBarcode = twodBarcode
            End Sub

            ' .. and show a copy to the outside world.
            Public Property twodBarcode() As String
                Get
                    Return _twodBarcode
                End Get
                Set(ByVal Value As String)
                    _twodBarcode = Value
                End Set
            End Property

        End Class

        <AttributeUsage(AttributeTargets.Property)> _
        Public Class help
            Inherits Attribute

            ' Keep a variable internally ...
            Private _help As String = "There is no help."

            ' The constructor is called when the attribute is set.
            Public Sub New(ByVal help As String)
                _help = help
            End Sub

            ' .. and show a copy to the outside world.
            Public Property help() As String
                Get
                    Return _help
                End Get
                Set(ByVal Value As String)
                    _help = Value
                End Set
            End Property

        End Class
    End Module

#End Region

    Public Class CustomProperties : Inherits PropertyInfo

#Region "Constructor"

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

                If TypeOf attr Is nType Then
                    _nType = CType(attr, nType).nType
                End If
                If TypeOf attr Is Scale Then
                    _Scale = CType(attr, Scale).Scale
                End If
                If TypeOf attr Is Precision Then
                    _Precision = CType(attr, Precision).Precision
                End If

                If TypeOf attr Is twodBarcode Then
                    _twodBarcode = CType(attr, twodBarcode).twodBarcode
                End If
                If TypeOf attr Is help Then
                    _help = CType(attr, help).help
                End If

            Next
        End Sub

#End Region

#Region "Custom property attributes"

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

        Private _nType As String
        Public ReadOnly Property nType() As String
            Get
                Return _nType
            End Get
        End Property

        Private _Scale As Integer
        Public ReadOnly Property Scale() As Integer
            Get
                Return _Scale
            End Get
        End Property

        Private _Precision As Integer
        Public ReadOnly Property Precision() As Integer
            Get
                Return _Precision
            End Get
        End Property

        Private _twodBarcode As String
        Public ReadOnly Property twodBarcode() As String
            Get
                Return _twodBarcode
            End Get
        End Property

        Private _help As String = "No help available."
        Public ReadOnly Property help() As String
            Get
                Return _help
            End Get
        End Property

#End Region

#Region "Derived from Property"

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

#End Region

#Region "Get / set Property"

        Public Overloads Overrides Function GetValue(ByVal obj As Object, ByVal invokeAttr As System.Reflection.BindingFlags, ByVal binder As System.Reflection.Binder, ByVal index() As Object, ByVal culture As System.Globalization.CultureInfo) As Object
            Return _me.GetValue(obj, invokeAttr, binder, index, culture)
        End Function

        Public Overloads Overrides Sub SetValue(ByVal obj As Object, ByVal value As Object, ByVal invokeAttr As System.Reflection.BindingFlags, ByVal binder As System.Reflection.Binder, ByVal index() As Object, ByVal culture As System.Globalization.CultureInfo)
            Select Case _nType
                Case "Edm.String"
                    _me.SetValue(obj, value, invokeAttr, binder, index, culture)

                Case "Edm.Decimal"
                    _me.SetValue(obj, Decimal.Parse(value), invokeAttr, binder, index, culture)

                Case "Edm.DateTimeOffset"
                    _me.SetValue(obj, DateTime.Parse(value), invokeAttr, binder, index, culture)

                Case "Edm.Int64"
                    _me.SetValue(obj, Int64.Parse(value), invokeAttr, binder, index, culture)

            End Select

        End Sub

#End Region

    End Class

End Namespace