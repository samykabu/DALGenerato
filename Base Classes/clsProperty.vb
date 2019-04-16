Public Class clsProperty
    Inherits clsBase
    Implements IDal

    Private _PrivateVariable As String
    Private _PropertyName As String
    Private _Datatype As String
    Private _Readonly As Boolean = False
    Private _WriteOnly As Boolean = False
    Private _CheckForNull As Boolean
    Private _SqlDataType As String
    Private _GetDataType As String
    Private _ISPrimaryKey As Boolean = False
    Private _ISIdentity As Boolean = False
    Private _PropertyNullCheck As String = String.Empty
    Private _PropertyParam As String = String.Empty

    Public Property IsPrimaryKey() As Boolean
        Get
            Return _ISPrimaryKey
        End Get
        Set(ByVal Value As Boolean)
            _ISPrimaryKey = Value
        End Set
    End Property
    Public Property IsIdentity() As Boolean
        Get
            Return _ISIdentity
        End Get
        Set(ByVal Value As Boolean)
            _ISIdentity = Value
        End Set
    End Property
    Public Property CheckForNull() As Boolean
        Get
            Return _CheckForNull
        End Get
        Set(ByVal Value As Boolean)
            _CheckForNull = Value
        End Set
    End Property
    Public Property SqlDataType() As String
        Get
            Return _SqlDataType
        End Get
        Set(ByVal Value As String)
            _SqlDataType = Value
        End Set
    End Property
    Public Property GetDataType() As String
        Get
            Return _GetDataType
        End Get
        Set(ByVal Value As String)
            _GetDataType = Value
        End Set
    End Property

    Public Property PrivateVariable() As String
        Get
            Return _PrivateVariable
        End Get
        Set(ByVal Value As String)
            _PrivateVariable = Value
        End Set
    End Property
    Public Property Name() As String
        Get
            Return _PropertyName
        End Get
        Set(ByVal Value As String)
            _PropertyName = Value
        End Set
    End Property
    Public Property DataType() As String
        Get
            Return _Datatype
        End Get
        Set(ByVal Value As String)
            _Datatype = Value
        End Set
    End Property
    Public Property [Readonly]() As Boolean
        Get
            Return _Readonly
        End Get
        Set(ByVal Value As Boolean)
            _Readonly = Value
        End Set
    End Property
    Public Property [Writeonly]() As Boolean
        Get
            Return _WriteOnly
        End Get
        Set(ByVal Value As Boolean)
            _WriteOnly = Value
        End Set
    End Property

    Public ReadOnly Property NullCheckStatement() As String
        Get
            If _CheckForNull Then _PropertyNullCheck = Me.PropertyNullCheck(Me)
            Return _PropertyNullCheck
        End Get
    End Property
    Public ReadOnly Property Parameter() As String
        Get
            _PropertyParam = Me.PropertyMakeParam(Me)
            Return _PropertyParam
        End Get
    End Property

    Public Function GenerateCode() As String Implements Interfaces.IDal.GenerateCode
        Dim sb As New System.Text.StringBuilder
        Try
            sb.Append(Indent(1) + "Public ")
            If _Readonly Then sb.Append("Readonly ")
            If _WriteOnly Then sb.Append("WriteOnly ")
            sb.Append("Property ")
            sb.Append(_PropertyName)
            sb.Append("() as ")
            sb.Append(_Datatype.ToString)
            sb.Append(vbCrLf)
            If Not _WriteOnly Then
                sb.Append(Indent(2) + "Get")
                sb.Append(vbCrLf)
                sb.Append(Indent(3) + "Return " + _PrivateVariable)
                sb.Append(vbCrLf)
                sb.Append(Indent(2) + "End Get")
                sb.Append(vbCrLf)
            End If
            If Not _Readonly Then
                sb.Append(Indent(2) + "set(Byval Value as " + _Datatype + ")")
                sb.Append(vbCrLf)
                sb.Append(Indent(3) + _PrivateVariable + "=Value")
                sb.Append(vbCrLf)
                If _ISPrimaryKey Then sb.Append(Indent(3) + "LoadValues()" + vbCrLf)
                sb.Append(Indent(2) + "End Set")
                sb.Append(vbCrLf)
            End If
            sb.Append(Indent(1) + "End Property")
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try
        Return sb.ToString
    End Function
    Public Function PrivateVariableDeclare() As String
        Return Indent(1) + "Private " + _PrivateVariable + " as " + _Datatype
    End Function
End Class
