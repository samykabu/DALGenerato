Public Class clsColumn
    Private _ColumnName As String
    Private _DataType As String
    Private _Legth As Integer
    Private _Precision As Integer
    Private _IsNullable As Boolean
    Private _IsPrimaryKey As Boolean
    Private _IsForginKey As Boolean
    Private _IsIdentity As Boolean
    Private _VBDataType As String
    Private _SQLType As String
    Private _CSharpType As String

    Public Property ColumnName() As String
        Get
            Return _ColumnName
        End Get
        Set(ByVal Value As String)
            _ColumnName = Value
        End Set
    End Property
    Public Property DataType() As String
        Get
            Return _DataType
        End Get
        Set(ByVal Value As String)
            _DataType = Value
            If _DataType.IndexOf("identity") > -1 Then IsIdentity = True
        End Set
    End Property
    Public Property Length() As Integer
        Get
            Return _Legth
        End Get
        Set(ByVal Value As Integer)
            _Legth = Value
        End Set
    End Property
    Public Property Precision() As Integer
        Get
            Return _Precision
        End Get
        Set(ByVal Value As Integer)
            _Precision = Value
        End Set
    End Property
    Public Property IsNullable() As Boolean
        Get
            Return _IsNullable
        End Get
        Set(ByVal Value As Boolean)
            _IsNullable = Value
        End Set
    End Property
    Public Property IsPrimaryKey() As Boolean
        Get
            Return _IsPrimaryKey
        End Get
        Set(ByVal Value As Boolean)
            _IsPrimaryKey = Value
        End Set
    End Property
    Public Property IsForgineKey() As Boolean
        Get
            Return _IsForginKey
        End Get
        Set(ByVal Value As Boolean)
            _IsForginKey = False
        End Set
    End Property
    Public Property IsIdentity() As Boolean
        Get
            Return _IsIdentity
        End Get
        Set(ByVal Value As Boolean)
            _IsIdentity = Value
        End Set
    End Property    
    Public ReadOnly Property VBDataType() As String
        Get
            Dim tbuf As String
            If _DataType.ToLower.IndexOf(" ") > -1 Then
                tbuf = _DataType.ToLower.Substring(0, _DataType.ToLower.IndexOf(" "))
            Else
                tbuf = _DataType.ToLower
            End If
            Select Case tbuf
                Case "bit" : Return "Boolean"
                Case "tinyint" : Return "Byte"
                Case "varbinary" : Return "Byte()"
                Case "timestamp" : Return "byte()"
                Case "datetime" : Return "Date"
                Case "decimal" : Return "Decimal"
                Case "float" : Return "Double"
                Case "real" : Return "Single"
                Case "uniqueidentifier" : Return "Guid"
                Case "smallint" : Return "int16"
                Case "int" : Return "int32"
                Case "bigint" : Return "int64"
                Case "variant" : Return "object"
                Case "nvarchar" : Return "String"
                Case "varchar" : Return "String"
                Case "money" : Return "Currency"
                Case "image" : Return "Byte()"
                Case "text" : Return "string"
                Case "ntext" : Return "string"
                Case "nchar" : Return "string"
                Case "char" : Return "string"
            End Select
        End Get
    End Property
    Public ReadOnly Property GetTypes() As String
        Get
            Dim tbuf As String
            If _DataType.ToLower.IndexOf(" ") > -1 Then
                tbuf = _DataType.ToLower.Substring(0, _DataType.ToLower.IndexOf(" "))
            Else
                tbuf = _DataType.ToLower
            End If
            Select Case tbuf
                Case "bit" : Return "GetBoolean"
                Case "tinyint" : Return "GetByte"
                Case "varbinary" : Return "GetBytes"
                Case "timestamp" : Return "Getbytes"
                Case "datetime" : Return "GetDateTime"
                Case "decimal" : Return "GetDecimal"
                Case "float" : Return "GetDouble"
                Case "real" : Return "GetSingle"
                Case "uniqueidentifier" : Return "Guid"
                Case "smallint" : Return "Getint16"
                Case "int" : Return "Getint32"
                Case "bigint" : Return "Getint64"
                Case "variant" : Return "Getobject"
                Case "nvarchar" : Return "GetString"
                Case "varchar" : Return "GetString"
                Case "money" : Return "GetCurrency"
                Case "image" : Return "GetBytes"
                Case "text" : Return "Getstring"
                Case "ntext" : Return "Getstring"
                Case "nchar" : Return "Getstring"
                Case "char" : Return "Getstring"
            End Select
        End Get
    End Property
    Public ReadOnly Property SqlDataType() As String
        Get
            Dim tbuf As String
            If _DataType.ToLower.IndexOf(" ") > -1 Then
                tbuf = _DataType.ToLower.Substring(0, _DataType.ToLower.IndexOf(" "))
            Else
                tbuf = _DataType.ToLower
            End If
            Select Case tbuf
                Case "bit" : Return "SqlDbType.Bit"
                Case "tinyint" : Return "SqlDbType.TinyInt"
                Case "varbinary" : Return "SqlDbType.VarBinary"
                Case "datetime" : Return "SqlDbType.DateTime"
                Case "decimal" : Return "SqlDbType.Decimal"
                Case "float" : Return "SqlDbType.Float"
                Case "real" : Return "SqlDbType.Real"
                Case "uniqueidentifier" : Return "SqlDbType.UniqueIdentifier"
                Case "smallint" : Return "SqlDbType.SmallInt"
                Case "int" : Return "SqlDbType.Int"
                Case "bigint" : Return "SqlDbType.BigInt"
                Case "variant" : Return "SqlDbType.Variant"
                Case "nvarchar" : Return "SqlDbType.NVarChar"
                Case "varchar" : Return "SqlDbType.VarChar"
                Case "money" : Return "SqlDbType.Money"
                Case "char" : Return "SqlDbType.Char"
                Case "nchar" : Return "SqlDbType.nChar"
                Case "timestamp" : Return "SqlDbType.timestamp"
                Case "image" : Return "SqlDbType.Image"
                Case "text" : Return "SqlDbType.Text"
                Case "ntext" : Return "SqlDbType.ntext"
            End Select
        End Get
    End Property

End Class
