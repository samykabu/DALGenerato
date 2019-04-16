Imports System.Data.SqlClient

Public Class clsTable
    Private _TableName As String
    Private _Columns As New clsColumns
    Private _PrimaryKeys As New ArrayList
    Private _ForginKeys As New ArrayList
    Private _StoredProcedure As New StoredProcedure
    Private ConnectionStr As String

    Public Property TableName() As String
        Get
            Return _TableName
        End Get
        Set(ByVal Value As String)
            _TableName = Value
            _StoredProcedure.SPName = _TableName
            FillColumns()
        End Set
    End Property
    Public ReadOnly Property StoredPro() As StoredProcedure
        Get
            Return _StoredProcedure
        End Get
    End Property

    Public ReadOnly Property Columns() As clsColumns
        Get
            Return _Columns
        End Get
    End Property

    Private Sub FillColumns()
        Dim DataReader As SqlDataReader
        Dim DataConnection As New SqlConnection(ConnectionStr)
        Dim DataCommand As New SqlCommand

        DataConnection.Open()
        DataCommand.Connection = DataConnection
        DataCommand.CommandText = "EXEC sp_pKeys '" + _TableName + "'"
        DataReader = DataCommand.ExecuteReader
        While DataReader.Read
            _PrimaryKeys.Add(DataReader("Column_Name"))
            Debug.Write(_TableName + "----->" + DataReader("Column_Name") + vbCrLf)
        End While
        If Not DataReader Is Nothing Then DataReader.Close()
        DataReader.Close()
        DataReader = Nothing


        DataCommand.CommandText = "EXEC sp_columns @table_name = '" + _TableName + "'"
        DataReader = DataCommand.ExecuteReader
        While DataReader.Read
            Dim DBColumn As New clsColumn
            DBColumn.ColumnName = DataReader("COLUMN_NAME")
            DBColumn.DataType = DataReader("TYPE_NAME")
            DBColumn.Length = DataReader("LENGTH")
            DBColumn.Precision = DataReader("PRECISION")
            DBColumn.IsNullable = DataReader("NULLABLE")
            If ISPrimaryKeyColumn(DBColumn.ColumnName) Then DBColumn.IsPrimaryKey = True
            If IsForginKeyColumn(DBColumn.ColumnName) Then DBColumn.IsForgineKey = True
            Select Case DBColumn.DataType.ToLower
                Case "varchar"
                    _StoredProcedure.Parameters(DBColumn.ColumnName, DBColumn.DataType + "(" + DBColumn.Length.ToString + ")", DBColumn.IsPrimaryKey, DBColumn.IsIdentity, DBColumn.Length)
                Case "nvarchar"
                    _StoredProcedure.Parameters(DBColumn.ColumnName, DBColumn.DataType + "(" + (DBColumn.Length / 2).ToString + ")", DBColumn.IsPrimaryKey, DBColumn.IsIdentity, DBColumn.Length)
                Case "char"
                    _StoredProcedure.Parameters(DBColumn.ColumnName, DBColumn.DataType + "(" + DBColumn.Length.ToString + ")", DBColumn.IsPrimaryKey, DBColumn.IsIdentity, DBColumn.Length)
                Case "nchar"
                    _StoredProcedure.Parameters(DBColumn.ColumnName, DBColumn.DataType + "(" + (DBColumn.Length / 2).ToString + ")", DBColumn.IsPrimaryKey, DBColumn.IsIdentity, DBColumn.Length)
                Case Else
                    _StoredProcedure.Parameters(DBColumn.ColumnName, DBColumn.DataType, DBColumn.IsPrimaryKey, DBColumn.IsIdentity, DBColumn.Length)
            End Select
            
            _Columns.Add(DBColumn)
        End While
        If Not DataReader Is Nothing Then DataReader.Close()
        DataReader.Close()
        DataReader = Nothing
    End Sub
    Private Function ISPrimaryKeyColumn(ByVal ColumnName As String) As Boolean
        Dim oStr As String
        For Each oStr In _PrimaryKeys
            If oStr.ToLower = ColumnName.ToLower Then Return True
        Next
        Return False
    End Function
    Private Function IsForginKeyColumn(ByVal ColumnName As String) As Boolean
        Dim oStr As String
        For Each oStr In _ForginKeys
            If oStr.ToLower = ColumnName.ToLower Then Return True
        Next
        Return False
    End Function
    Public Function FindColumn(ByVal ColumnName As String) As clsColumn
        Dim oc As clsColumn
        For Each oc In _Columns
            If oc.ColumnName = ColumnName Then
                Return oc
            End If
        Next
    End Function
    Public Sub New(ByVal ConnectionStr As String)
        Me.ConnectionStr = ConnectionStr
    End Sub
End Class
