Imports System.Data.SqlClient
Public Class clsTables
    Private _Tables As New ArrayList
    Private oStoredProcedures As New StoredProcedures

    Public Event ReadingTable(ByVal TableName As String)
    Public Event TotalTables(ByVal TablesCount As Integer)

    Property Table(ByVal index As Integer) As clsTable
        Get
            Return CType(_Tables(index), clsTable)
        End Get
        Set(ByVal Value As clsTable)
            _Tables(index) = Value
        End Set
    End Property
    Public ReadOnly Property [StoredProcedures]() As StoredProcedures
        Get
            Return oStoredProcedures
        End Get
    End Property

    Property Tables() As ArrayList
        Get
            Return _Tables
        End Get
        Set(ByVal Value As ArrayList)
            _Tables = Value
        End Set
    End Property
    Public Function FillTables(ByVal servername As String, ByVal uid As String, ByVal pass As String, ByVal database As String) As Boolean
        Dim DataReader As SqlDataReader
        Dim DataConnection As New SqlConnection("server=" + servername + ";uid=" + uid + ";pwd=" + pass + ";database=" + database + ";")
        Dim DataCommand As New SqlCommand

        DataConnection.Open()
        DataCommand.Connection = DataConnection
        DataCommand.CommandText = "exec sp_tables null,null,null,""'TABLE'"""
        DataReader = DataCommand.ExecuteReader
        While DataReader.Read
            Dim DBTable As New clsTable("server=" + servername + ";uid=" + uid + ";pwd=" + pass + ";database=" + database + ";")
            RaiseEvent ReadingTable(DataReader("TABLE_NAME"))
            DBTable.TableName = DataReader("TABLE_NAME")
            _Tables.Add(DBTable)
            oStoredProcedures.Add(DBTable.StoredPro)
        End While
        If Not DataReader Is Nothing Then DataReader.Close()
        DataReader.Close()
    End Function


End Class
