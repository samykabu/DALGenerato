Imports System.Text
Imports DALGenerator
Public Class ClsView
    Private strViewName As String
    Public Property Name() As String
        Get
            Return strViewName
        End Get
        Set(ByVal Value As String)
            strViewName = Value
        End Set
    End Property
    Public Shadows Function Tostring() As String
        Dim sb As New StringBuilder
        sb.Append("Imports System" + vbCrLf)
        sb.Append("Imports System.Data" + vbCrLf)
        sb.Append("Imports System.Data.SqlClient" + vbCrLf + vbCrLf)

        sb.Append("Public Class View" + strViewName + vbCrLf)
        sb.Append(Indent(1) + "Inherits DataObjectBase" + vbCrLf + vbCrLf)
        sb.Append(ClassTableName)
        sb.Append("End Class" + vbCrLf)
        Return sb.ToString
    End Function
    Private Function ClassTableName() As String
        Dim sb As New StringBuilder
        sb = WrLn(sb, Indent(1) + "Public Overrides ReadOnly Property TableName() As String")
        sb = WrLn(sb, Indent(2) + "Get")
        sb = WrLn(sb, Indent(3) + "Return " + Chr(34) + strViewName + Chr(34))
        sb = WrLn(sb, Indent(2) + "End Get")
        sb = WrLn(sb, Indent(1) + "End Property")
        Return sb.ToString
    End Function
    Private Function WrLn(ByVal sb As StringBuilder, ByVal str As String) As StringBuilder
        Return sb.Append(str + vbCrLf)
    End Function
    Private Function Indent(ByVal tab As Integer) As String
        Dim s As String
        Dim i As Integer
        For i = 0 To tab
            s = s + vbTab
        Next
        Return s
    End Function
End Class
