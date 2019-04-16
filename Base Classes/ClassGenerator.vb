Imports System.IO
Imports System.Text

Public Class ClassGenerator
    Private _Tables As ArrayList
    Private _Classes As New ArrayList
    Private OutputPath As String
    Private ClsPrf As String
    Public Event ClassSaved(ByVal ClassName As String)
    Public Event BeginConstructingClass(ByVal ClassName As String)
    Public Event ConstructingProperty(ByVal PropertyName As String)
    Public Event PropertyConstructed(ByVal PropertyName As String)
    Public Event PropertyFailed(ByVal PropertyName As String)
    Public Event ClassFaild(ByVal className As String)

    Public Sub ConstructClasses()

        Dim o As clsTable
        For Each o In _Tables
            Dim nClass As New clsClass
            Dim oculmns As clsColumn

            nClass.Name = o.TableName
            'Construct Class
            RaiseEvent BeginConstructingClass(nClass.Name)
            'Construct a property for each Column
            For Each oculmns In o.Columns
                Try
                    RaiseEvent ConstructingProperty(oculmns.ColumnName)
                    Dim oProperty As New clsProperty
                    oProperty.Name = oculmns.ColumnName
                    oProperty.PrivateVariable = "m_" + oculmns.ColumnName
                    oProperty.IsPrimaryKey = oculmns.IsPrimaryKey
                    oProperty.DataType = oculmns.VBDataType
                    oProperty.CheckForNull = Not oculmns.IsNullable
                    oProperty.SqlDataType = oculmns.SqlDataType
                    oProperty.GetDataType = oculmns.GetTypes
                    oProperty.IsIdentity = oculmns.IsIdentity
                    nClass.Properties.Add(oProperty)
                    RaiseEvent PropertyConstructed(oculmns.ColumnName)
                Catch ex As Exception
                    RaiseEvent PropertyFailed(oculmns.ColumnName)
                End Try
            Next
            If SaveClass(nClass) Then RaiseEvent ClassSaved(nClass.Name)
        Next
    End Sub
    Private Function SaveClass(ByVal [Class] As clsClass) As Boolean
        Dim oFile As FileStream = File.Open(OutputPath + ClsPrf + [Class].Name + ".vb", FileMode.Create)
        Try
            Dim ClassBuffer As String = [Class].GenerateCode
            oFile.Write(System.Text.Encoding.UTF8.GetBytes(ClassBuffer), 0, ClassBuffer.Length)
            oFile.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Sub New(ByVal ClassPrefix As String, ByVal Path As String, ByVal Tables As ArrayList)
        _Tables = Tables
        ClsPrf = ClassPrefix
        OutputPath = Path
    End Sub
End Class
