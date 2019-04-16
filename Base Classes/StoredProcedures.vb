Imports System.IO
Imports System.Text
Public Class StoredProcedures
    Inherits CollectionBase

    Public Function Add(ByVal obj As StoredProcedure) As Integer
        Return MyBase.List.Add(obj)
    End Function

    Public Function Item(ByVal index As Integer) As StoredProcedure
        Try
            Return CType(MyBase.List.Item(index), StoredProcedure)
        Catch ex As Exception
            Throw New Exception("Index out of range")
        End Try
    End Function

    Public Function Remove(ByVal obj As StoredProcedure) As Boolean
        Try
            MyBase.List.Remove(obj)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shadows Function RemoveAt(ByVal index As Integer) As Boolean
        Try
            MyBase.List.RemoveAt(index)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function SaveSP(ByVal Path As String) As Boolean
        Dim oFile As FileStream = File.Open(Path + "SP.sql", FileMode.Create)
        Try
            Dim ClassBuffer As String
            Dim oSp As StoredProcedure
            For Each oSp In Me.List
                ClassBuffer += oSp.InsertSP + vbCrLf
                ClassBuffer += oSp.UpdateSP + vbCrLf
                ClassBuffer += oSp.DeleteSP + vbCrLf
            Next
            oFile.Write(System.Text.Encoding.UTF8.GetBytes(ClassBuffer), 0, ClassBuffer.Length)
            oFile.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
