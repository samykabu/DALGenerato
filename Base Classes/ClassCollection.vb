Public Class ClassCollection
    Inherits CollectionBase

    Public Function Add(ByVal obj As clsClass) As Integer
        Return MyBase.List.Add(obj)
    End Function

    Public Function Item(ByVal index As Integer) As clsClass
        Try
            Return CType(MyBase.List.Item(index), clsClass)
        Catch ex As Exception
            Throw New Exception("Index out of range")
        End Try
    End Function

    Public Function Remove(ByVal obj As clsClass) As Boolean
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

End Class
