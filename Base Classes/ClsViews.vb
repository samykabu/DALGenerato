Public Class ClsViews
    Inherits CollectionBase
    Public Function Add(ByVal obj As ClsView) As Integer
        Return MyBase.List.Add(obj)
    End Function
    Public Function Item(ByVal indx As Integer) As ClsView
        Try
            Return DirectCast(MyBase.List(indx), ClsView)
        Catch ex As Exception
            Throw New Exception("Index out of range")
        End Try
    End Function
    Public Function Remove(ByVal obj As ClsView) As Boolean
        Try
            MyBase.List.Remove(obj)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Shadows Function Removeat(ByVal indx As Integer) As Boolean
        Try
            MyBase.List.RemoveAt(indx)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class
