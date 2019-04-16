Imports System.Text

Public Class StoredProcedure
    Public Structure Parameter
        Public Name As String
        Public Type As String
        Public PrimaryKey As Boolean
        Public IsIdentity As Boolean
        Public Length As Integer
    End Structure

    Private _SPName As String
    Private oParameters As New ArrayList
    Public Sub Parameters(ByVal Name, ByVal Type, ByVal IsPrimaryKey, ByVal IsIdentity, ByVal Length)
        Dim oparam As New Parameter
        oparam.Name = Name
        oparam.PrimaryKey = IsPrimaryKey
        oparam.Type = Type
        oparam.IsIdentity = IsIdentity
        oparam.Length = Length
        oParameters.Add(oparam)
    End Sub
    Public WriteOnly Property SPName() As String
        Set(ByVal Value As String)
            _SPName = Value
        End Set
    End Property

    Public Function InsertSP() As String
        Dim sb As New StringBuilder
        Dim tb As New StringBuilder
        Dim vb As New StringBuilder

        sb.Append("if exists (select * from dbo.sysobjects where id=object_id(N'[Insert_")
        sb.Append(_SPName)
        sb.Append("]') and OBJECTPROPERTY(id,N'IsProcedure')=1)")
        sb.Append("drop procedure [Insert_" + _SPName)
        sb.Append("]")
        sb.Append(vbCrLf)
        sb.Append("GO")
        sb.Append(vbCrLf)
        sb.Append("Create Procedure [Insert_" + _SPName + "]" + vbCrLf)

        Dim oparm As Parameter
        Dim OpKey As Parameter
        Dim oIdent As Parameter
        Dim str As String = ""
        sb.Append("(" + vbCrLf)
        For Each oparm In oParameters
            If Not oparm.IsIdentity AndAlso Not oparm.Type = "timestamp" Then
                sb.Append(str + vbCrLf + "@" + oparm.Name + " " + oparm.Type)
                If oparm.Type = "nvarchar" Or oparm.Type = "varchar" Or oparm.Type = "nchar" Or oparm.Type = "char" Then sb.Append("(" + oparm.Length.ToString + ")")
                tb.Append(str + vbCrLf + "[" + oparm.Name + "]")
                vb.Append(str + vbCrLf + "@" + oparm.Name)
                str = ","
                If oparm.PrimaryKey Then OpKey = oparm
            Else
                If oparm.IsIdentity Then oIdent = oparm
            End If
        Next
        sb.Append(vbCrLf + ")" + vbCrLf + " AS INSERT INTO [" + _SPName + "]" + vbCrLf)
        sb.Append("(" + vbCrLf)
        sb.Append(tb)
        sb.Append(vbCrLf)
        sb.Append(") " + vbCrLf + " Values " + vbCrLf + "(" + vbCrLf)
        sb.Append(vb)
        sb.Append(vbCrLf)
        sb.Append(")")
        sb.Append(vbCrLf)
        sb.Append(" IF  @@Error=0 ")
        sb.Append(vbCrLf)
        If Not oIdent.Name Is Nothing Then
            sb.Append("Select Result=@@Identity")
        Else
            sb.Append("Select Result=@" + OpKey.Name)
        End If
        sb.Append(vbCrLf)
        sb.Append("Else")
        sb.Append(vbCrLf)
        sb.Append("Select Result=-1")
        sb.Append(vbCrLf)
        sb.Append("go")
        Return sb.ToString
    End Function
    Public Function UpdateSP() As String
        Dim Params As New StringBuilder
        Dim ParamData As New StringBuilder
        Dim sb As New StringBuilder
        Dim OpKey As Parameter

        sb.Append("if exists (select * from dbo.sysobjects where id=object_id(N'[Update_")
        sb.Append(_SPName)
        sb.Append("]') and OBJECTPROPERTY(id,N'IsProcedure')=1)")
        sb.Append("drop procedure [Update_" + _SPName)
        sb.Append("]")
        sb.Append(vbCrLf)
        sb.Append("GO")
        sb.Append(vbCrLf)
        sb.Append("Create Procedure [Update_" + _SPName + "]" + vbCrLf)

        Dim oparm As Parameter
        Dim str As String = ""
        Dim Condition As String = String.Empty
        sb.Append("(" + vbCrLf)
        For Each oparm In oParameters
            If Not oparm.IsIdentity AndAlso Not oparm.Type = "timestamp" Then
                ParamData.Append(str + "[" + oparm.Name + "]=@" + oparm.Name)
                sb.Append(str + "@" + oparm.Name + " " + oparm.Type)
                If oparm.Type = "nvarchar" Or oparm.Type = "varchar" Or oparm.Type = "nchar" Or oparm.Type = "char" Then sb.Append("(" + oparm.Length.ToString + ")")
                str = "," + vbCrLf
                If oparm.PrimaryKey Then
                    Condition = "[" + oparm.Name + "]=@" + oparm.Name
                    OpKey = oparm
                End If
            Else
                If oparm.IsIdentity Then
                    sb.Append(str + "@" + oparm.Name + " " + oparm.Type.Substring(0, oparm.Type.IndexOf(" ")) + "," + vbCrLf)
                    Condition = "[" + oparm.Name + "]=@" + oparm.Name
                    OpKey = oparm
                End If
            End If
        Next
        sb.Append(vbCrLf + ")" + vbCrLf + " AS Update [" + _SPName + "]" + vbCrLf)
        sb.Append("Set " + vbCrLf)
        sb.Append(ParamData.ToString)
        sb.Append(vbCrLf)
        sb.Append("Where  " + vbCrLf)
        sb.Append(Condition)
        sb.Append(vbCrLf)
        sb.Append(" IF  @@Error=0 ")
        sb.Append(vbCrLf)
        If OpKey.Name Is Nothing Then
            sb.Append("Select Result=0")
        Else
            sb.Append("Select Result=@" + OpKey.Name)
        End If
        sb.Append(vbCrLf)
        sb.Append("Else")
        sb.Append(vbCrLf)
        sb.Append("Select Result=-1")
        sb.Append(vbCrLf)
        sb.Append("go")
        Return sb.ToString
    End Function
    Public Function DeleteSP() As String
        Dim sb As New StringBuilder
        Dim tb As New StringBuilder
        Dim vb As New StringBuilder

        sb.Append("if exists (select * from dbo.sysobjects where id=object_id(N'[Delete_")
        sb.Append(_SPName)
        sb.Append("]') and OBJECTPROPERTY(id,N'IsProcedure')=1)")
        sb.Append("drop procedure [Delete_" + _SPName)
        sb.Append("]")
        sb.Append(vbCrLf)
        sb.Append("GO")
        sb.Append(vbCrLf)
        sb.Append("Create Procedure [Delete_" + _SPName + "]" + vbCrLf)

        Dim oparm As Parameter
        Dim str As String = ""
        Dim condition As String = String.Empty
        sb.Append("(" + vbCrLf)
        For Each oparm In oParameters

            If oparm.IsIdentity Then
                sb.Append(str + "@" + oparm.Name + " " + oparm.Type.Substring(0, oparm.Type.IndexOf(" ")) + vbCrLf)
                condition = "[" + oparm.Name + "]=@" + oparm.Name
                Exit For
            ElseIf oparm.PrimaryKey Then
                sb.Append(str + "@" + oparm.Name + " " + oparm.Type + vbCrLf)
                condition = "[" + oparm.Name + "]=@" + oparm.Name
                Exit For
            End If
        Next
        sb.Append(")" + vbCrLf + " AS Delete From [" + _SPName + "]" + vbCrLf)
        sb.Append("Where  " + vbCrLf)
        sb.Append(condition)
        sb.Append(vbCrLf)
        sb.Append(" IF  @@Error=0 ")
        sb.Append(vbCrLf)
        sb.Append("Select Result=0")
        sb.Append(vbCrLf)
        sb.Append("Else")
        sb.Append(vbCrLf)
        sb.Append("Select Result=-1")
        sb.Append(vbCrLf)
        sb.Append("go")
        Return sb.ToString
    End Function
End Class
