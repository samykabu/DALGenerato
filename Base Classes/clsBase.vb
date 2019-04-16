Public MustInherit Class clsBase
    Public Function PropertyNullCheck(ByVal Prop As clsProperty) As String
        Select Case Prop.DataType.ToLower
            Case "int64", "int32", "int16", "byte", "decimal", "double", "single", "currency"
                If Prop.IsIdentity Or Prop.IsPrimaryKey Then
                    Return Indent(1) + "if " + Prop.PrivateVariable + "<=0 then " + vbCrLf + Indent(2) + " _Err =""" + Prop.Name + " Filed is required """ + vbCrLf + Indent(2) + "Return -1" + vbCrLf + Indent(1) + "End If" + vbCrLf
                ElseIf Prop.DataType.ToLower = "byte" Then
                    Return Indent(1) + "if " + Prop.PrivateVariable + "<0 then " + vbCrLf + Indent(2) + " _Err =""" + Prop.Name + " Filed is required """ + vbCrLf + Indent(2) + "Return -1" + vbCrLf + Indent(1) + "End If" + vbCrLf
                Else
                    Return Indent(1) + "if Not DatabaseUtility.Isnumber(" + Prop.PrivateVariable + ") then " + vbCrLf + Indent(2) + " _Err =""" + Prop.Name + " Filed is required """ + vbCrLf + Indent(2) + "Return -1" + vbCrLf + Indent(1) + "End If" + vbCrLf
                End If
            Case "datetime", "date"
                Return Indent(1) + "if Not DatabaseUtility.IsValidDate(" + Prop.PrivateVariable + ") then " + vbCrLf + Indent(2) + " _Err =""" + Prop.Name + " Filed is required """ + vbCrLf + Indent(2) + "Return -1" + vbCrLf + Indent(1) + "End If" + vbCrLf
            Case "byte()"
                Return Indent(1) + "IF " + Prop.PrivateVariable + ".length <0 then " + vbCrLf + Indent(2) + " _Err =""" + Prop.Name + " Filed is required """ + vbCrLf + Indent(2) + "Return -1" + vbCrLf + Indent(1) + "End If" + vbCrLf
            Case "string"
                Return Indent(1) + "If " + Prop.PrivateVariable + "=string.empty then " + vbCrLf + Indent(2) + " _Err =""" + Prop.Name + " Filed is required """ + vbCrLf + Indent(2) + "Return -1" + vbCrLf + Indent(1) + "End If" + vbCrLf
        End Select
    End Function

    Public Function PropertyMakeParam(ByVal prop As clsProperty) As String        
        If Not prop.CheckForNull Then
            Return "MakeParameter(""@" + prop.Name + Chr(34) + "," + GetHelperFunc(prop.DataType) + prop.PrivateVariable + "))" + vbCrLf
        Else
            Return "MakeParameter(""@" + prop.Name + Chr(34) + "," + prop.PrivateVariable + ")" + vbCrLf
        End If
    End Function

    Protected Function GetHelperFunc(ByVal datatype As String) As String
        Select Case datatype.ToLower
            Case "int64", "int32", "int16", "byte", "decimal", "double", "single", "currency"
                Return "DatabaseUtility.IntegerToDBNull("
            Case "datetime", "date"
                Return "DatabaseUtility.DateToDBNull("
            Case "string"
                Return "DatabaseUtility.StringToDBNull("
            Case Else
                Return "("
        End Select
    End Function
    Protected Function Indent(ByVal tab As Integer) As String
        Dim s As String
        Dim i As Integer
        For i = 0 To tab
            s = s + vbTab
        Next
        Return s
    End Function
End Class
