Imports System.Text
Public Class clsClass
    Implements IDal

    Private _Properties As New PropertyCollection
    Private _Name As String
    Private _CassPrefix As String = String.Empty
    Private _Namespace As String = String.Empty
    Private PrimaryKey As clsProperty

    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal Value As String)
            _Name = Value
        End Set
    End Property
    Public ReadOnly Property Properties() As PropertyCollection
        Get
            Return _Properties
        End Get
    End Property


    Public Function GenerateCode() As String Implements Interfaces.IDal.GenerateCode
        Dim PrivateVariablesDeclaration As New StringBuilder
        Dim PropertiesBuffer As New StringBuilder
        Dim MethodsBuffer As New StringBuilder
        Dim FunctionsBuffer As New StringBuilder

        Dim sb As New StringBuilder
        Dim oProperty As clsProperty

        For Each oProperty In _Properties
            PropertiesBuffer.Append(oProperty.GenerateCode + vbCrLf + vbCrLf)
            PrivateVariablesDeclaration.Append(oProperty.PrivateVariableDeclare + vbCrLf)
        Next

        sb.Append("Imports System" + vbCrLf)
        sb.Append("Imports System.Data" + vbCrLf)
        sb.Append("Imports System.Data.SqlClient" + vbCrLf + vbCrLf)


        If _Namespace <> String.Empty Then sb.Append("NameSpace " + _Namespace + vbCrLf + vbCrLf)
        sb.Append("Public Class " + _Name + vbCrLf)
        sb.Append(Indent(1) + "Inherits DataObjectBase" + vbCrLf + vbCrLf)
        sb.Append("#Region ""Private Variables" + Chr(34) + vbCrLf)
        sb.Append(PrivateVariablesDeclaration.ToString)
        sb.Append(Indent(1) + "Private _ExistInDB as boolean" + vbCrLf)
        sb.Append(Indent(1) + "Public  _Err As String = String.Empty" + vbCrLf + vbCrLf)
        sb.Append("#End Region " + vbCrLf)
        sb.Append("#Region ""Public Properties" + Chr(34) + vbCrLf)
        sb.Append(PropertiesBuffer.ToString + vbCrLf)
        sb.Append(Indent(1) + "Public readonly property ExistInDB as boolean" + vbCrLf)
        sb.Append(Indent(2) + "Get" + vbCrLf)
        sb.Append(Indent(3) + "Return _ExistInDb" + vbCrLf)
        sb.Append(Indent(2) + "End Get" + vbCrLf)
        sb.Append(Indent(1) + "End Property" + vbCrLf)
        sb.Append("#End Region " + vbCrLf)
        sb.Append("#Region ""Public Functions" + Chr(34) + vbCrLf)
        sb.Append(AddNewMethod() + vbCrLf)
        sb.Append(Update() + vbCrLf)
        sb.Append(Delete() + vbCrLf)
        sb.Append(LoadValues() + vbCrLf)
        sb.Append(ClassTableName() + vbCrLf)
        sb.Append("#End Region " + vbCrLf)
        sb.Append("#Region ""Constructors""" + vbCrLf)
        sb.Append(Constructors)
        sb.Append("#End Region " + vbCrLf)
        sb.Append("End Class" + vbCrLf)
        If _Namespace <> String.Empty Then sb.Append("End Namespace" + vbCrLf)
        PrivateVariablesDeclaration = Nothing
        PropertiesBuffer = Nothing
        MethodsBuffer = Nothing
        FunctionsBuffer = Nothing
        Return sb.ToString
    End Function
    Private Function Indent(ByVal tab As Integer) As String
        Dim s As String
        Dim i As Integer
        For i = 0 To tab
            s = s + vbTab
        Next
        Return s
    End Function

    Private Function Constructors() As String
        Dim sb As New StringBuilder
        sb.Append("Public Sub New()" + vbCrLf)
        sb.Append("End sub" + vbCrLf + vbCrLf)
        sb.Append("Public sub New(PK as " + PrimaryKey.DataType + ")")
        sb.Append(vbCrLf + Indent(1) + PrimaryKey.Name + "=PK" + vbCrLf)
        sb.Append("End sub" + vbCrLf + vbCrLf)
        Return sb.ToString
    End Function

    Private Function AddNewMethod() As String
        Dim sb As New StringBuilder
        sb = WrLn(sb, "Public Function AddNew() as int64")
        sb = WrLn(sb, Indent(1) + "Dim MyReader as SqlDataReader")


        Dim oProp As clsProperty
        Dim osbprop As New StringBuilder
        Dim osbparm As New StringBuilder
        Dim ParamIndex As Integer = 0
        Dim oKeyProp As clsProperty
        Try
            For Each oProp In _Properties
                If oProp.IsPrimaryKey Then
                    oKeyProp = oProp
                    PrimaryKey = oProp
                End If
                If Not oProp.IsIdentity And Not oProp.SqlDataType.ToLower.IndexOf("timestamp") > -1 Then
                    osbprop.Append(oProp.NullCheckStatement)
                    osbparm.Append(Indent(1) + "Params(" + ParamIndex.ToString + ")=" + oProp.Parameter)
                    ParamIndex += 1
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try
        sb = WrLn(sb, Indent(1) + "Dim Params(" + (ParamIndex - 1).ToString + ") as SqlParameter")
        sb.Append(vbCrLf)
        sb.Append(osbprop)
        sb.Append(vbCrLf)
        sb.Append(osbparm)
        sb.Append(vbCrLf)

        sb = WrLn(sb, Indent(1) + "Try")
        sb = WrLn(sb, Indent(2) + " RunProcedure(""insert_" + Chr(34) + " + TableName " + ",Params,myReader)")
        sb = WrLn(sb, Indent(2) + "If Not myReader is Nothing then ")
        sb = WrLn(sb, Indent(3) + "If (myReader.Read) andalso myReader(0)>0 then")
        sb = WrLn(sb, Indent(4) + " Dim Res as int64=myReader(0)")
        sb = WrLn(sb, Indent(4) + "myReader.close")
        sb = WrLn(sb, Indent(4) + "myReader=Nothing")
        sb = WrLn(sb, Indent(4) + " if bIsLocalConenction=true andalso oCurrentTransaction is nothing then")
        sb = WrLn(sb, Indent(5) + "oSqlConnection.close")
        sb = WrLn(sb, Indent(5) + "oSqlConnection=nothing")
        sb = WrLn(sb, Indent(4) + "End if")
        sb = WrLn(sb, Indent(4) + oKeyProp.PrivateVariable + "= res")
        sb = WrLn(sb, Indent(4) + "return Res")
        sb = WrLn(sb, Indent(3) + "Else")
        sb = WrLn(sb, Indent(4) + "_Err=""Error Hass been detected while Inserting record please try again"" ")
        sb = WrLn(sb, Indent(4) + "myreader.close")
        sb = WrLn(sb, Indent(4) + "myReader=Nothing")
        sb = WrLn(sb, Indent(4) + "return -1")
        sb = WrLn(sb, Indent(3) + "end if")
        sb = WrLn(sb, Indent(2) + "end if")
        sb = WrLn(sb, Indent(1) + "Catch ex as Exception")
        sb = WrLn(sb, Indent(2) + "if not myreader is nothing then ")
        sb = WrLn(sb, Indent(3) + "myreader=nothing")
        sb = WrLn(sb, Indent(2) + "End IF")
        sb = WrLn(sb, Indent(2) + "Throw new Exception(""Error while Adding Record To Databse :""+ex.message)")
        sb = WrLn(sb, Indent(1) + "end try")
        sb = WrLn(sb, "End Function")
        Return sb.ToString
    End Function
    Private Function Update() As String
        Dim sb As New StringBuilder
        sb = WrLn(sb, "Public Function Update() as boolean")
        sb = WrLn(sb, Indent(1) + "Dim MyReader as SqlDataReader")

        Dim oProp As clsProperty
        Dim osbprop As New StringBuilder
        Dim osbparm As New StringBuilder
        Dim ParamIndex As Integer = 0
        Dim oKeyProp As clsProperty

        For Each oProp In _Properties
            If Not oProp.SqlDataType.ToLower.IndexOf("timestamp") > -1 Then
                osbprop.Append(oProp.NullCheckStatement)
                osbparm.Append(Indent(1) + "Params(" + ParamIndex.ToString + ")=" + oProp.Parameter)
                ParamIndex += 1
                If oProp.IsPrimaryKey Then
                    oKeyProp = oProp
                End If
            Else
                osbprop.Append(oProp.NullCheckStatement)
                oKeyProp = oProp
            End If
        Next
        sb = WrLn(sb, Indent(1) + "Dim Params(" + (ParamIndex - 1).ToString + ") as SqlParameter")
        sb.Append(vbCrLf)
        sb.Append(osbprop)
        sb.Append(vbCrLf)
        sb.Append(osbparm)
        sb.Append(vbCrLf)

        sb = WrLn(sb, Indent(1) + "Try")
        sb = WrLn(sb, Indent(2) + " RunProcedure(""Update_" + Chr(34) + " + TableName " + ",Params,myReader)")
        sb = WrLn(sb, Indent(2) + "If Not myReader is Nothing then ")
        sb = WrLn(sb, Indent(3) + "If (myReader.Read) andalso myReader(0)>-1 then")
        sb = WrLn(sb, Indent(4) + " Dim Res as int64=myReader(0)")

        sb = WrLn(sb, Indent(4) + "myReader.close")
        sb = WrLn(sb, Indent(4) + "myReader=Nothing")
        sb = WrLn(sb, Indent(4) + " if bIsLocalConenction=true andalso oCurrentTransaction is nothing then")
        sb = WrLn(sb, Indent(5) + "oSqlConnection.close")
        sb = WrLn(sb, Indent(5) + "oSqlConnection=nothing")
        sb = WrLn(sb, Indent(4) + "End if")

        sb = WrLn(sb, Indent(4) + oKeyProp.PrivateVariable + "= res")
        sb = WrLn(sb, Indent(4) + "return true")
        sb = WrLn(sb, Indent(3) + "Else")
        sb = WrLn(sb, Indent(4) + "myreader.close")
        sb = WrLn(sb, Indent(4) + "myReader=Nothing")
        sb = WrLn(sb, Indent(4) + "_Err=""Error Hass been detected while Updating record please try again"" ")
        sb = WrLn(sb, Indent(4) + "return false")
        sb = WrLn(sb, Indent(3) + "end if")
        sb = WrLn(sb, Indent(2) + "end if")
        sb = WrLn(sb, Indent(1) + "Catch ex as Exception")
        sb = WrLn(sb, Indent(2) + " if bIsLocalConenction=true andalso oCurrentTransaction is nothing then")
        sb = WrLn(sb, Indent(3) + "oSqlConnection.close")
        sb = WrLn(sb, Indent(3) + "oSqlConnection=nothing")
        sb = WrLn(sb, Indent(2) + "End if")
        sb = WrLn(sb, Indent(2) + "Throw new Exception(""Error while Adding Record To Databse :""+ex.message)")
        sb = WrLn(sb, Indent(2) + "return false")
        sb = WrLn(sb, Indent(1) + "end try")
        sb = WrLn(sb, "End Function")
        Return sb.ToString
    End Function

    Private Function Delete() As String
        Dim sb As New StringBuilder
        sb = WrLn(sb, "Public Function Delete() as boolean")
        sb = WrLn(sb, Indent(1) + "Dim MyReader as SqlDataReader")
        sb = WrLn(sb, Indent(1) + "Dim Params(0)" + " as SqlParameter")

        Dim oProp As clsProperty
        Dim osbprop As New StringBuilder
        Dim osbparm As New StringBuilder
        Dim ParamIndex As Integer = 0
        Dim oKeyProp As clsProperty

        For Each oProp In _Properties
            If oProp.IsPrimaryKey Then
                osbprop.Append(Indent(1) + "If " + oProp.PrivateVariable + "=0 then return false" + vbCrLf)
                osbparm.Append(Indent(1) + "Params(" + ParamIndex.ToString + ")=" + oProp.Parameter)
                ParamIndex += 1
                oKeyProp = oProp
            End If
        Next
        sb.Append(vbCrLf)
        sb.Append(osbprop)
        sb.Append(vbCrLf)
        sb.Append(osbparm)
        sb.Append(vbCrLf)
        sb = WrLn(sb, Indent(1) + "Try")
        sb = WrLn(sb, Indent(2) + " RunProcedure(""Delete_" + Chr(34) + " + TableName " + ",Params,myReader)")

        sb = WrLn(sb, Indent(2) + "myReader.close")
        sb = WrLn(sb, Indent(2) + "myReader=Nothing")
        sb = WrLn(sb, Indent(2) + " if bIsLocalConenction=true andalso oCurrentTransaction is nothing then")
        sb = WrLn(sb, Indent(3) + "oSqlConnection.close")
        sb = WrLn(sb, Indent(3) + "oSqlConnection=nothing")
        sb = WrLn(sb, Indent(2) + "End if")

        sb = WrLn(sb, Indent(2) + "return True")
        sb = WrLn(sb, Indent(1) + "Catch ex as Exception")
        sb = WrLn(sb, Indent(4) + "myReader=Nothing")
        sb = WrLn(sb, Indent(4) + " if bIsLocalConenction=true andalso oCurrentTransaction is nothing then")
        sb = WrLn(sb, Indent(5) + "oSqlConnection.close")
        sb = WrLn(sb, Indent(5) + "oSqlConnection=nothing")
        sb = WrLn(sb, Indent(4) + "End if")
        sb = WrLn(sb, Indent(2) + "Throw new Exception(""Error while Adding Record To Databse :""+ex.message)")
        sb = WrLn(sb, Indent(2) + "return False")
        sb = WrLn(sb, Indent(1) + "end try")
        sb = WrLn(sb, "End Function")
        Return sb.ToString
    End Function
    Private Function LoadValues() As String
        Dim oProp As clsProperty
        Dim sb As New StringBuilder
        Dim Index As Integer

        For Each oProp In _Properties
            If oProp.IsPrimaryKey Then
                Exit For
            End If
        Next
        sb = WrLn(sb, "Private sub LoadValues() ")
        sb = WrLn(sb, Indent(1) + "Dim MyReader as SqlDataReader")
        sb = WrLn(sb, Indent(1) + "RunSqlST(""Select * from " + Me.Name + " where " + oProp.Name + "=" + Chr(34) + "+" + oProp.PrivateVariable + ".ToString,myReader)")
        sb = WrLn(sb, Indent(1) + "If Not myReader is nothing andalso myReader.read then")
        For Each oProp In _Properties
            If oProp.DataType.ToLower <> "byte()" AndAlso oProp.DataType.ToLower <> "timestamp" Then
                If Not oProp.CheckForNull Then
                    Select Case oProp.DataType.ToLower
                        Case "string", "char"
                            sb = WrLn(sb, Indent(2) + oProp.PrivateVariable + "=DatabaseUtility.DBNullToString(myReader(" + Index.ToString + "))")
                        Case "int64", "int16", "int32", "integer"
                            sb = WrLn(sb, Indent(2) + oProp.PrivateVariable + "=DatabaseUtility.DBNullToInteger(myReader(" + Index.ToString + "))")
                        Case "date", "datetime"
                            sb = WrLn(sb, Indent(2) + oProp.PrivateVariable + "=DatabaseUtility.DBNullToDate(myReader(" + Index.ToString + "))")                                                   
                        Case Else
                            sb = WrLn(sb, Indent(2) + oProp.PrivateVariable + "=myReader(" + Index.ToString + ")")
                    End Select
                Else
                    sb = WrLn(sb, Indent(2) + oProp.PrivateVariable + "=myReader." + oProp.GetDataType + "(" + Index.ToString + ")")
                End If
            ElseIf oProp.DataType.ToLower = "byte()" AndAlso oProp.DataType.ToLower <> "timestamp" Then
                sb = WrLn(sb, Indent(2) + oProp.PrivateVariable + "=ctype(myReader(" + Index.ToString + "),byte())")
            End If
            Index += 1
        Next
        sb = WrLn(sb, Indent(2) + "_ExistInDB=true")
        sb = WrLn(sb, Indent(2) + "Myreader.Close")
        sb = WrLn(sb, Indent(2) + "Myreader=nothing")

        sb = WrLn(sb, Indent(2) + " if bIsLocalConenction=true andalso oCurrentTransaction is nothing then")
        sb = WrLn(sb, Indent(3) + "oSqlConnection.close")
        sb = WrLn(sb, Indent(3) + "oSqlConnection=nothing")
        sb = WrLn(sb, Indent(2) + "End if")

        sb = WrLn(sb, Indent(1) + "else")
        sb = WrLn(sb, Indent(2) + "_ExistInDB=false")
        sb = WrLn(sb, Indent(2) + "Myreader=nothing")

        sb = WrLn(sb, Indent(2) + " if bIsLocalConenction=true andalso oCurrentTransaction is nothing then")
        sb = WrLn(sb, Indent(3) + "oSqlConnection.close")
        sb = WrLn(sb, Indent(3) + "oSqlConnection=nothing")
        sb = WrLn(sb, Indent(2) + "End if")

        sb = WrLn(sb, Indent(2) + "_Err=" + Chr(34) + "Record Not Fount" + Chr(34))
        sb = WrLn(sb, Indent(1) + "End  if")
        sb = WrLn(sb, Indent(2) + "End Sub")
        Return sb.ToString
    End Function

    Private Function ClassTableName() As String
        Dim sb As New StringBuilder
        sb = WrLn(sb, Indent(1) + "Public Overrides ReadOnly Property TableName() As String")
        sb = WrLn(sb, Indent(2) + "Get")
        sb = WrLn(sb, Indent(3) + "Return " + Chr(34) + Me.Name + Chr(34))
        sb = WrLn(sb, Indent(2) + "End Get")
        sb = WrLn(sb, Indent(1) + "End Property")
        Return sb.ToString
    End Function

    Private Function WrLn(ByVal sb As StringBuilder, ByVal str As String) As StringBuilder
        Return sb.Append(str + vbCrLf)
    End Function
    Private Function PropertyNullCheck(ByVal Prop As clsProperty) As String
        Select Case Prop.DataType.ToLower
            Case "int64"
                Return Prop.PrivateVariable + " <=0 then " + vbCrLf + " _Err =""" + Prop.Name + " Filed is required """ + vbCrLf + "Return -1" + vbCrLf + "End If"
        End Select
    End Function

End Class
