Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Tables As System.Windows.Forms.TabPage
    Friend WithEvents DataBase As System.Windows.Forms.TabPage
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Views As System.Windows.Forms.TabPage
    Friend WithEvents Options As System.Windows.Forms.TabPage
    Friend WithEvents SPOption As System.Windows.Forms.TabPage
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TreeView2 As System.Windows.Forms.TreeView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents TablesView As System.Windows.Forms.TreeView
    Friend WithEvents Log As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Tables = New System.Windows.Forms.TabPage
        Me.TablesView = New System.Windows.Forms.TreeView
        Me.DataBase = New System.Windows.Forms.TabPage
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.TextBox8 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Views = New System.Windows.Forms.TabPage
        Me.TreeView2 = New System.Windows.Forms.TreeView
        Me.Options = New System.Windows.Forms.TabPage
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SPOption = New System.Windows.Forms.TabPage
        Me.Button3 = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.Log = New System.Windows.Forms.TextBox
        Me.Tables.SuspendLayout()
        Me.DataBase.SuspendLayout()
        Me.Views.SuspendLayout()
        Me.Options.SuspendLayout()
        Me.SPOption.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tables
        '
        Me.Tables.Controls.Add(Me.TablesView)
        Me.Tables.Location = New System.Drawing.Point(4, 22)
        Me.Tables.Name = "Tables"
        Me.Tables.Size = New System.Drawing.Size(400, 597)
        Me.Tables.TabIndex = 1
        Me.Tables.Text = "Tables"
        '
        'TablesView
        '
        Me.TablesView.CheckBoxes = True
        Me.TablesView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TablesView.ImageIndex = -1
        Me.TablesView.Location = New System.Drawing.Point(0, 0)
        Me.TablesView.Name = "TablesView"
        Me.TablesView.SelectedImageIndex = -1
        Me.TablesView.Size = New System.Drawing.Size(400, 597)
        Me.TablesView.TabIndex = 0
        '
        'DataBase
        '
        Me.DataBase.Controls.Add(Me.Button4)
        Me.DataBase.Controls.Add(Me.Label8)
        Me.DataBase.Controls.Add(Me.TextBox8)
        Me.DataBase.Controls.Add(Me.Label7)
        Me.DataBase.Controls.Add(Me.TextBox7)
        Me.DataBase.Controls.Add(Me.Label6)
        Me.DataBase.Controls.Add(Me.TextBox6)
        Me.DataBase.Controls.Add(Me.Button2)
        Me.DataBase.Controls.Add(Me.Button1)
        Me.DataBase.Controls.Add(Me.Label3)
        Me.DataBase.Controls.Add(Me.TextBox3)
        Me.DataBase.Controls.Add(Me.TextBox2)
        Me.DataBase.Controls.Add(Me.Label2)
        Me.DataBase.Controls.Add(Me.TextBox1)
        Me.DataBase.Controls.Add(Me.Label1)
        Me.DataBase.Location = New System.Drawing.Point(4, 22)
        Me.DataBase.Name = "DataBase"
        Me.DataBase.Size = New System.Drawing.Size(400, 597)
        Me.DataBase.TabIndex = 0
        Me.DataBase.Tag = ""
        Me.DataBase.Text = "DataBase Connection"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(176, 320)
        Me.Button4.Name = "Button4"
        Me.Button4.TabIndex = 14
        Me.Button4.Text = "Exit"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(24, 171)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 16)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Output Path"
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(128, 168)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(200, 20)
        Me.TextBox8.TabIndex = 12
        Me.TextBox8.Text = "D:\SS\"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 149)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 16)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Class Prefix"
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(128, 144)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(200, 20)
        Me.TextBox7.TabIndex = 10
        Me.TextBox7.Text = "Dal"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 122)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 16)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "DataBase Name"
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(128, 120)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(200, 20)
        Me.TextBox6.TabIndex = 8
        Me.TextBox6.Text = ""
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(208, 264)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Disconnect"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(128, 264)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Connect"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Server Name"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(128, 96)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(200, 20)
        Me.TextBox3.TabIndex = 4
        Me.TextBox3.Text = "(local)"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(128, 69)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(200, 20)
        Me.TextBox2.TabIndex = 3
        Me.TextBox2.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Password"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(128, 43)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(200, 20)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "sa"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "User Name"
        '
        'Views
        '
        Me.Views.Controls.Add(Me.TreeView2)
        Me.Views.Location = New System.Drawing.Point(4, 22)
        Me.Views.Name = "Views"
        Me.Views.Size = New System.Drawing.Size(400, 597)
        Me.Views.TabIndex = 2
        Me.Views.Text = "Views"
        '
        'TreeView2
        '
        Me.TreeView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView2.ImageIndex = -1
        Me.TreeView2.Location = New System.Drawing.Point(0, 0)
        Me.TreeView2.Name = "TreeView2"
        Me.TreeView2.SelectedImageIndex = -1
        Me.TreeView2.Size = New System.Drawing.Size(400, 597)
        Me.TreeView2.TabIndex = 0
        '
        'Options
        '
        Me.Options.Controls.Add(Me.CheckBox3)
        Me.Options.Controls.Add(Me.CheckBox2)
        Me.Options.Controls.Add(Me.CheckBox1)
        Me.Options.Controls.Add(Me.TextBox5)
        Me.Options.Controls.Add(Me.Label5)
        Me.Options.Controls.Add(Me.TextBox4)
        Me.Options.Controls.Add(Me.Label4)
        Me.Options.Location = New System.Drawing.Point(4, 22)
        Me.Options.Name = "Options"
        Me.Options.Size = New System.Drawing.Size(400, 597)
        Me.Options.TabIndex = 3
        Me.Options.Text = "Options"
        '
        'CheckBox3
        '
        Me.CheckBox3.Location = New System.Drawing.Point(40, 128)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(296, 24)
        Me.CheckBox3.TabIndex = 6
        Me.CheckBox3.Text = "Create a Project for the generated classes"
        '
        'CheckBox2
        '
        Me.CheckBox2.Location = New System.Drawing.Point(40, 112)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(304, 24)
        Me.CheckBox2.TabIndex = 5
        Me.CheckBox2.Text = "User ConnectionString Form application section"
        '
        'CheckBox1
        '
        Me.CheckBox1.Location = New System.Drawing.Point(40, 96)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(224, 24)
        Me.CheckBox1.TabIndex = 4
        Me.CheckBox1.Text = "Support Transaction"
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(136, 40)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.TabIndex = 3
        Me.TextBox5.Text = "TextBox5"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(32, 42)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 16)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Property Prefix"
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(136, 16)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.TabIndex = 1
        Me.TextBox4.Text = "TextBox4"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(32, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 16)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Class Prefix"
        '
        'SPOption
        '
        Me.SPOption.Controls.Add(Me.Button3)
        Me.SPOption.Location = New System.Drawing.Point(4, 22)
        Me.SPOption.Name = "SPOption"
        Me.SPOption.Size = New System.Drawing.Size(400, 597)
        Me.SPOption.TabIndex = 5
        Me.SPOption.Text = "Stored Procedures Options"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(296, 320)
        Me.Button3.Name = "Button3"
        Me.Button3.TabIndex = 0
        Me.Button3.Text = "close"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.DataBase)
        Me.TabControl1.Controls.Add(Me.Tables)
        Me.TabControl1.Controls.Add(Me.Views)
        Me.TabControl1.Controls.Add(Me.Options)
        Me.TabControl1.Controls.Add(Me.SPOption)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Left
        Me.TabControl1.ItemSize = New System.Drawing.Size(44, 18)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(408, 623)
        Me.TabControl1.TabIndex = 0
        '
        'Log
        '
        Me.Log.Location = New System.Drawing.Point(408, 24)
        Me.Log.Multiline = True
        Me.Log.Name = "Log"
        Me.Log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Log.Size = New System.Drawing.Size(464, 592)
        Me.Log.TabIndex = 1
        Me.Log.Text = ""
        '
        'Form1
        '
        Me.AutoScale = False
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(874, 623)
        Me.ControlBox = False
        Me.Controls.Add(Me.Log)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Form1"
        Me.Tables.ResumeLayout(False)
        Me.DataBase.ResumeLayout(False)
        Me.Views.ResumeLayout(False)
        Me.Options.ResumeLayout(False)
        Me.SPOption.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub Logit(ByVal Str As String)
        Log.Text += "Reading (" + Str + ") Table" + vbCrLf
    End Sub
    Public Sub Logit(ByVal Str As Integer)
        Log.Text += "Reading (" + Str.ToString + ") Table" + vbCrLf
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim tb As New clsTables
        AddHandler tb.ReadingTable, AddressOf Logit
        AddHandler tb.TotalTables, AddressOf Logit
        tb.FillTables(TextBox3.Text, TextBox1.Text, TextBox2.Text, TextBox6.Text)

        tb.StoredProcedures.SaveSP(TextBox8.Text)
        'Me.SuspendLayout()
        Dim otable As clsTable
        Dim ocolumn As clsColumn
        Dim RootNode As New TreeNode("AraOp")
        TablesView.Nodes.Add(RootNode)
        For Each otable In tb.Tables
            Dim crnode As New TreeNode(otable.TableName)
            RootNode.Nodes.Add(crnode)
            For Each ocolumn In otable.Columns
                crnode.Nodes.Add(New TreeNode(ocolumn.ColumnName))
            Next
        Next
        'Me.ResumeLayout()
        Dim clsgen As New ClassGenerator(TextBox7.Text, TextBox8.Text, tb.Tables)
        clsgen.ConstructClasses()
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Application.Exit()
    End Sub


#Region " Event Handlers "
    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Application.Exit()
    End Sub



#End Region

End Class
