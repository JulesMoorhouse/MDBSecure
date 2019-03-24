Imports Microsoft.Win32
Friend Class frmOpenMDB
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Friend Sub New()
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
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblAccessPath As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblAccessPath = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnOpen
        '
        Me.btnOpen.Location = New System.Drawing.Point(200, 240)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.TabIndex = 1
        Me.btnOpen.Text = "&Open"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(272, 16)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.TabIndex = 0
        Me.btnBrowse.Text = "&Browse"
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(16, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(344, 23)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "MDBSecure will try and open your database in MSAccess."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(16, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(368, 32)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "However, as MSAccess does not allow the database password to be passed in as a pa" & _
        "rameter, this process is a little bit more complicated."
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(16, 160)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(344, 40)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "MDBSecure will enter the password as if you had typed it yourself, if this proces" & _
        "s fails (due to the timing involved) you may have to enter the database password" & _
        " yourself!"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(280, 240)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "&Cancel"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(160, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "The path to MS Access is :"
        '
        'lblAccessPath
        '
        Me.lblAccessPath.Location = New System.Drawing.Point(40, 48)
        Me.lblAccessPath.Name = "lblAccessPath"
        Me.lblAccessPath.Size = New System.Drawing.Size(312, 32)
        Me.lblAccessPath.TabIndex = 8
        Me.lblAccessPath.Text = "MSAccess Not Found!"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.CheckFileExists = False
        Me.OpenFileDialog1.CheckPathExists = False
        Me.OpenFileDialog1.DefaultExt = "exe"
        Me.OpenFileDialog1.Filter = "MS Access|*.exe"
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Blue
        Me.Label5.Location = New System.Drawing.Point(16, 208)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(344, 23)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "You may need to minimize all other open applications!"
        '
        'frmOpenMDB
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 272)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label5, Me.lblAccessPath, Me.Label4, Me.btnCancel, Me.Label3, Me.Label2, Me.Label1, Me.btnBrowse, Me.btnOpen})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmOpenMDB"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Open Database"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim mstrCommParam As String
    Dim mstrDBPassword As String
    Dim lbooDoneOnce As Boolean = False
    Friend Property DBPassword() As String
        Get

            Return mstrDBPassword
        End Get
        Set(ByVal Value As String)
            mstrDBPassword = Value
        End Set
    End Property
    Friend Property CommParam() As String
        Get

            Return mstrCommParam
        End Get
        Set(ByVal Value As String)
            mstrCommParam = Value
        End Set
    End Property

    Function GetMSAccessPath() As String
        Dim ExeFilePath As String

        AddDebugComment("frmOpenMDB.GetMSAccessPath 1")

        Try
            Dim sKey As String
            Dim aKey As RegistryKey = Registry.ClassesRoot
            aKey = aKey.OpenSubKey("Access.Application\CurVer")
            Registry.ClassesRoot.GetValue("Access.Application\CurVer", "")
            sKey = aKey.GetValue("")

            AddDebugComment("frmOpenMDB.GetMSAccessPath 2")

            Dim sKey2 As String
            Dim aKey2 As RegistryKey = Registry.ClassesRoot
            aKey2 = aKey2.OpenSubKey(sKey & "\shell\New\command")
            Registry.ClassesRoot.GetValue(sKey & "\shell\New\command", "")
            sKey2 = aKey2.GetValue("")

            AddDebugComment("frmOpenMDB.GetMSAccessPath 3")

            ExeFilePath = Microsoft.VisualBasic.Left(sKey2, InStrGet(sKey2.ToUpper, "MSACCESS.EXE") + "MSACCESS.EXE".Length)

            ExeFilePath = ExeFilePath.Replace(ChrGet(34), "")

            AddDebugComment("frmOpenMDB.GetMSAccessPath 4")

        Catch
            AddDebugComment("frmOpenMDB.GetMSAccessPath Catch")
            ExeFilePath = ""
        End Try

        If System.IO.File.Exists(ExeFilePath) = True Then
            Return ExeFilePath
        End If

    End Function

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        AddDebugComment("frmOpenMDB.btnOpen_Click 1")

        If lbooDoneOnce = True Then
            Exit Sub
        Else
            lbooDoneOnce = True
        End If

        Try
            Me.SendToBack()

            Dim psInfo As New System.Diagnostics.ProcessStartInfo(lblAccessPath.Text, mstrCommParam)
            psInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized

            psInfo.CreateNoWindow = True
            AddDebugComment("frmOpenMDB.btnOpen_Click 2")

            Dim myProcess As Process = System.Diagnostics.Process.Start(psInfo)

            AddDebugComment("frmOpenMDB.btnOpen_Click 3")

            Dim lintArrInc As Integer
            For lintArrInc = 0 To 69999
                System.Windows.Forms.Application.DoEvents()
            Next lintArrInc

            SendKeys.SendWait(mstrDBPassword)
            SendKeys.SendWait("{Enter}")

            AddDebugComment("frmOpenMDB.btnOpen_Click 4")

            Me.Close()
        Catch
        End Try

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        AddDebugComment("frmOpenMDB.Form1_Load 1")

        'check config file first
        Dim lstrAccessExeFile As String = GetSetting("Access Exe File", "", InitalXMLConfig.XmlConfigType.AppSettings, "")
        If System.IO.File.Exists(lstrAccessExeFile) = True Then
            lblAccessPath.Text = lstrAccessExeFile
        Else
            lblAccessPath.Text = GetMSAccessPath()
        End If

        AddDebugComment("frmOpenMDB.Form1_Load 2")

        If lblAccessPath.Text = "" Then
            btnOpen.Enabled = False
        Else
            btnOpen.Enabled = True
        End If

        If IsAboveOrEqualWinXp() = True Then
            btnOpen.FlatStyle = FlatStyle.System
            btnCancel.FlatStyle = FlatStyle.System
            btnBrowse.FlatStyle = FlatStyle.System
        End If

        SetBackcolors()

        AddDebugComment("frmOpenMDB.Form1_Load 3")

    End Sub
    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click

        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName <> "" Then
            lblAccessPath.Text = OpenFileDialog1.FileName
        End If


        If System.IO.File.Exists(lblAccessPath.Text) = True Then
            btnOpen.Enabled = True
            SaveSetting("Access Exe File", lblAccessPath.Text, InitalXMLConfig.XmlConfigType.AppSettings, "")
        Else
            btnOpen.Enabled = False
        End If

    End Sub
    Private Sub SetBackcolors()

        btnOpen.BackColor = Color.FromArgb(0, btnOpen.BackColor)
        btnCancel.BackColor = Color.FromArgb(0, btnCancel.BackColor)
        btnBrowse.BackColor = Color.FromArgb(0, btnBrowse.BackColor)
        Label1.BackColor = Color.FromArgb(0, Label1.BackColor)
        Label2.BackColor = Color.FromArgb(0, Label2.BackColor)
        Label3.BackColor = Color.FromArgb(0, Label3.BackColor)
        Label4.BackColor = Color.FromArgb(0, Label4.BackColor)
        Label5.BackColor = Color.FromArgb(0, Label5.BackColor)
        lblAccessPath.BackColor = Color.FromArgb(0, lblAccessPath.BackColor)

    End Sub
    Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)

        Dim PaintBack As New UIStyle.Painting
        PaintBack.PaintBackground(pevent, Me)

    End Sub
End Class