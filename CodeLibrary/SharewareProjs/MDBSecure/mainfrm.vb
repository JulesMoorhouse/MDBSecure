Imports ADOX
Imports System
Imports System.Windows.Forms
Imports System.IO
<Obfuscate()> Public Class mainfrm
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
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents chkCopyTables As System.Windows.Forms.CheckBox
    Friend WithEvents lblSourceDB As System.Windows.Forms.Label
    Friend WithEvents btnSelectSource As System.Windows.Forms.Button
    Friend WithEvents lstSourceTables As System.Windows.Forms.ListBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtNewPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtSuperUser As System.Windows.Forms.TextBox
    Friend WithEvents txtSuperGroup As System.Windows.Forms.TextBox
    Friend WithEvents txtAdminPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpCFU As System.Windows.Forms.MenuItem
    Friend WithEvents cboDBVersion As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents grbSourceDB As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtDatabasePassword As System.Windows.Forms.TextBox
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpLicense As System.Windows.Forms.MenuItem
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents btnSelectAll As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnClearAll As System.Windows.Forms.Button
    Friend WithEvents mnuHelpSupport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents btnOpenDB As System.Windows.Forms.Button
    Friend WithEvents chkDynHelp As System.Windows.Forms.CheckBox
    Friend WithEvents AxWebbrowser1 As WinOnly.WebOCHostCtrl
    Friend WithEvents mnuHelpInstallPack As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpEnterCode As System.Windows.Forms.MenuItem
    Friend WithEvents btnAddJetString As System.Windows.Forms.Button
    Friend WithEvents menuEnhancer As WinOnly.EnhancedMenu
    Friend WithEvents rdoCreateNew As System.Windows.Forms.RadioButton
    Friend WithEvents rdoOpenSecured As System.Windows.Forms.RadioButton
    Friend WithEvents grpUserChoices As System.Windows.Forms.GroupBox
    Friend WithEvents mnuHelpReportProblem As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(mainfrm))
        Me.btnCreate = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.chkCopyTables = New System.Windows.Forms.CheckBox()
        Me.lblSourceDB = New System.Windows.Forms.Label()
        Me.btnSelectSource = New System.Windows.Forms.Button()
        Me.lstSourceTables = New System.Windows.Forms.ListBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txtNewPassword = New System.Windows.Forms.TextBox()
        Me.txtSuperUser = New System.Windows.Forms.TextBox()
        Me.txtSuperGroup = New System.Windows.Forms.TextBox()
        Me.txtAdminPassword = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtDatabasePassword = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboDBVersion = New System.Windows.Forms.ComboBox()
        Me.grbSourceDB = New System.Windows.Forms.GroupBox()
        Me.btnClearAll = New System.Windows.Forms.Button()
        Me.btnSelectAll = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.mnuHelp = New System.Windows.Forms.MenuItem()
        Me.mnuHelpLicense = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuHelpCFU = New System.Windows.Forms.MenuItem()
        Me.mnuHelpInstallPack = New System.Windows.Forms.MenuItem()
        Me.mnuHelpSupport = New System.Windows.Forms.MenuItem()
        Me.mnuHelpReportProblem = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuHelpEnterCode = New System.Windows.Forms.MenuItem()
        Me.mnuHelpAbout = New System.Windows.Forms.MenuItem()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btnOpenDB = New System.Windows.Forms.Button()
        Me.AxWebbrowser1 = New WinOnly.WebOCHostCtrl()
        Me.chkDynHelp = New System.Windows.Forms.CheckBox()
        Me.btnAddJetString = New System.Windows.Forms.Button()
        Me.menuEnhancer = New WinOnly.EnhancedMenu(Me.components)
        Me.grpUserChoices = New System.Windows.Forms.GroupBox()
        Me.rdoOpenSecured = New System.Windows.Forms.RadioButton()
        Me.rdoCreateNew = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.grbSourceDB.SuspendLayout()
        Me.grpUserChoices.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(121, 428)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(104, 23)
        Me.btnCreate.TabIndex = 13
        Me.btnCreate.Text = "&Create Database"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.FileName = "doc1"
        '
        'chkCopyTables
        '
        Me.chkCopyTables.Location = New System.Drawing.Point(17, 236)
        Me.chkCopyTables.Name = "chkCopyTables"
        Me.chkCopyTables.Size = New System.Drawing.Size(344, 24)
        Me.chkCopyTables.TabIndex = 8
        Me.chkCopyTables.Text = "Basic copy of tables from another mdb (not password protected)"
        '
        'lblSourceDB
        '
        Me.lblSourceDB.Location = New System.Drawing.Point(8, 24)
        Me.lblSourceDB.Name = "lblSourceDB"
        Me.lblSourceDB.Size = New System.Drawing.Size(272, 23)
        Me.lblSourceDB.TabIndex = 2
        Me.lblSourceDB.Text = "SourceDB"
        '
        'btnSelectSource
        '
        Me.btnSelectSource.Enabled = False
        Me.btnSelectSource.Font = New System.Drawing.Font("Arial Black", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelectSource.Location = New System.Drawing.Point(296, 24)
        Me.btnSelectSource.Name = "btnSelectSource"
        Me.btnSelectSource.Size = New System.Drawing.Size(24, 24)
        Me.btnSelectSource.TabIndex = 9
        Me.btnSelectSource.Text = "..."
        '
        'lstSourceTables
        '
        Me.lstSourceTables.BackColor = System.Drawing.SystemColors.Control
        Me.lstSourceTables.Enabled = False
        Me.lstSourceTables.Location = New System.Drawing.Point(8, 48)
        Me.lstSourceTables.Name = "lstSourceTables"
        Me.lstSourceTables.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstSourceTables.Size = New System.Drawing.Size(240, 82)
        Me.lstSourceTables.TabIndex = 10
        '
        'txtNewPassword
        '
        Me.txtNewPassword.Location = New System.Drawing.Point(168, 96)
        Me.txtNewPassword.MaxLength = 20
        Me.txtNewPassword.Name = "txtNewPassword"
        Me.txtNewPassword.Size = New System.Drawing.Size(104, 20)
        Me.txtNewPassword.TabIndex = 6
        Me.txtNewPassword.Text = ""
        '
        'txtSuperUser
        '
        Me.txtSuperUser.Location = New System.Drawing.Point(168, 72)
        Me.txtSuperUser.MaxLength = 20
        Me.txtSuperUser.Name = "txtSuperUser"
        Me.txtSuperUser.Size = New System.Drawing.Size(104, 20)
        Me.txtSuperUser.TabIndex = 5
        Me.txtSuperUser.Text = ""
        '
        'txtSuperGroup
        '
        Me.txtSuperGroup.Location = New System.Drawing.Point(168, 120)
        Me.txtSuperGroup.MaxLength = 20
        Me.txtSuperGroup.Name = "txtSuperGroup"
        Me.txtSuperGroup.Size = New System.Drawing.Size(104, 20)
        Me.txtSuperGroup.TabIndex = 7
        Me.txtSuperGroup.Text = ""
        '
        'txtAdminPassword
        '
        Me.txtAdminPassword.Location = New System.Drawing.Point(168, 24)
        Me.txtAdminPassword.MaxLength = 20
        Me.txtAdminPassword.Name = "txtAdminPassword"
        Me.txtAdminPassword.Size = New System.Drawing.Size(104, 20)
        Me.txtAdminPassword.TabIndex = 3
        Me.txtAdminPassword.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 23)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Super User Password :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 23)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Super Username :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(48, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 23)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Super User Group :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(32, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 23)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Admin User Password :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label7, Me.txtDatabasePassword, Me.Label1, Me.cboDBVersion, Me.Label4, Me.Label3, Me.txtNewPassword, Me.Label5, Me.Label2, Me.txtAdminPassword, Me.txtSuperGroup, Me.txtSuperUser})
        Me.GroupBox1.Location = New System.Drawing.Point(17, 68)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(328, 152)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Security Atrributes and Settings"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(32, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(128, 23)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Database Password :"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'txtDatabasePassword
        '
        Me.txtDatabasePassword.Location = New System.Drawing.Point(168, 48)
        Me.txtDatabasePassword.MaxLength = 20
        Me.txtDatabasePassword.Name = "txtDatabasePassword"
        Me.txtDatabasePassword.Size = New System.Drawing.Size(104, 20)
        Me.txtDatabasePassword.TabIndex = 4
        Me.txtDatabasePassword.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(280, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 23)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Database Version :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label1.Visible = False
        '
        'cboDBVersion
        '
        Me.cboDBVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDBVersion.Items.AddRange(New Object() {"1.0 ", "1.1", "2.0", "95 / 97", "2000 / 2002"})
        Me.cboDBVersion.Location = New System.Drawing.Point(288, 112)
        Me.cboDBVersion.Name = "cboDBVersion"
        Me.cboDBVersion.Size = New System.Drawing.Size(24, 21)
        Me.cboDBVersion.TabIndex = 5
        Me.cboDBVersion.Visible = False
        '
        'grbSourceDB
        '
        Me.grbSourceDB.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnClearAll, Me.lblSourceDB, Me.lstSourceTables, Me.btnSelectSource, Me.btnSelectAll, Me.Label6})
        Me.grbSourceDB.Enabled = False
        Me.grbSourceDB.Location = New System.Drawing.Point(17, 260)
        Me.grbSourceDB.Name = "grbSourceDB"
        Me.grbSourceDB.Size = New System.Drawing.Size(328, 160)
        Me.grbSourceDB.TabIndex = 7
        Me.grbSourceDB.TabStop = False
        Me.grbSourceDB.Text = "Source Database"
        '
        'btnClearAll
        '
        Me.btnClearAll.Location = New System.Drawing.Point(256, 88)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(64, 23)
        Me.btnClearAll.TabIndex = 12
        Me.btnClearAll.Text = "C&lear All"
        '
        'btnSelectAll
        '
        Me.btnSelectAll.Location = New System.Drawing.Point(256, 56)
        Me.btnSelectAll.Name = "btnSelectAll"
        Me.btnSelectAll.Size = New System.Drawing.Size(64, 23)
        Me.btnSelectAll.TabIndex = 11
        Me.btnSelectAll.Text = "Select &All"
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Blue
        Me.Label6.Location = New System.Drawing.Point(8, 136)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(240, 16)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Select which tables you want to copy"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelp})
        '
        'mnuHelp
        '
        Me.menuEnhancer.SetImageIndex(Me.mnuHelp, -1)
        Me.mnuHelp.Index = 0
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpLicense, Me.MenuItem2, Me.mnuHelpCFU, Me.mnuHelpInstallPack, Me.mnuHelpSupport, Me.mnuHelpReportProblem, Me.MenuItem1, Me.mnuHelpEnterCode, Me.mnuHelpAbout})
        Me.mnuHelp.OwnerDraw = True
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpLicense
        '
        Me.menuEnhancer.SetImageIndex(Me.mnuHelpLicense, -1)
        Me.mnuHelpLicense.Index = 0
        Me.mnuHelpLicense.OwnerDraw = True
        Me.mnuHelpLicense.Text = "&License Agreement"
        '
        'MenuItem2
        '
        Me.menuEnhancer.SetImageIndex(Me.MenuItem2, -1)
        Me.MenuItem2.Index = 1
        Me.MenuItem2.OwnerDraw = True
        Me.MenuItem2.Text = "-"
        '
        'mnuHelpCFU
        '
        Me.menuEnhancer.SetImageIndex(Me.mnuHelpCFU, -1)
        Me.mnuHelpCFU.Index = 2
        Me.mnuHelpCFU.OwnerDraw = True
        Me.mnuHelpCFU.Text = "&Check for updates"
        '
        'mnuHelpInstallPack
        '
        Me.menuEnhancer.SetImageIndex(Me.mnuHelpInstallPack, -1)
        Me.mnuHelpInstallPack.Index = 3
        Me.mnuHelpInstallPack.OwnerDraw = True
        Me.mnuHelpInstallPack.Text = "&Install add-on"
        '
        'mnuHelpSupport
        '
        Me.menuEnhancer.SetImageIndex(Me.mnuHelpSupport, -1)
        Me.mnuHelpSupport.Index = 4
        Me.mnuHelpSupport.OwnerDraw = True
        Me.mnuHelpSupport.Text = "S&upport Info"
        '
        'mnuHelpReportProblem
        '
        Me.menuEnhancer.SetImageIndex(Me.mnuHelpReportProblem, -1)
        Me.mnuHelpReportProblem.Index = 5
        Me.mnuHelpReportProblem.OwnerDraw = True
        Me.mnuHelpReportProblem.Text = "&Report Problem"
        '
        'MenuItem1
        '
        Me.menuEnhancer.SetImageIndex(Me.MenuItem1, -1)
        Me.MenuItem1.Index = 6
        Me.MenuItem1.OwnerDraw = True
        Me.MenuItem1.Text = "-"
        '
        'mnuHelpEnterCode
        '
        Me.menuEnhancer.SetImageIndex(Me.mnuHelpEnterCode, -1)
        Me.mnuHelpEnterCode.Index = 7
        Me.mnuHelpEnterCode.OwnerDraw = True
        Me.mnuHelpEnterCode.Text = "&Enter License Code"
        '
        'mnuHelpAbout
        '
        Me.menuEnhancer.SetImageIndex(Me.mnuHelpAbout, -1)
        Me.mnuHelpAbout.Index = 8
        Me.mnuHelpAbout.OwnerDraw = True
        Me.mnuHelpAbout.Text = "&About"
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(9, 460)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(336, 40)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "This program will make your databases MORE secure.  If you have any ideas, sugges" & _
        "tions or comments to make about this program, please visit our forum on our webs" & _
        "ite."
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Location = New System.Drawing.Point(9, 508)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(344, 16)
        Me.LinkLabel1.TabIndex = 17
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "http://www.example.com"
        Me.LinkLabel1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblProgress
        '
        Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblProgress.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.Location = New System.Drawing.Point(321, 428)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(24, 24)
        Me.lblProgress.TabIndex = 17
        Me.lblProgress.Visible = False
        '
        'Label9
        '
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label9.Enabled = False
        Me.Label9.ForeColor = System.Drawing.Color.Blue
        Me.Label9.Location = New System.Drawing.Point(0, 534)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(610, 16)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "MDBSecure™   Developed by Mindwarp Consultancy Ltd © " & gYear
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnOpenDB
        '
        Me.btnOpenDB.Location = New System.Drawing.Point(17, 428)
        Me.btnOpenDB.Name = "btnOpenDB"
        Me.btnOpenDB.Size = New System.Drawing.Size(96, 23)
        Me.btnOpenDB.TabIndex = 15
        Me.btnOpenDB.Text = "&Open Database"
        '
        'AxWebbrowser1
        '
        Me.AxWebbrowser1.BrowserContextMenu = False
        Me.AxWebbrowser1.Location = New System.Drawing.Point(360, 8)
        Me.AxWebbrowser1.Name = "AxWebbrowser1"
        Me.AxWebbrowser1.Size = New System.Drawing.Size(232, 495)
        Me.AxWebbrowser1.TabIndex = 0
        '
        'chkDynHelp
        '
        Me.chkDynHelp.Checked = True
        Me.chkDynHelp.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDynHelp.Location = New System.Drawing.Point(361, 510)
        Me.chkDynHelp.Name = "chkDynHelp"
        Me.chkDynHelp.Size = New System.Drawing.Size(144, 16)
        Me.chkDynHelp.TabIndex = 16
        Me.chkDynHelp.Text = "Show help dynamically"
        '
        'btnAddJetString
        '
        Me.btnAddJetString.Location = New System.Drawing.Point(233, 428)
        Me.btnAddJetString.Name = "btnAddJetString"
        Me.btnAddJetString.Size = New System.Drawing.Size(80, 23)
        Me.btnAddJetString.TabIndex = 14
        Me.btnAddJetString.Text = "&Jet String"
        '
        'menuEnhancer
        '
        Me.menuEnhancer.ColorsControl = System.Drawing.SystemColors.Control
        Me.menuEnhancer.ColorsHighlight = System.Drawing.SystemColors.Highlight
        Me.menuEnhancer.ColorsWindow = System.Drawing.SystemColors.Window
        Me.menuEnhancer.ImageListMarks = Nothing
        '
        'grpUserChoices
        '
        Me.grpUserChoices.Controls.AddRange(New System.Windows.Forms.Control() {Me.rdoOpenSecured, Me.rdoCreateNew})
        Me.grpUserChoices.Location = New System.Drawing.Point(16, 13)
        Me.grpUserChoices.Name = "grpUserChoices"
        Me.grpUserChoices.Size = New System.Drawing.Size(328, 44)
        Me.grpUserChoices.TabIndex = 20
        Me.grpUserChoices.TabStop = False
        Me.grpUserChoices.Text = "User Choices"
        '
        'rdoOpenSecured
        '
        Me.rdoOpenSecured.Location = New System.Drawing.Point(174, 13)
        Me.rdoOpenSecured.Name = "rdoOpenSecured"
        Me.rdoOpenSecured.Size = New System.Drawing.Size(147, 24)
        Me.rdoOpenSecured.TabIndex = 2
        Me.rdoOpenSecured.Text = "Open Secured Database"
        '
        'rdoCreateNew
        '
        Me.rdoCreateNew.Checked = True
        Me.rdoCreateNew.Location = New System.Drawing.Point(11, 12)
        Me.rdoCreateNew.Name = "rdoCreateNew"
        Me.rdoCreateNew.Size = New System.Drawing.Size(134, 24)
        Me.rdoCreateNew.TabIndex = 1
        Me.rdoCreateNew.TabStop = True
        Me.rdoCreateNew.Text = "Create New Database"
        '
        'mainfrm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(610, 550)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.grpUserChoices, Me.btnAddJetString, Me.chkDynHelp, Me.AxWebbrowser1, Me.btnOpenDB, Me.Label9, Me.lblProgress, Me.LinkLabel1, Me.Label8, Me.GroupBox1, Me.chkCopyTables, Me.grbSourceDB, Me.btnCreate})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.Name = "mainfrm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MDBSecure"
        Me.GroupBox1.ResumeLayout(False)
        Me.grbSourceDB.ResumeLayout(False)
        Me.grpUserChoices.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim lbooSplashShown As Boolean = False
    Dim Overide As Boolean = False
    Dim lbooLoadDataNow As Boolean = True
    Dim lstrCFUCode As String
    Enum WebHelpPages
        welcome
        AdminUser
        DatabasePassword
        SuperUser
        SuperUserGroup
        Append
        Open
        Create
        JetString
    End Enum

    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click

        AddDebugComment("mainfrm.btnCreate_Click - Start")

        If CheckFields() = False Then
            Exit Sub
        End If

        Overide = True

        btnCreate.Enabled = False
        btnSelectAll.Enabled = False
        btnSelectSource.Enabled = False
        btnOpenDB.Enabled = False
        btnAddJetString.Enabled = False

        Dim NewDB As ADOX.Catalog = New ADOX.Catalog
        Dim lstrSystemDB As String
        Dim lstrPreDB As String
        Dim lstrNewConnString As String
        Dim lcnn1 As New ADODB.Connection

        Dim gconstrConnectionProvider As String
        Dim gconstrConnectionJustProvider As String

        Dim lstrDetails(3) As String
        lstrDetails(0) = "33IHGPFEDPIHGPFEDP"
        lstrDetails(1) = "jHK;DQV;UH=<mQW;HKG;KTF;LpUKFLk?nuv;xLF@LJ?PpUF;ZknKuvx;@vYKFYX;YGU;Zj"
        lstrDetails(2) = "YKGGC;KHV;=GHPORV;GGY;ACDP?eGKUHZ<qv=<qje;GUH;?jYKGGC;KHV;=wuKtJI@tuG<"
        lstrDetails(3) = "vDD@yw?<vYF;YZg?KEH;WU=<mK"

        gconstrConnectionProvider = Decrypt("", "", flame.Encops.EncryptString, lstrDetails)
        ReDim lstrDetails(0)

        ReDim lstrDetails(1)
        lstrDetails(0) = "33IHGPFEDPIHGPFEDP"
        lstrDetails(1) = "jHK;DQV;UH=<mQW;HKG;KTF;LpUKFLk?nuv;xLF@LJ?PvYF;YZg?KEH;WU=<mK"

        gconstrConnectionJustProvider = Decrypt("", "", flame.Encops.EncryptString, lstrDetails)
        ReDim lstrDetails(0)

        Busy(Me, True)
        System.Windows.Forms.Application.DoEvents()

        lstrPreDB = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\~predb.tmp"
        Try : System.IO.File.Delete(lstrPreDB) : Catch : End Try

        With SaveFileDialog1
            .CheckPathExists = True
            .AddExtension = True
            .FileName = "Project"
            .Filter = "Database file (*.mdb)|*.mdb"
            .DefaultExt = "mdb"
        End With

        AddDebugComment("mainfrm.btnCreate_Click - About to Show SFD")

        Dim eh As CustomExceptionHandler = New CustomExceptionHandler
        Try

            If SaveFileDialog1.ShowDialog() <> DialogResult.Cancel And SaveFileDialog1.FileName <> "" Then
                lstrSystemDB = LeftGet(SaveFileDialog1.FileName, SaveFileDialog1.FileName.Length - 1) & "w"

                AddDebugComment("mainfrm.btnCreate_Click - About to Delete Old Files")

                System.Windows.Forms.Application.DoEvents()

                Try : System.IO.File.Delete(lstrSystemDB) : Catch : End Try
                Try : System.IO.File.Delete(SaveFileDialog1.FileName) : Catch : End Try

                AddDebugComment("mainfrm.btnCreate_Click - B1a Command Build")

                Dim lstrComm(0) As String
                Dim lintCLArrInc1 As Integer
                Dim lstrCommandLine1 As String = "DBNWMAKE@" & _
                     lstrPreDB & "@" & lstrSystemDB & "@Admin@none@" & _
                    txtNewPassword.Text & "@" & txtSuperUser.Text & "@" & txtSuperGroup.Text & "@"

                Encrypt(lstrCommandLine1, "dsdfsffsfsaik", lstrComm)
                lstrCommandLine1 = ""
                For lintCLArrInc1 = 0 To lstrComm.GetUpperBound(0)
                    lstrCommandLine1 = lstrCommandLine1 & lstrComm(lintCLArrInc1) & "X3X"
                Next lintCLArrInc1

                AddDebugComment("mainfrm.btnCreate_Click - About to Launch Beside01a")

                Dim Beside01FirstCall As New AppRun

                If Beside01FirstCall.Run(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\", _
                    "Beside01.exe", lstrCommandLine1, 40, False) = False Then
                    AddDebugComment("mainfrm.btnCreate_Click - Beside01a Kill")
                    Try : System.IO.File.Delete(lstrPreDB) : Catch : End Try
                    Busy(Me, False)

                    MessageBox.Show("Unable to create file! This may be because the program " & Environment.NewLine() & _
                                    "wasn't able to run fast enough on this computer! Error code B11", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    btnCreate.Enabled = True
                    btnSelectAll.Enabled = True
                    btnSelectSource.Enabled = True
                    btnOpenDB.Enabled = True
                    btnAddJetString.Enabled = True
                    Overide = False
                    Exit Sub
                End If

                System.Windows.Forms.Application.DoEvents()

                '*' Things to do
                '*' ---------------
                '*' create new db - to create IPUser ownership

                Dim lstrComm2(0) As String
                Dim lintCLArrInc2 As Integer
                Dim lstrCommandLine2 As String = "DBMAKE@" & _
                    SaveFileDialog1.FileName & "@" & lstrSystemDB & "@" & txtSuperUser.Text & "@" & txtNewPassword.Text & "@NONE@NONE@NONE@"
                AddDebugComment("mainfrm.btnCreate_Click - B1b Command Build")
                Encrypt(lstrCommandLine2, "dsdfsffsfsaik", lstrComm2)
                lstrCommandLine2 = ""
                For lintCLArrInc2 = 0 To lstrComm2.GetUpperBound(0)
                    lstrCommandLine2 = lstrCommandLine2 & lstrComm2(lintCLArrInc2) & "X3X"
                Next lintCLArrInc2

                AddDebugComment("mainfrm.btnCreate_Click - About to Launch Beside01b")

                Dim Beside01SecondCall As New AppRun
                If Beside01SecondCall.Run(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\", _
                    "Beside01.exe", lstrCommandLine2, 40, False) = False Then
                    AddDebugComment("mainfrm.btnCreate_Click - Beside01b Kill")
                    Try : System.IO.File.Delete(lstrPreDB) : Catch : End Try
                    Busy(Me, False)
                    MessageBox.Show("Unable to create file! This may be because the program " & Environment.NewLine() & _
                              "wasn't able to run fast enough on this computer! Error Code B12", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    btnCreate.Enabled = True
                    btnSelectAll.Enabled = True
                    btnSelectSource.Enabled = True
                    btnOpenDB.Enabled = True
                    btnAddJetString.Enabled = True
                    Overide = False
                    Exit Sub
                End If

                System.Windows.Forms.Application.DoEvents()

                AddDebugComment("mainfrm.btnCreate_Click - DB Trans 1")

                '*' open new db
                lcnn1 = New ADODB.Connection

                lstrNewConnString = gconstrConnectionJustProvider & SaveFileDialog1.FileName & _
                    ";User ID=" & txtSuperUser.Text & ";Password=" & txtNewPassword.Text & ";Jet OLEDB:System database=" & lstrSystemDB

                AddDebugComment("mainfrm.btnCreate_Click - Chinese Err " & lstrNewConnString)

                lcnn1.Open(lstrNewConnString)

                AddDebugComment("mainfrm.btnCreate_Click - DB Trans 1a")

                NewDB.ActiveConnection = lcnn1

                AddDebugComment("mainfrm.btnCreate_Click - DB Trans 1b")

                NewDB.Users("Admin").ChangePassword("", txtAdminPassword.Text)
                ReDim lstrDetails(0)

                NewDB = Nothing
                lcnn1.Close()

                AddDebugComment("mainfrm.btnCreate_Click - DB Trans 2")

                Dim lcnn2 As New ADODB.Connection
                lcnn2 = New ADODB.Connection

                ReDim lstrDetails(2)
                '#### E4 - Part Connection String / IP User password

                lcnn2.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
                lcnn2.Open(lstrNewConnString)

                '#### E5 - Alter SQL DB Password
                Dim lstrSQL As String = "Alter Database password [" & txtDatabasePassword.Text & "] NULL"
                gstrLastSQL = lstrSQL
                lcnn2.Execute(lstrSQL)

                lstrNewConnString = lstrNewConnString & ";Jet OLEDB:Database Password=" & txtDatabasePassword.Text

                AddDebugComment("mainfrm.btnCreate_Click - Point A")

                ''#### E6 - Alter SQL IPUser Password

                Try : System.IO.File.Delete(lstrPreDB) : Catch : End Try

                lcnn2.Close()

                If chkCopyTables.Checked = True And lstSourceTables.Items.Count >= 1 Then
                    lblProgress.Visible = True
                    AddDebugComment("mainfrm.btnCreate_Click - Point B")
                    If AppendTables(lblSourceDB.Text, lstrNewConnString) = False Then
                        Exit Sub
                    End If
                End If

                lblProgress.BackColor = System.Drawing.SystemColors.Control
                lblProgress.Visible = False

                btnCreate.Enabled = True
                btnSelectAll.Enabled = True
                btnSelectSource.Enabled = True
                btnOpenDB.Enabled = True
                btnAddJetString.Enabled = True
                Overide = False
                Busy(Me, False)

                MessageBox.Show("Database created!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            btnCreate.Enabled = True
            btnSelectAll.Enabled = True
            btnSelectSource.Enabled = True

            AddDebugComment("mainfrm.btnCreate_Click - Point C")

        Catch ex As Exception
            AddDebugComment("<Font color=Red>MSG:" & ex.ToString & "</font>")
            eh.OnThreadException(Nothing, Nothing)

        End Try

        rdoOpenSecured.Checked = True
        rdoUserChoice_Click(Nothing, Nothing)
        btnOpenDB_Enter(Nothing, Nothing)

        Busy(Me, False)

    End Sub
    Private Sub SourceSchema(ByVal pstrSourceDB As String)

        AddDebugComment("mainfrm.SourceSchema")

        Dim lngIndex As Long
        Dim cnnS As New ADODB.Connection
        Dim catS As New ADOX.Catalog, tblS As New ADOX.Table

        Try
            cnnS = New ADODB.Connection
            cnnS.Open("Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source= " & pstrSourceDB & ";")
            catS.ActiveConnection = cnnS

            Dim tbl As New ADOX.Table

            For Each tbl In catS.Tables
                If tbl.Type.ToUpper = "TABLE" Then
                    lstSourceTables.Items.Add(tbl.Name)
                End If
            Next

            Dim db As Object
            db = Microsoft.VisualBasic.CreateObject("DAO.Database")

            Dim objReportDocuments As Object
            objReportDocuments = Microsoft.VisualBasic.CreateObject("DAO.Documents")

            Dim objReport As Object
            objReport = Microsoft.VisualBasic.CreateObject("DAO.Document")

            Dim mydbengine As Object
            mydbengine = Microsoft.VisualBasic.CreateObject("DAO.DBEngine")

            db = mydbengine("c:\Projects\Firm Reporting\Profile 2000.mdb")

            objReportDocuments = db.Containers("Reports").Documents

            For Each objReport In objReportDocuments
                lstSourceTables.Items.Add(objReport.Name)
            Next objReport

            db.Close()
            db = Nothing
            objReportDocuments = Nothing
            objReport = Nothing

        Catch ex As Exception
            AddDebugComment("mainfrm.SourceSchema " & ex.ToString)

            MessageBox.Show(ex.Message, NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Try
                tblS = Nothing
                catS = Nothing
                cnnS.Close()
                cnnS = Nothing
            Catch e As Exception
            End Try
        End Try
    End Sub
    Private Sub btnSelectSource_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectSource.Click

        AddDebugComment("mainfrm.btnSelectSource_Click")

        lstSourceTables.Items.Clear()
        With OpenFileDialog1
            .Filter = "Database file (*.mdb)|*.mdb"
            .DefaultExt = "mdb"
            .ShowDialog()
            If .FileName <> "" Then
                System.Windows.Forms.Application.DoEvents()
                SourceSchema(.FileName)
                lblSourceDB.Text = .FileName
            End If
        End With

    End Sub
    Private Function AppendTables(ByVal pstrSourceDB As String, ByVal lstrDestConnString As String) As Boolean

        AppendTables = True

        AddDebugComment("mainfrm.AppendTables - Start")

        Dim lngIndex As Long
        Dim msgReturn As DialogResult

        Dim cnnS As New ADODB.Connection
        Dim catS As New ADOX.Catalog, tblS As New ADOX.Table
        Dim cnnD As New ADODB.Connection
        Dim catD As New ADOX.Catalog, tblD As New ADOX.Table

        Dim lintProgressCtr As Integer

        cnnS = New ADODB.Connection
        cnnD = New ADODB.Connection
        Try
            cnnS.Open("Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source= " & pstrSourceDB & ";")
            catS.ActiveConnection = cnnS
        Catch ex As Exception
            MessageBox.Show("There was a problem opening your source database!" & CR() & _
                "Please ensure that it is unprotected and compatible!" & CR() & CR() & _
                ex.Message, NameMe("Import Database Access"), MessageBoxButtons.OK, MessageBoxIcon.Error)

            AppendTables = False
            Exit Function
        End Try

        cnnD.Open(lstrDestConnString)
        catD.ActiveConnection = cnnD

        ' Copy datas from Source table to Destination table
        AddDebugComment("mainfrm.AppendTables - Point A")

        Dim Item As Object

        'create tables
        Try
            For Each Item In lstSourceTables.SelectedItems
                Dim lstrSelectedTableName As String = (CType(Item, String))
                Dim oTable As ADOX.Table
                oTable = New ADOX.Table
                oTable.Name = lstrSelectedTableName
                catD.Tables.Append(oTable)
                lintProgressCtr += 1 : If lintProgressCtr > 30 Then : Progress() : lintProgressCtr = 0 : End If
            Next Item
        Catch ex As Exception
        End Try

        AddDebugComment("mainfrm.AppendTables - Point B")

        'Append columns to table
        Dim lbooTableAlreadyAppended As Boolean = False
        For Each Item In lstSourceTables.SelectedItems
            Dim lstrSelectedTableName As String = (CType(Item, String))
            tblD.Name = lstrSelectedTableName

            tblS = catS.Tables(lstrSelectedTableName)
            With tblS.Columns
                Dim x As ADOX.Column
                For Each x In tblS.Columns

                    Dim oColumn As ADOX.Column
                    oColumn = New ADOX.Column
                    With oColumn
                        .Name = x.Name
                        .Type = x.Type
                        .ParentCatalog = catD     ' Must set before setting properties
                        .DefinedSize = x.DefinedSize
                        .NumericScale = x.NumericScale
                        .Precision = x.Precision

                        'These 3 lines work, but set AZL to yes when should be no
                        If x.Attributes <> ColumnAttributesEnum.adColNullable Then
                            .Properties("Jet OLEDB:Allow Zero Length").Value = True
                        End If

                        If x.Attributes = ColumnAttributesEnum.adColNullable Then
                            .Attributes = ColumnAttributesEnum.adColNullable
                        End If

                        If x.Properties("Autoincrement").Value = True Then
                            .Properties("Autoincrement").Value = True
                        End If

                    End With
                    catD.Tables(lstrSelectedTableName).Columns.Append(oColumn)

                    lintProgressCtr += 1 : If lintProgressCtr > 30 Then : Progress() : lintProgressCtr = 0 : End If
                Next
            End With
            catD.Tables.Refresh()
        Next Item

        AddDebugComment("mainfrm.AppendTables - Point C")

        'append data
        For Each Item In lstSourceTables.SelectedItems
            Dim lstrSelectedTableName As String = (CType(Item, String))
            tblD.Name = lstrSelectedTableName

            tblS = catS.Tables(lstrSelectedTableName)

            Dim rstS As New ADODB.Recordset
            Dim rstD As New ADODB.Recordset
            rstS.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            Dim lstrSQL As String = "SELECT * FROM [" & tblD.Name & "]"
            gstrLastSQL = lstrSQL

            rstS.Open(lstrSQL, cnnS, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)

            rstD.CursorLocation = ADODB.CursorLocationEnum.adUseClient

            Dim lstrSQL2 As String = "SELECT * FROM [" & tblD.Name & "]"
            gstrLastSQL = lstrSQL2

            rstD.Open(lstrSQL2, cnnD, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)

            ' Add all datas into the destination table
            With rstS
                While Not (.EOF Or .BOF)
                    rstD.AddNew()
                    Dim y As ADODB.Field
                    For Each y In rstS.Fields
                        rstD.Fields(y.Name).Value = rstS.Fields(y.Name).Value

                        lintProgressCtr += 1 : If lintProgressCtr > 30 Then : Progress() : lintProgressCtr = 0 : End If
                    Next
                    rstD.UpdateBatch(ADODB.AffectEnum.adAffectCurrent)
                    .MoveNext()
                End While
            End With
            rstS.Close()
            rstS = Nothing
            rstD.Close()
            rstD = Nothing
        Next

        AddDebugComment("mainfrm.AppendTables - Point D")

        ' Release all the objects
        tblS = Nothing
        catS = Nothing
        cnnS.Close()
        cnnS = Nothing

        tblD = Nothing
        catD = Nothing
        cnnD.Close()
        cnnD = Nothing
        AddDebugComment("mainfrm.AppendTables - Point E")
    End Function
    Private Sub Encrypt(ByVal pstrInput As String, ByVal pstrView As String, Optional ByRef pstrInp() As String = Nothing)
        Dim strHead As String
        Dim strT As String
        Dim strA As String
        Dim cphX As New Fix
        Dim lngN As Long

        strA = pstrInput & CR()
        strT = Hash(Date.Today & CStr(Microsoft.VisualBasic.Timer))
        strHead = "33" & strT & Hash(strT & pstrView)

        cphX.KeyString = strHead
        cphX.Text = strA
        cphX.DoXor()
        cphX.Stretch()
        strA = cphX.Text

        Dim lintArrInc As Integer
        ReDim pstrInp(0)
        pstrInp(0) = strHead
        lngN = 1
        Do
            lintArrInc = lintArrInc + 1
            ReDim Preserve pstrInp(lintArrInc)
            pstrInp(lintArrInc) = MidGet(strA, lngN, 70)
            lngN = lngN + 70
        Loop Until lngN > (strA).Length

    End Sub
    Private Function Hash(ByVal strA As String) As String
        Dim cphHash As New Fix

        cphHash.KeyString = strA & "123456"
        cphHash.Text = strA & "123456"
        cphHash.DoXor()
        cphHash.Stretch()
        cphHash.KeyString = cphHash.Text
        cphHash.Text = "123456"
        cphHash.DoXor()
        cphHash.Stretch()
        Hash = cphHash.Text

    End Function

    Private Sub mainfrm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        gstrProbComtStack = "mainfrm_Load - Start"
        Me.Text = NameMe(Me.Text)

        gstrProbComtStack &= " #MF1"

        cboDBVersion.Text = "2000 / 2002"
        LinkLabel1.Links.Add(0, LinkLabel1.Text.Length, LinkLabel1.Text)

        gstrProbComtStack &= " #MF2"

        If IsAboveOrEqualWinXp() = True Then
            btnCreate.FlatStyle = FlatStyle.System
            btnSelectSource.FlatStyle = FlatStyle.System
            btnSelectAll.FlatStyle = FlatStyle.System
            btnClearAll.FlatStyle = FlatStyle.System
            btnAddJetString.FlatStyle = FlatStyle.System
            btnOpenDB.FlatStyle = FlatStyle.System
            rdoCreateNew.FlatStyle = FlatStyle.System
            rdoOpenSecured.FlatStyle = FlatStyle.System
        End If

        gstrProbComtStack &= " #MF3"
        SetBackcolors()

        gstrProbComtStack &= " #MF4"
        ShowWebHelpPage(WebHelpPages.welcome)

        btnOpenDB.Enabled = False

        lstrCFUCode = GetSetting("CFU Code", "10", InitalXMLConfig.XmlConfigType.AppSettings, "")

        gstrProbComtStack &= " #MFEnd" : AddDebugComment(gstrProbComtStack) : gstrProbComtStack = ""

    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Private Sub mnuHelpCFU_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpCFU.Click

        AddDebugComment("mainfrm.mnuHelpCFU_Click")

        Application.DoEvents()

        Dim StandLangText As System.Resources.ResourceManager = New _
            System.Resources.ResourceManager("AppBasic.AppBasic", _
            System.Reflection.Assembly.Load("AppBasic"))

        If hasMultipleInstances("MDBSecure", NameMe(""), Me.Handle, StandLangText) = True Then
            Exit Sub
        End If

        Dim NewCFU As New frmCFU(True)
        With NewCFU
            .Caption = NameMe("")
            .FormIcon = Me.Icon
            .strManifestSite(gstrManifestSite)
            .Owner = Me
            .ShowDialog()
        End With

        If gbooNeedToRestartProgAfterCFU = True Then
            Me.Close()
        End If
    End Sub
    Private Sub chkCopyTables_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCopyTables.CheckedChanged
        Dim lbooTrueOrFlase As Boolean = False

        If chkCopyTables.Checked = True Then
            lbooTrueOrFlase = True
            lstSourceTables.BackColor = System.Drawing.SystemColors.Window
        Else
            lstSourceTables.BackColor = System.Drawing.SystemColors.Control
        End If

        grbSourceDB.Enabled = lbooTrueOrFlase
        lstSourceTables.Enabled = lbooTrueOrFlase
        btnSelectSource.Enabled = lbooTrueOrFlase

        ShowWebHelpPage(WebHelpPages.Append)

    End Sub
    Private Sub mnuHelpAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click

        AddDebugComment("mainfrm.mnuHelpAbout_Click")

        Dim NewAbout As New frmAbout
        NewAbout.Owner = Me
        NewAbout.ShowDialog()

    End Sub
    Private Sub mnuHelpLicense_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpLicense.Click

        AddDebugComment("mainfrm.mnuHelpLicense_Click")

        Dim LicenseBox As New LicenceBox
        With LicenseBox
            .FormIcon = Me.Icon
            If InStrGet((NameMe("")).ToUpper, "TRIAL") = 0 Then
                .LicenseSectionSkip = "loan, copy, "
                'LOCALISATION DIFFERENCE HERE
            End If
            .ProdName = NameMe("").ToUpper
            .SetPageSettings = m_PageSettings
            .Owner = Me
            .ShowDialog()
            m_PageSettings = .SetPageSettings
        End With

    End Sub
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        AddDebugComment("mainfrm.LinkLabel1_LinkClicked")

        BrowseToUrl("http://www.example.com", Me)

    End Sub
    Private Sub btnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAll.Click

        AddDebugComment("mainfrm.btnSelectAll_Click")
        Dim i As Integer
        For i = 0 To Me.lstSourceTables.Items.Count - 1
            Me.lstSourceTables.SetSelected(i, True)
        Next

    End Sub
    Private Sub Progress()

        Select Case lblProgress.BackColor.Name
            Case Color.DarkBlue.Name : lblProgress.BackColor = Color.MediumBlue
            Case Color.MediumBlue.Name : lblProgress.BackColor = Color.Blue
            Case Color.Blue.Name : lblProgress.BackColor = Color.LightSkyBlue
            Case Color.LightSkyBlue.Name : lblProgress.BackColor = System.Drawing.SystemColors.Control
            Case System.Drawing.SystemColors.Control.Name : lblProgress.BackColor = Color.DarkBlue
        End Select
        System.Windows.Forms.Application.DoEvents()
    End Sub
    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click

        AddDebugComment("mainfrm.btnClearAll_Click")

        Dim i As Integer
        For i = 0 To Me.lstSourceTables.Items.Count - 1
            Me.lstSourceTables.SetSelected(i, False)
        Next

    End Sub
    Private Sub mnuHelpSupport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpSupport.Click

        AddDebugComment("mainfrm.mnuHelpSupport_Click")

        BrowseToUrl("http://www.example.com/supporthelp.php", Me)

    End Sub
    Private Sub btnOpenDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenDB.Click

        If CheckFields() = False Then
            Exit Sub
        End If

        AddDebugComment("mainfrm.btnOpenDB_Click 1")

        Application.DoEvents()

        Dim lstrMDBFile As String = SaveFileDialog1.FileName
        If System.IO.File.Exists(lstrMDBFile) = False Or System.IO.File.Exists(lstrMDBFile.Replace(".mdb", ".mdw")) = False Then
            lstrMDBFile = JustSpecify()
            If System.IO.File.Exists(lstrMDBFile) = False Or System.IO.File.Exists(lstrMDBFile.Replace(".mdb", ".mdw")) = False Then
                lstrMDBFile = JustSpecify()
                If System.IO.File.Exists(lstrMDBFile) = False Or _
                    System.IO.File.Exists(lstrMDBFile.Replace(".mdb", ".mdw")) = False Or _
                    lstrMDBFile = "" Then
                    Exit Sub
                End If
            End If
            If lstrMDBFile = "" Then
                Exit Sub
            End If
        End If

        AddDebugComment("mainfrm.btnOpenDB_Click 2")

        Dim MDW As String = LeftGet(lstrMDBFile, lstrMDBFile.Length - 1) & "w"
        Dim OpenMDB As New frmOpenMDB
        With OpenMDB
            .CommParam = ChrGet(34) & lstrMDBFile & ChrGet(34) & " /wrkgrp " & ChrGet(34) & MDW & ChrGet(34) & " /user " & txtSuperUser.Text & " /pwd " & txtNewPassword.Text & " /Excl"
            .DBPassword = txtDatabasePassword.Text
            .Owner = Me
            .ShowDialog()
        End With

        AddDebugComment("mainfrm.btnOpenDB_Click 3")

    End Sub
    Private Function JustSpecify() As String

        With OpenFileDialog1
            .CheckPathExists = True
            .AddExtension = True
            .FileName = "Project"
            .Filter = "Database file (*.mdb)|*.mdb"
            .DefaultExt = "mdb"
        End With

        If OpenFileDialog1.ShowDialog() <> DialogResult.Cancel And OpenFileDialog1.FileName <> "" Then
            Return OpenFileDialog1.FileName
        Else
            Return ""
        End If
    End Function
    Private Sub ShowWebHelpPage(ByVal Page As WebHelpPages)

        AddDebugComment("mainfrm.ShowWebHelpPage - start")

        If chkDynHelp.Checked = False Then
            Exit Sub
        End If

        If Overide = False Then
            Busy(Me, True)
            Dim lstrPageFile As String
            Dim lstrExt As String = ".html"
            Select Case Page
                Case WebHelpPages.welcome : lstrPageFile = "Welcome"
                Case WebHelpPages.AdminUser : lstrPageFile = "adminuser"
                Case WebHelpPages.DatabasePassword : lstrPageFile = "databasepassword"
                Case WebHelpPages.SuperUser : lstrPageFile = "superuser"
                Case WebHelpPages.SuperUserGroup : lstrPageFile = "superusergroup"
                Case WebHelpPages.Append : lstrPageFile = "append"
                Case WebHelpPages.Open : lstrPageFile = "open"
                Case WebHelpPages.Create : lstrPageFile = "create"
                Case WebHelpPages.JetString : lstrPageFile = "jetstring"
            End Select

            AddDebugComment("mainfrm.ShowWebHelpPage " & lstrPageFile)

            lstrPageFile = lstrPageFile & lstrExt

            If File.Exists(Application.StartupPath & "\help\" & lstrPageFile) = True Then
                Try
                    AxWebbrowser1.Navigate(Application.StartupPath & "\help\" & lstrPageFile)
                Catch
                    AddDebugComment("mainfrm.ShowWebHelpPage Error 1")
                    AxWebbrowser1 = Nothing
                    AddDebugComment("mainfrm.ShowWebHelpPage Error 2")
                    AxWebbrowser1 = New WinOnly.WebOCHostCtrl
                    AddDebugComment("mainfrm.ShowWebHelpPage Error 3")
                    AxWebbrowser1.Navigate(Application.StartupPath & "\help\" & lstrPageFile)
                End Try
            End If

            AddDebugComment("mainfrm.ShowWebHelpPage - done")
            Busy(Me, False)
        End If

    End Sub
    Private Sub txtNewPassword_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNewPassword.Enter
        ShowWebHelpPage(WebHelpPages.SuperUser)
    End Sub
    Private Sub txtSuperUser_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSuperUser.Enter
        ShowWebHelpPage(WebHelpPages.SuperUser)
    End Sub
    Private Sub txtSuperGroup_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSuperGroup.Enter
        ShowWebHelpPage(WebHelpPages.SuperUserGroup)
    End Sub
    Private Sub txtAdminPassword_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdminPassword.Enter
        ShowWebHelpPage(WebHelpPages.AdminUser)
    End Sub
    Private Sub txtDatabasePassword_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDatabasePassword.Enter
        ShowWebHelpPage(WebHelpPages.DatabasePassword)
    End Sub
    Private Sub btnSelectSource_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSelectSource.Enter
        ShowWebHelpPage(WebHelpPages.Append)
    End Sub
    Private Sub btnSelectAll_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSelectAll.Enter
        ShowWebHelpPage(WebHelpPages.Append)
    End Sub
    Private Sub btnClearAll_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClearAll.Enter
        ShowWebHelpPage(WebHelpPages.Append)
    End Sub
    Private Sub lstSourceTables_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSourceTables.Enter
        ShowWebHelpPage(WebHelpPages.Append)
    End Sub
    Private Sub btnOpenDB_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenDB.Enter
        ShowWebHelpPage(WebHelpPages.Open)
    End Sub
    Private Sub btnCreate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreate.Enter
        ShowWebHelpPage(WebHelpPages.Create)
    End Sub
    Private Sub mnuHelpEnterCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpEnterCode.Click

        AddDebugComment("mainfrm.mnuHelpEnterCode_Click")

        If AcceptLicense(Me) = True Then
            Me.Text = NameMe("")
            StandardUpgradeTidy()
        End If

    End Sub
    Private Sub mnuHelpInstallPack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpInstallPack.Click
        Dim lstrAddOnFile As String

        AddDebugComment("mainfrm.mnuHelpInstallPack_Click")

        Application.DoEvents()

        Dim StandLangText As System.Resources.ResourceManager = New _
            System.Resources.ResourceManager("AppBasic.AppBasic", _
            System.Reflection.Assembly.Load("AppBasic"))
        If hasMultipleInstances("MDBSecure", NameMe(""), Me.Handle, StandLangText) = True Then
            Exit Sub
        End If

        Dim OpenAddon As New OpenFileDialog
        With OpenAddon
            .CheckFileExists = True
            .CheckPathExists = True
            .Filter = "Mindwarp Consultancy Ltd AddOn (*.mcla)|*.mcla"
            If .ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If
            lstrAddOnFile = .FileName

        End With
        Dim lstrDat As Date = Date.Now

        gstrCFUTempDir = System.IO.Path.GetDirectoryName( _
            System.Reflection.Assembly.GetEntryAssembly.Location.ToString()) & "\Temp-" & _
            lstrDat.ToString("dddd-dd-MMM-yyyy-HH-mm-ss")
        Dim LangText2 As System.Resources.ResourceManager = New _
            System.Resources.ResourceManager("AppBasic.AppBasic", _
            System.Reflection.Assembly.Load("AppBasic"))

        Try
            System.IO.Directory.CreateDirectory(gstrCFUTempDir & "\unzip")
            AppBasic.UpdateFuncs.Unzip(lstrAddOnFile, gstrCFUTempDir & "\unzip\")

            Dim InitialConfig As New InitalXMLConfig(InitalXMLConfig.XmlConfigType.AppSettings, "", gstrCFUTempDir & "\unzip\addon.dat")
            With InitialConfig

                If AppBasic.IsCompatible(.GetValue("AppVersion", "")) = False Then
                    Directory.Delete(gstrCFUTempDir, True)
                    Throw New Exception(" ")
                End If

            End With

            MessageBox.Show(LangText2.GetString("CFU_ProgRestart"), NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Information)
            gbooNeedToRestartProgAfterCFU = True
            Me.Close()
        Catch
            MessageBox.Show(LangText2.GetString("CFU_DownloadIncompatible"), NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

        Try
            File.Delete(gstrCFUTempDir & "\unzip\addon.dat")
        Catch
            '
        End Try

    End Sub
    Private Sub mainfrm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        AddDebugComment("mainfrm_Closing")
        DeleteTempFiles()

        If InStrGet((NameMe("")).ToUpper, "TRIAL") > 0 Then
            Welcome(lbooSplashShown, Me)
        End If

        AddDebugComment("mainfrm_Closing About to set AxWebbrowser1 = nothing")

        Try
            AxWebbrowser1 = Nothing
        Catch
            AddDebugComment("<Font color=Blue>mainfrm_Closing AxWebbrowser1 = nothing</font>")
        End Try


    End Sub
    Private Sub btnAddJetString_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddJetString.Click

        AddDebugComment("mainfrm.btnAddJetString_Click")

        If CheckFields() = False Then
            Exit Sub
        End If

        Dim lstrMDBFile As String = SaveFileDialog1.FileName
        If System.IO.File.Exists(lstrMDBFile) = False Then
            lstrMDBFile = JustSpecify()
            If lstrMDBFile = "" Then
                Exit Sub
            End If
        End If

        Dim MDW As String = LeftGet(lstrMDBFile, lstrMDBFile.Length - 1) & "w"

        Dim s As String = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=" & txtDatabasePassword.Text & _
        ";User ID=" & txtSuperUser.Text & ";Password=" & txtNewPassword.Text & ";Data Source=" & _
        SaveFileDialog1.FileName & ";Jet OLEDB:System database=" & MDW

        If InStrGet((NameMe("")).ToUpper, "TRIAL") > 0 Then
            s &= Environment.NewLine & Environment.NewLine & NameMe("")
        End If

        Clipboard.SetDataObject(s, True)
        MessageBox.Show("The Jet Connection string for this database has been added to the clipboard!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub
    Private Sub btnAddJetString_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddJetString.Enter
        ShowWebHelpPage(WebHelpPages.JetString)
    End Sub
    Private Function CheckFields() As Boolean

        AddDebugComment("mainfrm.CheckFields")

        CheckFields = False

        If txtAdminPassword.Text = "" Then
            MessageBox.Show("You must enter an Admin Password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtAdminPassword)
            Exit Function
        ElseIf txtAdminPassword.Text.ToUpper = "NONE" Then
            MessageBox.Show("You may not use NONE as an Admin Password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtAdminPassword)
            Exit Function
        ElseIf LeftGet(txtAdminPassword.Text, 1) = " " Then
            MessageBox.Show("Your Admin Password must not begin with a space!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtAdminPassword)
            Exit Function
        ElseIf RightGet(txtAdminPassword.Text, 1) = " " Then
            MessageBox.Show("Your Admin Password must not end with a space!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtAdminPassword)
            Exit Function
        End If

        If txtDatabasePassword.Text = "" Then
            MessageBox.Show("You must enter a Database Password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtDatabasePassword)
            Exit Function
        ElseIf txtDatabasePassword.Text.ToUpper = "NONE" Then
            MessageBox.Show("You may not use NONE as a Database Password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtDatabasePassword)
            Exit Function
        ElseIf LeftGet(txtDatabasePassword.Text, 1) = " " Then
            MessageBox.Show("Your Database Password must not begin with a space!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtAdminPassword)
            Exit Function
        ElseIf RightGet(txtDatabasePassword.Text, 1) = " " Then
            MessageBox.Show("Your Admin Password must not end with a space!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtAdminPassword)
            Exit Function
        End If

        If InStrGet(txtDatabasePassword.Text, "-") > 0 Then
            MessageBox.Show("You cannot have a hyphen in your Database Password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtDatabasePassword)
            Exit Function
        End If

        If txtNewPassword.Text = "" Then
            MessageBox.Show("You must enter a Super User Password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtNewPassword)
            Exit Function
        ElseIf txtNewPassword.Text.ToUpper = "NONE" Then
            MessageBox.Show("You may not use NONE as a Super User Password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtNewPassword)
            Exit Function
        ElseIf LeftGet(txtNewPassword.Text, 1) = " " Then
            MessageBox.Show("Your Super User Password must not begin with a space!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtAdminPassword)
            Exit Function
        ElseIf RightGet(txtNewPassword.Text, 1) = " " Then
            MessageBox.Show("Your Super User Password must not end with a space!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtAdminPassword)
            Exit Function
        End If

        If txtSuperGroup.Text = "" Then
            MessageBox.Show("You must enter a Super Group name!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtSuperGroup)
            Exit Function
        ElseIf txtSuperGroup.Text.ToUpper = "NONE" Then
            MessageBox.Show("You may not use NONE a Super Group name!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtSuperGroup)
            Exit Function
        End If

        If txtSuperUser.Text = "" Then
            MessageBox.Show("You must enter a Super User name!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtSuperUser)
            Exit Function
        ElseIf txtSuperUser.Text.ToUpper = "NONE" Then
            MessageBox.Show("You may not use NONE as a Super User name!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtSuperUser)
            Exit Function
        End If

        If txtDatabasePassword.Text.ToUpper = "DATABASE" Then
            MessageBox.Show("'Database' is not permitted as a database password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtDatabasePassword)
            Exit Function
        ElseIf txtDatabasePassword.Text.ToUpper = "NONE" Then
            MessageBox.Show("You may not use NONE as a database password!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtDatabasePassword)
            Exit Function
        End If

        If txtSuperUser.Text.ToUpper = txtSuperGroup.Text.ToUpper Then
            MessageBox.Show("You may not use the same name for Super User and Super Group!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SelectText(txtSuperUser)
            Exit Function
        End If

        CheckFields = True

    End Function
    Private Sub SelectText(ByVal pobjText As TextBox)

        With pobjText
            .SelectionStart = 0
            .SelectionLength = .Text.Length
            .Select()
        End With

    End Sub
    Private Sub mainfrm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        If lbooLoadDataNow = True Then
            AddDebugComment("mainfrm_Activated - lbooLoadDataNow = True")
            Me.TopMost = True
            Me.TopMost = False

            lbooLoadDataNow = False
            AddDebugComment("mainfrm_Activated - 2")

            If InStrGet((NameMe("")).ToUpper, "TRIAL") > 0 Then
                If lstrCFUCode = "30" Then
                    AddDebugComment("mainfrm_Activated - 3")
                    SaveSetting("CFU Code", "", InitalXMLConfig.XmlConfigType.AppSettings, "")
                    MessageBox.Show("This version of the program will from now on normally be a 30 day trial.  As you have upgraded from an older version of the program (90 day version) you are entitled to the 90 day (approx) trial. Please ignore any reference to 30 days in the package!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

            AddDebugComment("mainfrm_Activated - 6M")
            Dim lbooSixMonthVersionCFUDone As Boolean = CBool(GetSetting("SixMonthVersionCFUDone", False, InitalXMLConfig.XmlConfigType.AppSettings, ""))
            If lbooSixMonthVersionCFUDone = False Then
                Dim BuildDate As Date = IO.File.GetLastWriteTime(System.Reflection.Assembly.GetEntryAssembly.Location.ToString())
                Dim Now As Date = Date.Now '.AddMonths(7) 'Add 7 months for testing purposes
                If BuildDate.AddMonths(6) < Now Then
                    Dim DlgRes As DialogResult = MessageBox.Show("This program is at least six months old, there may now be a newser version." & Environment.NewLine & Environment.NewLine & _
                    "Would you like to check for updates?", NameMe("Six Month Update Check"), MessageBoxButtons.YesNo)
                    If DlgRes = DialogResult.Yes Then
                        mnuHelpCFU_Click(Nothing, Nothing)
                    End If
                End If
                SaveSetting("SixMonthVersionCFUDone", True, InitalXMLConfig.XmlConfigType.AppSettings, "")
            End If
        End If

    End Sub
    Private Sub SetBackcolors()

        btnCreate.BackColor = Color.FromArgb(0, btnCreate.BackColor)
        btnSelectSource.BackColor = Color.FromArgb(0, btnSelectSource.BackColor)
        btnSelectAll.BackColor = Color.FromArgb(0, btnSelectAll.BackColor)
        btnClearAll.BackColor = Color.FromArgb(0, btnClearAll.BackColor)
        btnAddJetString.BackColor = Color.FromArgb(0, btnAddJetString.BackColor)
        btnOpenDB.BackColor = Color.FromArgb(0, btnOpenDB.BackColor)
        chkCopyTables.BackColor = Color.FromArgb(0, chkCopyTables.BackColor)
        chkDynHelp.BackColor = Color.FromArgb(0, chkDynHelp.BackColor)
        grbSourceDB.BackColor = Color.FromArgb(0, grbSourceDB.BackColor)
        GroupBox1.BackColor = Color.FromArgb(0, GroupBox1.BackColor)
        LinkLabel1.BackColor = Color.FromArgb(0, LinkLabel1.BackColor)
        Label8.BackColor = Color.FromArgb(0, Label8.BackColor)
        Label9.BackColor = Color.FromArgb(0, Label9.BackColor)
        grpUserChoices.BackColor = Color.FromArgb(0, grpUserChoices.BackColor)

    End Sub
    Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)

        Dim PaintBack As New UIStyle.Painting
        PaintBack.PaintBackground(pevent, Me)

        If IsAboveOrEqualWinXp() = True Then
            '
        Else
            chkCopyTables.BackgroundImage = PaintBack.PaintBackgroundToFit(pevent, Me.Height, Me.Width, chkCopyTables.Top, chkCopyTables.Left, chkCopyTables.Width, chkCopyTables.Height)
            chkDynHelp.BackgroundImage = PaintBack.PaintBackgroundToFit(pevent, Me.Height, Me.Width, chkDynHelp.Top, chkDynHelp.Left, chkDynHelp.Width, chkDynHelp.Height)
        End If

    End Sub

    Private Sub rdoUserChoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoCreateNew.Click, rdoOpenSecured.Click

        AddDebugComment("mainfrm.rdoUserChoice_Click - " & rdoCreateNew.Checked)

        If rdoCreateNew.Checked = True Then
            chkCopyTables.Enabled = True
            grbSourceDB.Enabled = True
            btnOpenDB.Enabled = False
            btnCreate.Enabled = True
            btnAddJetString.Enabled = True
        Else
            chkCopyTables.Enabled = False
            grbSourceDB.Enabled = False
            btnOpenDB.Enabled = True
            btnCreate.Enabled = False
            btnAddJetString.Enabled = False
        End If

    End Sub
    Private Sub rdoCreateNew_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoCreateNew.Enter
        ShowWebHelpPage(WebHelpPages.welcome)
    End Sub
    Private Sub rdoOpenSecured_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoOpenSecured.Enter
        ShowWebHelpPage(WebHelpPages.welcome)
    End Sub
    Private Sub mnuHelpReportProblem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpReportProblem.Click

        Application.DoEvents()

        Dim ErrRep As New ProbHand.BugReport(True)
        With ErrRep
            AddDebugComment("<Font color=Blue>Manual Problem Report</font>")
            DebugDBComment()

            .Caption = NameMe("")
            .Owner = Me
            .SysStartTime = gdatSystemStart
            .FormIcon = New System.Drawing.Icon( _
                System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("MDBSecure.mcl32.ico"))
            .ShowDialog()
        End With

    End Sub
    Private Sub AxWebbrowser1_DocumentComplete(ByVal sender As Object, ByVal e As Object) Handles AxWebbrowser1.DocumentComplete

        AddDebugComment("AxWebbrowser1_DocumentComplete")

    End Sub
    Private Sub AxWebbrowser1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxWebbrowser1.Leave

        AddDebugComment("AxWebbrowser1_Leave")

    End Sub
End Class
