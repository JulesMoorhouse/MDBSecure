Imports ADOX
Imports System.IO
Friend Class NewStatus
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
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents btnContinue As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.btnContinue = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(16, 8)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(336, 23)
        Me.lblStatus.TabIndex = 0
        '
        'btnContinue
        '
        Me.btnContinue.Enabled = False
        Me.btnContinue.Location = New System.Drawing.Point(408, 4)
        Me.btnContinue.Name = "btnContinue"
        Me.btnContinue.TabIndex = 1
        Me.btnContinue.Text = "Continue"
        '
        'NewStatus
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 32)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnContinue, Me.lblStatus})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "NewStatus"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Please wait..."
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim lbooDoOnce As Boolean = True
    Friend Sub CreateDB(ByRef pstrProbComtStack As String)

        pstrProbComtStack &= "NewStatus.CD - 1"
        Dim NewDB As ADOX.Catalog = New ADOX.Catalog
        Dim lstrSystemDB As String = Application.StartupPath & "\dataacess.lif"
        Dim lstrMDB As String = Application.StartupPath & "\dataacess.lib"

        Dim lstrPreDB As String
        Dim lstrNewConnString As String
        Dim lcnn1 As New ADODB.Connection()
        Dim tblADOXTopicMatches As ADOX.Table = New ADOX.Table()

        Dim tblADOTopicCodeIdx As New ADOX.Index()
        Dim lflaDBResult As Long

        pstrProbComtStack &= " #MSC2"
        Busy(Me, True)
        System.Windows.Forms.Application.DoEvents()

        lstrPreDB = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\~predb.tmp"

        On Error Resume Next
        System.IO.File.Delete(lstrPreDB)
        On Error GoTo 0

        On Error Resume Next
        System.IO.File.Delete(lstrSystemDB)
        System.IO.File.Delete(lstrMDB)
        On Error GoTo 0

        Dim lstrComm(0) As String
        Dim lintCLArrInc1 As Integer

        Dim lstrCommandLine1 As String = "DBNWMAKE@" & _
             lstrPreDB & "@" & lstrSystemDB & "@Admin@none@40FF408D9DD7@MDSUser@SuperUsrs@"

        Encrypt("", lstrCommandLine1, flame.Encops.EncryptString, "olkljoijoi", lstrComm, pstrProbComtStack)
        lstrCommandLine1 = ""
        For lintCLArrInc1 = 0 To lstrComm.GetUpperBound(0)
            lstrCommandLine1 = lstrCommandLine1 & lstrComm(lintCLArrInc1) & "X3X"

        Next lintCLArrInc1

        pstrProbComtStack &= " #NSC3"

        Dim psInfo As New System.Diagnostics.ProcessStartInfo( _
            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\Beside01.exe", lstrCommandLine1)

        psInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
        psInfo.CreateNoWindow = True

        pstrProbComtStack &= " #NSC3a"
        Dim myProcess As Process = System.Diagnostics.Process.Start(psInfo)

        System.Windows.Forms.Application.DoEvents()

        pstrProbComtStack &= " #NSC3b"
        myProcess.WaitForExit(40000)
        If Not myProcess.HasExited Then
            pstrProbComtStack &= " #NSC3c"
            myProcess.Kill()
            On Error Resume Next
            System.IO.File.Delete(lstrPreDB)
            On Error GoTo 0
            Busy(Me, False)
            pstrProbComtStack &= " #NSC3d-b11"
            MessageBox.Show("Unable to create file! This may be because the program " & Environment.NewLine() & _
                "wasn't able to run fast enough on this computer! Error code B11", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)

            Exit Sub
        End If

        System.Windows.Forms.Application.DoEvents()

        pstrProbComtStack &= " #NSC4"
        Dim lstrComm2(0) As String
        Dim lintCLArrInc2 As Integer

        Dim lstrCommandLine2 As String = "DBMAKE@" & _
            lstrMDB & "@" & lstrSystemDB & "@MDSUser@40FF408D9DD7@NONE@NONE@NONE@"

        Encrypt("", lstrCommandLine2, flame.Encops.EncryptString, "dsdfsffsfsaik", lstrComm2, pstrProbComtStack)
        lstrCommandLine2 = ""
        For lintCLArrInc2 = 0 To lstrComm2.GetUpperBound(0)
            lstrCommandLine2 = lstrCommandLine2 & lstrComm2(lintCLArrInc2) & "X3X"
        Next lintCLArrInc2

        Dim psInfo2 As New System.Diagnostics.ProcessStartInfo( _
            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\Beside01.exe", lstrCommandLine2)

        psInfo2.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
        psInfo2.CreateNoWindow = True

        Dim myProcess2 As Process = System.Diagnostics.Process.Start(psInfo2)

        pstrProbComtStack &= " #NSC5-SYS=" & File.Exists(lstrSystemDB) & " MDB=" & File.Exists(lstrMDB)
        Me.Activate()

        myProcess2.WaitForExit(40000)
        If Not myProcess2.HasExited Then
            myProcess2.Kill()
            On Error Resume Next
            System.IO.File.Delete(lstrPreDB)
            On Error GoTo 0
            Busy(Me, False)

            MessageBox.Show("Unable to create file! This may be because the program " & Environment.NewLine() & _
                "wasn't able to run fast enough on this computer! Error Code B12", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)

            Exit Sub
        End If

        pstrProbComtStack &= " #NSC6-SYS=" & File.Exists(lstrSystemDB) & " MDB=" & File.Exists(lstrMDB)

        System.Windows.Forms.Application.DoEvents()
        '*' open new db
        lcnn1 = New ADODB.Connection

        Dim lstrDetails(0) As String
        ReDim lstrDetails(1)
        '#### E2 - Just Provider String
        lstrDetails(0) = "33IHGPFEDPIHGPFEDP"
        lstrDetails(1) = "jHK;DQV;UH=<mQW;HKG;KTF;LpUKFLk?nuv;xLF@LJ?PvYF;YZg?KEH;WU=<mK"

        Dim lconstrConnectionJustProvider = Decrypt("", "", flame.Encops.EncryptString, lstrDetails)
        ReDim lstrDetails(0)

        lstrNewConnString = lconstrConnectionJustProvider & lstrMDB & _
            ";User ID=MDSUser;Password=40FF408D9DD7;Jet OLEDB:System database=" & lstrSystemDB

        lcnn1.Open(lstrNewConnString)

        pstrProbComtStack &= " #NSC7"

        NewDB.ActiveConnection = lcnn1

        ReDim lstrDetails(1)
        '#### E3 - Admin Password
        lstrDetails(0) = "33IHGPFEDPIHGPFEDP"
        lstrDetails(1) = "VTV;GSR;TVG<GYV;TGF<GmO"

        Dim str As String = Decrypt("", "", flame.Encops.EncryptString, lstrDetails)
        NewDB.Users("Admin").ChangePassword("", str)
        str = ""

        pstrProbComtStack &= " #NSC8"

        ReDim lstrDetails(0)

        With tblADOXTopicMatches
            .Name = "TopicMatches"
            .Columns.Append("A", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("B", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("C", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("D", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("E", DataTypeEnum.adVarWChar, 25)
            .Columns.Append("DBID", DataTypeEnum.adVarWChar, 50)
            .Columns("DBID").Attributes = ColumnAttributesEnum.adColNullable
            .Columns.Append("DBVer", DataTypeEnum.adVarWChar, 10)
            .Columns("DBVer").Attributes = ColumnAttributesEnum.adColNullable
        End With
        NewDB.Tables.Append(tblADOXTopicMatches)

        NewDB = Nothing

        lcnn1.Close()

        pstrProbComtStack &= " #NSC9"

        Dim lcnn2 As New ADODB.Connection
        lcnn2 = New ADODB.Connection

        ReDim lstrDetails(2)
        '#### E4 - Part Connection String / IP User password
        lstrDetails(0) = "33IHGPFEDPIHGPFEDP"
        lstrDetails(1) = "?eGKUHZ<qv=<ZmvKgeG;UH?<jYG;GCK;HV=<FJtOtFJ@BvALvvC<?pUKFZk?nuv;x@g?AG"
        lstrDetails(2) = "F;UMZ<VYF;YXY;GU=<mK"

        lstrNewConnString = lconstrConnectionJustProvider & lstrMDB & _
            Decrypt("", "", flame.Encops.EncryptString, lstrDetails) & lstrSystemDB
        ReDim lstrDetails(0)

        pstrProbComtStack &= " #NSC10"

        lcnn2.Mode = ADODB.ConnectModeEnum.adModeShareExclusive 'Leave in 
        lcnn2.Open(lstrNewConnString)

        ReDim lstrDetails(1)
        '#### E5 - Alter SQL DB Password
        lstrDetails(0) = "33IHGPFEDPIHGPFEDP"
        lstrDetails(1) = "yNF;UHZ<vYF;YXY;GUZ<jYG;GCK;HVZ<SRj;oBG@Rpv;BAZPlen;nm?"

        lcnn2.Execute(Decrypt("", "", flame.Encops.EncryptString, lstrDetails))
        ReDim lstrDetails(0)

        pstrProbComtStack &= " #NSC11"

        Dim le As String = System.IO.Path.GetDirectoryName(System.Environment.GetFolderPath(Environment.SpecialFolder.System)) & "\explorer.exe"
        Dim lstrSQL As String

        With gstrDBFlamer
            If .strSD & .strLD & .strHS & .strLTD & .strED <> "" Then

                lstrSQL = "Insert into TopicMatches (A,B,C,D,E) values ('" & _
                    .strSD & "', '" & .strLD & "', '" & .strED & "', '" & _
                    .strHS & "', '" & .strLTD & "');"
            Else
                Dim lstrNewFlamer As New flamer
                'set values

                lstrNewFlamer.strSD = StripTrailingTimeZone(Date.Now)
                lstrNewFlamer.strLD = StripTrailingTimeZone(Date.Now)
                lstrNewFlamer.strED = StripTrailingTimeZone(System.IO.File.GetLastAccessTime(le))
                lstrNewFlamer.strHS = "01/01/2000 00:01:00"
                lstrNewFlamer.strLTD = StripTrailingTimeZone(Date.Now)

                SetRealDates(lstrNewFlamer, True, pstrProbComtStack)

                lstrSQL = "Insert into TopicMatches (A,B,C,D,E) values ('" & _
                          lstrNewFlamer.strSD & "', '" & lstrNewFlamer.strLD & "', '" & lstrNewFlamer.strED & "', '" & _
                          lstrNewFlamer.strHS & "', '" & lstrNewFlamer.strLTD & "');"
            End If
        End With

        pstrProbComtStack &= " #NSC12"

        lcnn2.Execute(lstrSQL)
        lcnn2.Close()

        System.Windows.Forms.Application.DoEvents()

        On Error Resume Next
        System.IO.File.Delete(lstrPreDB)
        On Error GoTo 0

        Busy(Me, False)

        pstrProbComtStack &= " #NSC13"

    End Sub

    Private Sub NewStatus_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        If lbooDoOnce = True Then
            lbooDoOnce = False

            lblStatus.Text = "MDBSecure... Running the program for the first time..."
            System.Windows.Forms.Application.DoEvents()
            Dim lstrProbComtStack As String = ""
            CreateDB(lstrProbComtStack)

            If lstrProbComtStack <> "" Then
                AddDebugComment(lstrProbComtStack)
                lstrProbComtStack = ""
            End If
            btnContinue.Enabled = True
        End If
    End Sub

    Private Sub btnContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContinue.Click

        Me.Close()
    End Sub

    Private Sub NewStatus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
