Imports System.Drawing.Printing
Imports System.IO
Friend Module mdbsecdemo
    Private Const mstrTitle As String = "Take a minute to understand what MDBSecure can do for you!"
    Private Const mstrBullet1 As String = "       Takes the pain out of securing your databases"
    Private Const mstrBullet2 As String = "       Easy to use, step-by-step help included!"
    Private Const mstrBullet3 As String = "       Caters for creativity and future change"
    Private Const mstrBullet4 As String = "       Free email support"
    Public gstrDecryptProbTrace As String = ""
    Friend Sub Welcome(ByRef pbooSplashShown As Boolean, ByVal powner As Form)
        If pbooSplashShown = False Then
            pbooSplashShown = True
            Dim BetaSplash As New strat3welcome
            BetaSplash.Title = mstrTitle
            BetaSplash.Bullet1 = mstrBullet1
            BetaSplash.Bullet2 = mstrBullet2
            BetaSplash.Bullet3 = mstrBullet3
            BetaSplash.Bullet4 = mstrBullet4
            BetaSplash.Owner = powner
            BetaSplash.ShowDialog()
        End If
        powner.Activate()

    End Sub
    Friend mintVersion As Integer
    Friend lstrTempFiles(0) As String
    Sub main()

        Dim lstrEssentialFiles() As String = {"ADODB.dll", "ADOX.dll", "AppBasic.dll", "AxInterop.SHDocVw.dll", _
            "Beside01.exe", "Beside02.exe", "MDBSecure.exe", "JRO.dll", "MCLCore.dll", "SharpZipLib.dll", "SHDocVw.dll", _
            "WinOnly.dll", "help\databasepassword.html", "help\superusergroup.html", "help\welcome.html", _
            "help\adminuser.html", "help\superuser.html", "help\append.html", "help\open.html", "help\create.html", _
            "help\jetstring.html"}

        Dim lstrPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString())

        If EssentialFileCheck(lstrEssentialFiles, lstrPath) = False Then
            MessageBox.Show("Some essential program files are missing, please re-install the program!", "MDBSecure", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If


        mainStart()
    End Sub
    Sub mainStart()

        Dim lflaDBResult As Long

        Dim eh1 As CustomExceptionHandler = New CustomExceptionHandler
        Try
            gstrProbComtStack = "mainStart - Topmost"

            gdatSystemStart = Date.Now

            Dim lintThreads As Integer = 259

            With gstrManifestSite(0)
                .strSitePath = "http://www.example.com"
                .strManifestDir = "updates/newton/"
                .strManifestFile = "MDBSeDemo.xml"
            End With

            With gstrManifestSite(1)
                .strSitePath = "http://www.example.com"
                .strManifestDir = "updates/newton/"
                .strManifestFile = "MDBSeDemo.xml"
            End With

            If System.IO.File.Exists(System.Reflection.Assembly.GetEntryAssembly.Location.ToString() & ".dat") = False Then
                GetSetting("CFU Code", "10", InitalXMLConfig.XmlConfigType.AppSettings, "")
            End If

            gstrProbComtStack &= " #MS1"

            If flamenow() Then ' Soft ICE Check
                Dim lstrDetails2(2) As String
                Dim lstrDetails1(1) As String
                Dim lstrRetVal As String
                Dim lstrRetVal1 As String

                lstrDetails2(0) = "33IHGPFEDPIHGPFEDP"
                lstrDetails2(1) = "ZRYKGLS<FZS?KFZ<ULK;ESR;ZGAKGFU;MZH?UGK;EHW;UGZ<FKZ<HEL;ZJHKKJU;HNA;Lm"
                lstrDetails2(2) = "pPmK"

                lstrRetVal = Decrypt("", "", flame.Encops.EncryptString, lstrDetails2)

                lstrDetails1(0) = "33IHGPFEDPIHGPFEDP"
                lstrDetails1(1) = "jNU;YGU;ZWNKKGU;ZGKKMUZ<KTZ<AKE;HZK?FRU;HZJ?HKS;HYM;GZY?LVZ<HUF;HAY<mK"

                lstrRetVal1 = Decrypt("", "", flame.Encops.EncryptString, lstrDetails1)

                MessageBox.Show(NameMe("") & lstrRetVal & Environment.NewLine & Environment.NewLine & lstrRetVal1, NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)

                lstrRetVal1 = ""
                lstrRetVal = ""
                End
                GoTo Exhaust
            End If

            gstrProbComtStack &= " #MS2"

            '------------------ crc check -------------------------
            Dim IRes As Integer
            IRes = GetWrittenCRC(AppExe)

            Dim lstrDetails(2) As String
            Dim lstrRetVal3 As String
            lstrDetails(0) = "33IHGPFEDPIHGPFEDP"
            lstrDetails(1) = "wKJ;AHQ;SRF;ZHJPJGZPmQL;VCY;HJZ<wKL;GEN;FYL;WAZ<nFV;LmpPyNN;ZHQKSRF;GZ"
            lstrDetails(2) = "H?UGU;HDU;VLm@pmO"

            lstrRetVal3 = Decrypt("", "", flame.Encops.EncryptString, lstrDetails)

#If Not Debug Then
            If IRes = 1 Then
                '
            ElseIf IRes = -1 Then
                MessageBox.Show(NameMe("") & Environment.NewLine & Environment.NewLine & lstrRetVal3, NameMe("") & "   ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End
                GoTo Exhaust
            Else
                MessageBox.Show(NameMe("") & Environment.NewLine & Environment.NewLine & lstrRetVal3, NameMe("") & "   ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End
                GoTo Exhaust
                End If
#End If
            '------------------ crc check -------------------------

            gstrProbComtStack &= " #MS3"

            Dim Dets As strat1.UnlockDetails
            TakeCare(lintThreads, Dets)

            Dim lstrSystemDB As String = Application.StartupPath & "\dataacess.lif"
            Dim lstrMDB As String = Application.StartupPath & "\dataacess.lib"

            gstrProbComtStack &= " #MS3a SYS=" & File.Exists(lstrSystemDB) & " MDB=" & File.Exists(lstrMDB)

            Dim StandLangText As System.Resources.ResourceManager = New _
                System.Resources.ResourceManager("AppBasic.AppBasic", _
                System.Reflection.Assembly.Load("AppBasic"))

            gstrProbComtStack &= " #MS4"

            Try
                If GetMDACVerNum() < 2.7 Then
                    MessageBox.Show(StandLangText.GetString("StandText_MDACReq").Replace("[Program]", NameMe("")), NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Catch
                MessageBox.Show(StandLangText.GetString("StandText_MDACReq").Replace("[Program]", NameMe("")), NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            gstrProbComtStack &= " #MS5"

            Try
                If IsJet4Installed() <> 1 Then
                    MessageBox.Show(StandLangText.GetString("StandText_MDACReq").Replace("[Program]", NameMe("")).Replace("Data Access Components (MDAC) 2.7", "Jet 4.0 Service Pack 8"), NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Catch
                MessageBox.Show(StandLangText.GetString("StandText_MDACReq").Replace("[Program]", NameMe("")).Replace("Data Access Components (MDAC) 2.7", "Jet 4.0 Service Pack 8"), NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            gstrProbComtStack &= " #MS6"
            AddDebugComment(gstrProbComtStack) : gstrProbComtStack = ""

            If File.Exists(lstrMDB) = False Then
                Dim NewDB As New NewStatus
                NewDB.ShowDialog()
            End If

            gstrProbComtStack &= " #MS6B - SYS=" & File.Exists(lstrSystemDB) & " MDB=" & File.Exists(lstrMDB)

            gstrProbComtStack &= " #MS7"

            ReDim lstrDetails(3)
            '#### E1 - Connection String

            lstrDetails(0) = "33IHGPFEDPIHGPFEDP"
            lstrDetails(1) = "jHK;DQV;UH=<mQW;HKG;KTF;LpUKFLk?nuv;xLF@LJ?PpUF;ZknKuvx;@vYKFYX;YGU;Zj"
            lstrDetails(2) = "YKGGC;KHV;=SRKjoB<GRpKvBA@?eGKUHZ<qv=<mvg;eGU;H?j?YGG;CKH;V=F@JttKFJBP"
            lstrDetails(3) = "vAv?vC?@vYF;YZg?KEH;WU=<mpmPpmO"

            Dim lconstrConnectionProvider As String = Decrypt("", "", flame.Encops.EncryptString, lstrDetails)
            ReDim lstrDetails(0)

            lconstrConnectionProvider = lconstrConnectionProvider.Replace(Environment.NewLine, "")

            gstrConnectionString = lconstrConnectionProvider & lstrMDB & ";Jet OLEDB:System database=" & lstrSystemDB

            gstrProbComtStack &= " #MS7a"

            'Try
            lflaDBResult = GetWindowsDir(gstrDBFlamer, gstrProbComtStack)
            gstrProbComtStack &= " Basic Checks done" : AddDebugComment(gstrProbComtStack) : gstrProbComtStack = ""
        Catch ex As Exception
            AddDebugComment("<Font color=Red>MSG:" & ex.ToString & "</font>")
            eh1.OnThreadException(Nothing, Nothing)

        End Try

        If lflaDBResult = 2 Then
            Dim mdbsMain As New mainfrm
            Try
                Dim eh As CustomExceptionHandler = New CustomExceptionHandler
                AddHandler Application.ThreadException, AddressOf eh.OnThreadException

                Application.Run(mdbsMain)
                AddDebugComment("MDBSecure.mainStart - After Main Screen 1")

            Catch ex As Exception
                '
            End Try
            AddDebugComment("MDBSecure.mainStart - After Main Screen 2")
            Dim lstrProbComtStack As String = ""
            lflaDBResult = GetWindowsDir(gstrDBFlamer, lstrProbComtStack)

            If lstrProbComtStack <> "" Then
                AddDebugComment(lstrProbComtStack)
                lstrProbComtStack = ""
            End If

            AddDebugComment("MDBSecure.mainStart - After Main Screen 3")
            ProcessAnyCFU()
            AddDebugComment("MDBSecure.mainStart - After Main Screen 4")
        Else

            'need to show register screen, so people can still unlock and buy it.

            Dim BetaSplash As New strat3welcome
            BetaSplash.ShowInTaskbar = True
            BetaSplash.Expired = True
            BetaSplash.Title = mstrTitle
            BetaSplash.Bullet1 = mstrBullet1
            BetaSplash.Bullet2 = mstrBullet2
            BetaSplash.Bullet3 = mstrBullet3
            BetaSplash.Bullet4 = mstrBullet4
            BetaSplash.BuyNowURL = "http://www.example.com/purchase.php"
            BetaSplash.Icon = New System.Drawing.Icon( _
                System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("MDBSecure.mcl32.ico"))
            BetaSplash.ShowDialog()

        End If

Exhaust:
    End Sub
    Friend Function GetWindowsDir(ByVal pstrResource As flamer, ByRef pstrProbComtStack As String) As Long

        If InStrGet((NameMe("")).ToUpper, "TRIAL") = 0 Then
            Return 2
        Else
            Return CheckDates(pstrResource, 1, pstrProbComtStack)
        End If

    End Function
    Friend Function NameMe(ByVal pstrCaption As String) As String


        Dim lstrVersion As String
        If mintVersion = 32 Then
            lstrVersion = "MDBSecure " & gYear
        Else
            lstrVersion = "MDBSecure Trial Version"

        End If

        If pstrCaption <> "" Then
            If (pstrCaption).ToUpper = "MDBSECURE" Then
                NameMe = lstrVersion
            Else
                NameMe = lstrVersion & " - " & pstrCaption
            End If
        Else
            NameMe = lstrVersion
        End If

    End Function

    Private Function x(ByVal pstrString As String) As String
        Dim lintArrInc As Integer
        Dim lstrChar As String

        For lintArrInc = 0 To pstrString.Length - 1
            lstrChar = Microsoft.VisualBasic.Mid(pstrString, lintArrInc + 1, 1)
            If Microsoft.VisualBasic.Asc(lstrChar) >= 48 And Microsoft.VisualBasic.Asc(lstrChar) <= 57 Then
                x &= lstrChar
            End If
        Next lintArrInc
    End Function
    Friend Function AcceptLicense(ByVal pform As Form) As Boolean
        Dim dlgResult As DialogResult
        Dim lAcceptReg As New AcceptReg
        With lAcceptReg

            AcceptLicense = False
            .Caption = NameMe("")

            .ProdName = "MDBSecure " & gYear
            dlgResult = .ShowDialog()

            If dlgResult = DialogResult.OK Then
                'create temp license file
                Dim lstrTemp As String
                Dim clsEnc As New MyCrypto

                If x(.LicenseCode) = "" Then
                    MessageBox.Show("Your license code was not accepted!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Function
                End If
                lstrTemp = clsEnc.Encrypt(x(.LicenseCode), "bUnn1es#j*mp@thr")
                clsEnc = Nothing

                Dim lstrEncFile As String = System.IO.Path.GetDirectoryName( _
                    System.Reflection.Assembly.GetEntryAssembly.Location.ToString()) & "\Temp-" & _
                    Date.Now.ToString("dddd-dd-MMM-yyyy-HH-mm-ss") & ".tmp"


                ReDim Preserve lstrTempFiles(lstrTempFiles.GetUpperBound(0) + 1)
                lstrTempFiles(lstrTempFiles.GetUpperBound(0)) = lstrEncFile

                Dim lintFreeFile As Integer = Microsoft.VisualBasic.FreeFile()
                Microsoft.VisualBasic.FileOpen(lintFreeFile, lstrEncFile, Microsoft.VisualBasic.OpenMode.Output)
                Microsoft.VisualBasic.Print(lintFreeFile, lstrTemp)
                Microsoft.VisualBasic.FileClose(lintFreeFile)

                'check license
                Dim lintCheck As Integer = 16
                Try
                    lintCheck = Unlock(lstrEncFile, Nothing, "", "")
                Catch

                End Try

                If lintCheck <> 245 + 12 Then
                    MessageBox.Show("Your license code was not accepted!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Warning)

                    Try
                        System.IO.File.Delete(lstrEncFile)
                    Catch
                        '
                    End Try
                Else

                    'store license
                    Try
                        System.IO.File.Delete(System.IO.Path.GetDirectoryName( _
                        System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\keyfile.mcl")
                    Catch
                    End Try

                    System.IO.File.Copy(lstrEncFile, System.IO.Path.GetDirectoryName( _
                        System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\keyfile.mcl")

                    MessageBox.Show("Your license code was accepted!", NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Information)

                    AcceptLicense = True
                    mintVersion = 32
                End If
            End If
        End With
    End Function
    Friend Sub StandardUpgradeTidy()

        If InStrGet((NameMe("")).ToUpper, "TRIAL") = 0 Then

            Dim lstrDemoBuyPage As String = _
            System.Environment.GetFolderPath(Environment.SpecialFolder.StartMenu.Programs) & _
                "MDBSecure\How to Buy MDBSecure.url"

            Dim lbooSucess As Boolean = False
            If System.IO.File.Exists(lstrDemoBuyPage) = True Then
                Try
                    System.IO.File.Delete(lstrDemoBuyPage)
                    lbooSucess = True
                Catch ex As Exception
                    '
                End Try
            End If


            If lbooSucess = False Then
                lstrDemoBuyPage = _
               System.Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) & _
                   "\Programs\MDBSecure\How to Buy MDBSecure.url"
                If System.IO.File.Exists(lstrDemoBuyPage) = True Then
                    Try
                        System.IO.File.Delete(lstrDemoBuyPage)
                        lbooSucess = True
                    Catch ex As Exception
                        '
                    End Try
                End If
            End If
        End If



        Dim lstrFileStr As String
        Dim lstrPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()) & "\Help\Welcome.html"

        Dim OpenFile As FileStream = New FileStream(lstrPath, FileMode.Open, FileAccess.Read, FileShare.Read)

        Dim StreamReader As StreamReader = New StreamReader(OpenFile)
        lstrFileStr = StreamReader.ReadToEnd '.Read 'Line()
        StreamReader.Close()
        OpenFile.Close()

        Dim RepStr As String = "Please remember you only have a 30 day evaluation. After this time you will be unable to use the program. <a href='http://www.example.com/buy-products.html' target='_blank'>Click here to Buy!</a><BR><BR>"
        lstrFileStr = lstrFileStr.Replace(RepStr, "")

        Dim lintFreeFile As Integer = Microsoft.VisualBasic.FreeFile()
        Microsoft.VisualBasic.FileOpen(lintFreeFile, lstrPath, Microsoft.VisualBasic.OpenMode.Output)
        Microsoft.VisualBasic.Print(lintFreeFile, lstrFileStr)
        Microsoft.VisualBasic.FileClose(lintFreeFile)

    End Sub
    Friend Sub DeleteTempFiles()
        Dim lintArrInc As Integer

        Try
            For lintArrInc = 0 To lstrTempFiles.GetUpperBound(0)
                If lstrTempFiles(lintArrInc) <> "" Then
                    If RightGet(lstrTempFiles(lintArrInc), 1) = "\" Then 'directory
                        Try
                            Directory.Delete(lstrTempFiles(lintArrInc), True)
                        Catch
                        End Try
                    Else
                        Try
                            File.Delete(lstrTempFiles(lintArrInc))
                        Catch
                        End Try
                    End If
                End If
            Next lintArrInc
        Catch ex As Exception
        End Try
    End Sub
    Friend Sub BlackKeys(ByRef lstrKeys() As String)
        ReDim lstrKeys(0)
        lstrKeys(0) = "xxx"

    End Sub
    Friend Function DataFileProductIdent(ByVal pstrProductNumber As String) As String

        'not used in MDBSecure

    End Function
    Friend Sub DebugDBComment()
        'added

        If gstrProbComtStack <> "" Then
            AddDebugComment(gstrProbComtStack)
        End If
        AddDebugComment("main.LastSQL: " & gstrLastSQL)
        AddDebugComment("main.ConnStr: " & gstrConnectionString)

        Dim BesideErr As String
        Try : BesideErr = GetSetting("Beside01Err", "Null", InitalXMLConfig.XmlConfigType.AppSettings, "") : Catch : End Try
        AddDebugComment("main.BesideErr: " & BesideErr)

    End Sub
    Function RBDecypt(ByVal pstrInputFile As String) As String
        'not in use in MDBSecure
    End Function
End Module
Friend Class CustomExceptionHandler 'created
    Friend Sub OnThreadException(ByVal sender As Object, ByVal t As System.Threading.ThreadExceptionEventArgs)

        Dim lstrErrMsg As String = "Unknown Error"

        Try
            If Not t Is Nothing Then
                If InStrGet(t.Exception.ToString, "System.Runtime.InteropServices.COMException") > 0 Then
                    lstrErrMsg = t.Exception.Message
                    lstrErrMsg = lstrErrMsg.Replace("IPUser", "ADMIN")
                Else
                    lstrErrMsg = t.Exception.Message & Environment.NewLine
                End If
            End If
            'CFU etc here
            Try
                If lstrErrMsg <> "" Then

                    MessageBox.Show(lstrErrMsg & "[CR2]Please use the 'Check for Updates' feature, which will be shown after this message.[CR]A newer version of the program should have fixed any problems in older versions.[CR]If you have tried the 'Check for Updates' feature and already have the latest version[CR]of the program please refer to the help file for support details.".Replace("[CR2]", _
                        Environment.NewLine & Environment.NewLine).Replace("[CR]", _
                        Environment.NewLine), NameMe(""), MessageBoxButtons.OK, MessageBoxIcon.Error)

                    Dim NewCFU As New frmCFU(True)
                    With NewCFU
                        .Caption = NameMe("")

                        .FormIcon = New System.Drawing.Icon( _
                            System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("MDBSecure.mcl32.ico"))

                        .strManifestSite(gstrManifestSite)

                        Application.DoEvents()
                        .ShowDialog()
                    End With

                    Dim lstrError As String = ""
                    Dim lstrErrorEnc As String = ""
                    If Not t Is Nothing Then
                        lstrError = "<Font color=Red>MSG:" & t.ToString & "<BR>{" & lstrErrMsg.Replace(Environment.NewLine, "") & "}</font>"

                        lstrErrorEnc = "<Font color=Red>MSG:" & EncodeToHtml(t.ToString) & "<BR>{" & lstrErrMsg.Replace(Environment.NewLine, "") & "}</font>"
                    End If

                    If lstrError <> "" Then
                        AddDebugComment(lstrError)
                        AddDebugComment(lstrErrorEnc)
                    End If

                    DebugDBComment()

                    If gbooNeedToRestartProgAfterCFU = True And gstrCFUTempDir <> "" Then
                        Dim lstrReportNames() As String
                        LoadReportsNames(lstrReportNames)
                        DumpThisErrorLog(NameMe(""), lstrReportNames, gdatSystemStart)
                    Else
                        Dim ErrRep As New ProbHand.BugReport(True)
                        With ErrRep
                            .SysStartTime = gdatSystemStart
                            .Caption = NameMe("")
                            .FormIcon = New System.Drawing.Icon( _
                                System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("MDBSecure.mcl32.ico"))

                            Application.DoEvents()

                            .ShowDialog()
                        End With
                    End If

                    ProcessAnyCFU()

                End If
            Catch
                '
            End Try
        Finally
            Environment.Exit(0)
        End Try

    End Sub
    Function EncodeToHtml(ByVal pstrSentence As String) As String

        Dim lintArrInc As Integer
        Dim lstrOutPut As String

        pstrSentence = Web.HttpUtility.HtmlEncode(pstrSentence)

        For lintArrInc = 1 To pstrSentence.Length
            Dim lintChar As Integer = AscGet(MidGet(pstrSentence, lintArrInc, 1))
            Select Case lintChar
                Case 32 To 127, 160 To 255
                    lstrOutPut &= ChrGet(lintChar)
                Case Else
                    lstrOutPut &= "&#" & (lintChar) & ";"
            End Select
        Next lintArrInc

        Return lstrOutPut

    End Function
End Class
<DoNotObfuscateAttribute()> Friend Module Base
    <DoNotObfuscateAttribute()> Friend gstrManifestSite(1) As ManifestInfo
    <DoNotObfuscateAttribute()> Friend m_PageSettings As New PageSettings
    <DoNotObfuscateAttribute()> Friend gdatSystemStart As Date
    <DoNotObfuscateAttribute()> Friend gstrConnectionString As String
    <DoNotObfuscateAttribute()> Friend gbooDebug As Boolean = True 'false
    <DoNotObfuscateAttribute()> Friend gYear As String = "2005"
    <DoNotObfuscateAttribute()> Friend gstrLastSQL As String
    <DoNotObfuscateAttribute()> Friend gstrProbComtStack As String = ""

End Module
<System.AttributeUsage(AttributeTargets.Class Or AttributeTargets.Field _
Or AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.Enum)> _
Friend Class ObfuscateAttribute
    Inherits System.Attribute
End Class 'ObfuscateAttribute

<System.AttributeUsage(AttributeTargets.Class Or AttributeTargets.Field _
Or AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.Enum)> _
Friend Class DoNotObfuscateAttribute
    Inherits System.Attribute
End Class 'DoNotObfuscateAttribute

<System.AttributeUsage(AttributeTargets.Assembly, allowmultiple:=True)> _
Friend Class ObfuscateBlockAttribute
    Inherits System.Attribute
    Private BlockString As String
    Public Sub New(ByVal BlockString As String)
        MyClass.BlockString = BlockString
    End Sub
End Class 'ObfuscateBlockAttribute
'<Assembly: ObfuscateBlock("SomeString")>