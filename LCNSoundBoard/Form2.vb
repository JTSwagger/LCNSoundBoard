Imports System.Runtime.InteropServices
Imports System.Speech
Imports System.Speech.Recognition
Imports NDde.Client
Imports mshtml
Imports OpenQA.Selenium.Firefox
Imports System.Threading
Public Class Form2
    Dim Au As New AutoItX3Lib.AutoItX3
    Dim vars() As String

    Dim dde As New DdeClient("Firefox", "WWW_GetWindowInfo")
    Dim GramBuild As GrammarBuilder

    Dim headElement As HtmlElement
    Dim scriptElement As HtmlElement
    Dim xTypecast As HTMLElementCollection
    Dim element As IHTMLScriptElement
    Dim theurl As String = ""
    Dim DickGram As Recognition.DictationGrammar
    Dim objection As Integer
    Dim WaveOut As New NAudio.Wave.WaveOut()
    Dim WaveOut2 As New NAudio.Wave.WaveOut()
    Dim SECONDARIES As Boolean
    Public synth As New Speech.Synthesis.SpeechSynthesizer
    Public WithEvents recognizer As New Speech.Recognition.SpeechRecognitionEngine
    Dim CURRENTlblQuestion(27) As String
    Dim HOMEOWNER As Boolean = False
    Dim isplaying As Boolean
    Dim INSCO As String
    Dim POLSTART As String
    Dim POLEnd As String
    Dim VEHICLE As String
    Dim CurrentQ As Integer
    Dim CustomerName As String
    Dim globalFile As String
    Dim globalFile2 As String
    Dim globalfile3 As String
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpWindowText As String, ByVal cch As Integer) As Integer
    Public Delegate Function CallBack(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean
    Public Declare Function EnumWindows Lib "user32" (ByVal Adress As CallBack, ByVal y As Integer) As Integer
    Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Const SW_HIDE As Integer = 0
    Private Const SW_RESTORE As Integer = 9
    Private hWnd As Integer
    Public selectedIndex As Integer
    Private ActiveWindows As New System.Collections.ObjectModel.Collection(Of IntPtr)
    Public Const MOD_ALT As Integer = &H1 'Alt key
    Public Const MOD_CONTROL As Integer = &H2
    Public Const WM_HOTKEY As Integer = &H312
    Dim deviceNum1 As Integer = 0
    Dim DeviceNum2 As Integer = 0
    Dim totalCalls As Integer
    Dim totalLeads As Integer

    Dim NewMod As String = ""
    Public Function GetActiveWindows() As ObjectModel.Collection(Of IntPtr)

        EnumWindows(AddressOf Enumerator, 0)

        Return ActiveWindows

    End Function
    Dim theOldWindowTitle As String
    Dim theWindowTitle
    Dim CustName(1) As String
    Dim oldCust(1) As String
    Private Function Enumerator(ByVal hwnd As IntPtr, ByVal lParam As Integer) As Boolean      'FINDS THE AUTO INSURANCE LEAD WINDOW AND ASSIGNS SOUNDCLIP FILEPATHS BASED ON NAME OF CUSTOMER
        Dim customerName As String = ""

        Dim text As String = Space(Int16.MaxValue)
        Dim i As Integer = 0

        If IsWindowVisible(hwnd) Then
            GetWindowText(hwnd, text, Int16.MaxValue)
            If text.Contains("Auto Insurance") Then
                theWindowTitle = text
                Do While text.Substring(i, 1) <> "|"
                    customerName = customerName & text.Substring(i, 1)
                    i = i + 1


                Loop
                Dim name() As String = customerName.Split
                CustName(0) = name(0)
                CustName(1) = name(1)

                globalFile = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 1.wav"
                globalFile2 = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 3.wav"
                globalfile3 = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 2.wav"

            End If
        End If

        Return True

    End Function
    Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr,
                                     ByVal childAfter As IntPtr,
                                     ByVal lclassName As String,
                                     ByVal windowTitle As String) As IntPtr
    End Function
    Public Sub RollTheClip(Clip As String)                          'THIS PLAYS THE WAV FILE SOUNDCLIP THAT IS PASSED TO IT VIA STRING FILEPATH AND NAME
        WaveOut.Dispose()
        WaveOut2.Dispose()
        WaveOut = New NAudio.Wave.WaveOut()
        WaveOut2 = New NAudio.Wave.WaveOut()
        If deviceNum1 <> DeviceNum2 Then
            Dim WaveFile As New NAudio.Wave.WaveFileReader(Clip)
            Dim WaveFile2 As New NAudio.Wave.WaveFileReader(Clip)
            WaveOut.DeviceNumber = deviceNum1
            WaveOut.Init(WaveFile)
            WaveOut.Play()
            WaveOut2.DeviceNumber = DeviceNum2
            WaveOut2.Init(WaveFile2)
            WaveOut2.Play()
        Else

            Dim WaveFile As New NAudio.Wave.WaveFileReader(Clip)
            WaveOut.DeviceNumber = deviceNum1
            WaveOut.Init(WaveFile)
            WaveOut.Play()
        End If

    End Sub

    <DllImport("User32.dll")>
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr,                     'ASSIGNS GLOBAL HOTKEYS 
                  ByVal id As Integer, ByVal fsModifiers As Integer,
                  ByVal vk As Integer) As Integer
    End Function

    <DllImport("User32.dll")>
    Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer) As Integer

    End Function
    Private Sub Button69_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button19_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button26_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button25_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RegisterHotKey(Me.Handle, 169, Keys.None, Keys.Escape)
        RegisterHotKey(Me.Handle, 170, Keys.None, Keys.Right)
        RegisterHotKey(Me.Handle, 171, MOD_CONTROL, Keys.Right)
        RegisterHotKey(Me.Handle, 172, Keys.None, Keys.Left)

    End Sub

    Public Sub PushCallForward(callPos As Integer) 'ASKS THE NEXT QUESTION IN THE CALL PROGRESSION AND PLACES THE FOCUS ON THE CONTROL FOR PROPER DATA INPUT  
Redo:
        Select Case callPos
            Case 1                                                          'HELLO 
                If My.Computer.FileSystem.FileExists(globalFile2) Then
                    lblQUESTION.Text = CustomerName
                Else
                    lblQUESTION.Text = "This is..."
                End If
                RollTheClip("c:\soundboard\cheryl\INTRO\HELLO.wav")
                CurrentQ += 1
            Case 2                                                          'ASKS FOR FIRST NAME IF WE HAVE IT, IF NOT GOES INTO INTRO
                Try
                    RollTheClip(globalFile2)
                    lblQUESTION.Text = CustomerName

                Catch
                    RollTheClip("c:\soundboard\cheryl\INTRO\INTRO.wav")
                    CurrentQ += 3
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                End Try
            Case 3
                RollTheClip("c:\soundboard\cheryl\INTRO\CHERYLCALLING.wav")       '
                CurrentQ += 1
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
            Case 4                                                                  'INTRO
                RollTheClip("c:\soundboard\cheryl\INTRO\INTRO.wav")
                CurrentQ = CurrentQ + 1
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
            Case 5                                                                                  'WHO IS THE AUTO INSURANCE COMPANY YOURE WITH RIGHT NOW
                txtInsuranceProvider.Focus()
                CurrentQ = CurrentQ + 1

                RollTheClip("c:\soundboard\cheryl\INSURANCE INFO\WHO Is THE CURRENT AUTO INSURANCE COMPANY THAT YOU'RE WITH.wav")
                TControlQuestions.SelectedIndex = 1
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
            Case 6                                                                                                                                          'IF NONE, SKIPS TO VEHICLE, OTHERWISE ASKS FOR EXPIRATION AND START DATES
                If txtInsuranceProvider.Text.Length > 0 And txtInsuranceProvider.Text <> "none" And txtInsuranceProvider.Text <> "no insurance" Then
                    CurrentQ = CurrentQ + 1
                    RollTheClip("c:\soundboard\cheryl\INSURANCE INFO\EXPIRATION.WAV")
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                    txtPolicyExpiration.Focus()
                Else
                    CurrentQ = CurrentQ + 3
                    RESULTS.Text = RESULTS.Text & vbNewLine & " HAS NO INSURANCE"
                    RollTheClip("c:\soundboard\cheryl\VEHICLE INFO\YEAR OF VEHICLE.WAV")
                    txtVehicle.Focus()
                    lblQUESTION.Text = "YEAR OF VEHICLE"
                End If

            Case 7                                                                                                          'START 
                CurrentQ = CurrentQ + 1

                RollTheClip("c:\soundboard\cheryl\INSURANCE INFO\HOW MANY YEARS HAVE YOU BEEN WITH THEM 2.wav")

                txtPolicyStart.Focus()
                lblQUESTION.Text = "YEAR OF VEHICLE:"

            Case 8
                CurrentQ = CurrentQ + 2

                RollTheClip("c:\soundboard\cheryl\VEHICLE INFO\year of vehicle.wav")

                txtVehicle.Focus()


                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)

            Case 9



                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
            Case 10

                RollTheClip("c:\soundboard\cheryl\VEHICLE INFO\OTHER VEHICLES ON THAT POLICY.WAV")
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                CurrentQ += 1
                txtResults.Text += vbNewLine & "VEHICLE: " & txtVehicle.Text
            Case 11
                If cmbMoreVehicles.Text = "YES" Then
                    CurrentQ = 9
                    GoTo Redo
                Else
                    CurrentQ += 1
                    GoTo Redo
                End If

            Case 12
                RollTheClip("c:\soundboard\cheryl\DRIVER INFO\JUST ABOUT DONE DOB.WAV")
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)

            Case 13
                CurrentQ = CurrentQ + 1
                RollTheClip("c:\soundboard\cheryl\DRIVER INFO\MARITAL STATUS.WAV")
                cmbMaritalStatus.Focus()
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
            Case 14
                If cmbMaritalStatus.Text = "Married" Then
                    CurrentQ = CurrentQ + 1
                    RollTheClip("c:\soundboard\cheryl\DRIVER INFO\SPOUSES FIRST NAME.WAV")
                    txtSpouseName.Focus()
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                Else
                    CurrentQ = CurrentQ + 3
                    txtResults.Text += vbNewLine & "DOB: " & txtDOB.Text & vbNewLine & "MARITAL STATUS: " & "Single"
                    RollTheClip("c:\soundboard\cheryl\PERSONAL INFO\DO YOU OWN OR RENT THE HOME.WAV")
                    cmbOwnRent.Focus()
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)

                End If
            Case 15
                CurrentQ = CurrentQ + 1

                RollTheClip("c:\soundboard\cheryl\DRIVER INFO\SPOUSES DATE OF BIRTH.WAV")
                txtSpouseDOB.Focus()

                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)

            Case 16
                CurrentQ = CurrentQ + 1

                RESULTS.Text += vbNewLine & "DOB: " & txtDOB.Text & vbNewLine & "MARITAL STATUS: " & cmbMaritalStatus.Text
                RESULTS.Text += vbNewLine & "SPOUSE' NAME: " & txtSpouseName.Text
                RESULTS.Text += vbNewLine & "SPOUSE' DOB: " & txtSpouseDOB.Text


                RollTheClip("c:\soundboard\cheryl\PERSONAL INFO\DO YOU OWN OR RENT THE HOME.WAV")
                cmbOwnRent.Focus()
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)


            Case 17

                If cmbOwnRent.Text <> "OTHER" Then
                    CurrentQ = CurrentQ + 1
                    RollTheClip("c:\soundboard\cheryl\PERSONAL INFO\HOMETYPE.WAV")
                    cmbHomeType.Focus()
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                Else
                    CurrentQ = CurrentQ + 2
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\could you please verify your address.wav")
                    RESULTS.Text = RESULTS.Text & vbNewLine & "RENTS HOME"
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                    txtAddress.Focus()
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                End If


            Case 18
                CurrentQ = CurrentQ + 1

                RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\Could you please verify your address.WAV")
                txtAddress.Focus()
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)

            Case 19
                CurrentQ = CurrentQ + 1

                RollTheClip("C:\SoundBoard\Cheryl\PERSONAL INFO\email.wav")
                txtEmail.Focus()
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)


            Case 20
                CurrentQ = CurrentQ + 1

                cmbCredit.Focus()
                RollTheClip("c:\soundboard\cheryl\PERSONAL INFO\cmbCREDIT.WAV")
                cmbCredit.Focus()

                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
            Case 21
                CurrentQ = CurrentQ + 1

                RollTheClip("c:\soundboard\cheryl\PERSONAL INFO\cmbPhoneType.WAV")
                cmbPhoneType.Focus()
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
            Case 22
                CurrentQ = CurrentQ + 1

                RollTheClip("c:\soundboard\cheryl\PERSONAL INFO\LAST NAME.WAV")
                txtName.Focus()
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                AgeFromProg()
            Case 23
                CurrentQ = CurrentQ + 1
                RESULTS.Text += vbNewLine & "ADDRESS: " & txtAddress.Text
                RESULTS.Text += vbNewLine & "EMAIL: " & txtEmail.Text
                RESULTS.Text += vbNewLine & "cmbCREDIT: " & cmbCredit.Text
                RESULTS.Text += vbNewLine & "PHONE TYPE : " & cmbPhoneType.Text
                RESULTS.Text += vbNewLine & "FIRST/LAST NAME: " & txtName.Text
                WrapUp.Visible = True
                If chkLife.Visible = True And chkHome.Visible = True And chkMedicare.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Home life medicare.wav")
                ElseIf chkLife.Visible = True And chkHome.Visible = True And chkHealth.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Home Life and Health.wav")
                ElseIf chkLife.Visible = False And chkHome.Visible = True And chkMedicare.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Life and Medicare.wav")
                ElseIf chkLife.Visible = False And chkHome.Visible = True And chkHealth.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\ Life and Health.wav")
                ElseIf chkLife.Visible = True And chkRenters.Visible = True And chkMedicare.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Rental Life and Medicare.wav")
                ElseIf chkLife.Enabled = True And chkRenters.Visible = True And chkHealth.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Renters Health and Life.wav")
                ElseIf chkLife.Enabled = False And chkRenters.Visible = True And chkMedicare.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Renteral and Medicare.wav")
                ElseIf chkLife.Enabled = False And chkRenters.Visible = True And chkHealth.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Renters Health.wav")
                ElseIf chkLife.Visible = True And chkHome.Visible = False And chkRenters.Visible = False And chkMedicare.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Life and Medicare.wav")
                ElseIf chkLife.Visible = True And chkHome.Visible = False And chkRenters.Visible = False And chkHealth.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Life and Health.wav")
                ElseIf chkLife.Visible = False And chkRenters.Visible = False And chkHome.Visible = False And chkMedicare.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Medicare.wav")
                ElseIf chkLife.Visible = False And chkRenters.Visible = False And chkHome.Visible = False And chkHealth.Visible = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Health.wav")
                End If
                lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)


            Case 24
                cmbSecondaries.Focus()

                If SECONDARIES = True Then
                    RollTheClip("C:\SoundBoard\Cheryl\WRAPUP\Which Secondaries.wav")
                    CurrentQ = CurrentQ + 1
                Else
                    cmbTCPA.Focus()
                    RollTheClip("c:\soundboard\cheryl\WRAPUP\TCPA.WAV")
                    CurrentQ = CurrentQ + 4
                End If
            Case 25


                If SECONDARIES = True Then
                    If chkHealth.Checked Then
                        RESULTS.Text += "HOMEOWNERS INSURANCE"
                    ElseIf chkMedicare.Checked = True Then
                        RESULTS.Text += "MEDICARE INSURANCE"
                    End If
                    If chkLife.Checked Then
                        RESULTS.Text += "MEDICARE INSURANCE"
                    End If

                    If chkHome.Checked = True Then
                        txtYearBuilt.Visible = True
                        txtYearBuilt.Visible = True

                        RollTheClip("c:\soundboard\cheryl\WRAPUP\txtYearBuilt.WAV")
                        txtYearBuilt.Focus()
                        CurrentQ = CurrentQ + 1
                        lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                        RESULTS.Text += "HOMEOWNERS INSURANCE"

                    ElseIf chkRenters.Checked = True Then
                        RollTheClip("c:\soundboard\cheryl\WRAPUP\txtYearBuilt.WAV")
                        txtYearBuilt.Focus()
                        CurrentQ = CurrentQ + 1
                        lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                    Else
                        RollTheClip("c:\soundboard\cheryl\WRAPUP\TCPA.WAV")
                        CurrentQ = CurrentQ + 3
                        lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)

                    End If

                End If
            Case 26
                If chkHome.Checked = True Then

                    RollTheClip("c:\soundboard\cheryl\WRAPUP\SQUARE FOOTAGE.WAV")
                    txtSquareFeet.Focus()
                    CurrentQ = CurrentQ + 1
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                ElseIf chkRenters.Checked = True Then
                    RollTheClip("c:\soundboard\cheryl\WRAPUP\PPCoverage.WAV")
                    txtSquareFeet.Focus()
                    CurrentQ = CurrentQ + 1
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                Else
                    RollTheClip("c:\soundboard\cheryl\WRAPUP\TCPA.WAV")
                    CurrentQ = CurrentQ + 2
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                End If

            Case 27

                RollTheClip("c:\soundboard\cheryl\WRAPUP\TCPA.WAV")
                CurrentQ += 1
            Case 28


                If Strings.LCase(cmbTCPA.Text) = "YES" Then
                    RollTheClip("c:\soundboard\cheryl\WRAPUP\ENDCALL.WAV")
                    CurrentQ = 1
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                    totalLeads += 1
                    totalCalls += 1
                    txtInsuranceProvider.Clear()
                    txtPolicyExpiration.Clear()

                    txtVehicle.Clear()

                    txtDOB.Clear()

                    txtSpouseName.Clear()
                    txtSpouseDOB.Clear()


                    txtAddress.Clear()

                    txtName.Clear()

                Else
                    RollTheClip("c:\soundboard\cheryl\WRAPUP\HAVEAGOODDAY.WAV")
                    CurrentQ = 1
                    lblQUESTION.Text = CURRENTlblQuestion(CurrentQ)
                    totalCalls += 1

                    txtInsuranceProvider.Clear()
                    txtPolicyExpiration.Clear()

                    txtVehicle.Clear()

                    txtDOB.Clear()

                    txtSpouseName.Clear()
                    txtSpouseDOB.Clear()

                    txtAddress.Clear()

                    txtName.Clear()

                End If

        End Select


    End Sub

    Dim LifeQual As Boolean
    Dim HealthQual As Boolean
    Dim MediQual As Boolean

    Public Sub AgeFromProg()
        Dim dates() As String = txtDOB.Text.Split("/")
        Dim theMonth As Integer = CInt(dates(0))

        Dim theDay As Integer = CInt(dates(1))
        Dim theYear As Integer = CInt("19" & dates(2))
        Dim theAge As Integer = Today.Year - theYear
        If theMonth > Today.Month Then
            theAge -= 1
        End If
        Console.WriteLine("CUSTOMER IS " & theAge & " YEARS OLD")
        If theAge <= 80 And theAge >= 25 Then
            LifeQual = True
            chkLife.Visible = True
        Else
            LifeQual = False
            chkLife.Visible = False

        End If
        If theAge >= 64 Then
            chkHealth.Visible = False
            HealthQual = False
            MediQual = True
            chkMedicare.Visible = True
        Else
            HealthQual = True
            MediQual = False
            chkMedicare.Visible = False
            chkHealth.Visible = True

        End If
    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)            'HANDLES GLOBAL KEYPRESS
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                Case "170"

            End Select
        End If

    End Sub
    Private Sub VERSIONToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub OldYellowToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim f1 As New Form1
        f1.Show()
        Me.Close()


    End Sub

    Private Sub NewestToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form2_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub

    Private Sub Button59_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button56_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button55_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button31_Click_2(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button49_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles btnName.Click
        Try
            RollTheClip(globalFile2)
        Catch
        End Try
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles btnIntro.Click
        RollTheClip("c:\soundboard\cheryl\INTRO\INTRO.wav")
    End Sub

    Private Sub ButtonStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

    End Sub

    Private Sub txtPolicyExpiration_TextChanged(sender As Object, e As EventArgs) Handles txtPolicyStart.TextChanged

    End Sub

    Private Sub PersonalInfo_Click(sender As Object, e As EventArgs) Handles PersonalInfo.Click

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbOwnRent.SelectedIndexChanged

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSecondaries.SelectedIndexChanged

    End Sub

    Private Sub btnHello_Click(sender As Object, e As EventArgs) Handles btnHello.Click
        RollTheClip("c:\soundboard\cheryl\INTRO\HELLO.wav")
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        RollTheClip("C:\Soundboard\Cheryl\REACTIONS\OK.wav")
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles btnExcellent.Click
        RollTheClip("C:\Soundboard\Cheryl\REACTIONS\EXCELLENT 2.wav")
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles btnGreat.Click
        RollTheClip("C:\Soundboard\Cheryl\REACTIONS\GREAT 2.wav")
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles btnWonderful.Click
        RollTheClip("c:\soundboard\cheryl\REACTIONS\Wonderful.wav")
    End Sub

    Private Sub Button93_Click(sender As Object, e As EventArgs) Handles btnUnderstand.Click
        RollTheClip("c:\soundboard\cheryl\REBUTTALS\I COMPLETELY UNDERSTAND.WAV")
    End Sub

    Private Sub Button46_Click(sender As Object, e As EventArgs) Handles btnGreatQuestion.Click
        RollTheClip("c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT lblQuestion.WAV")
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles btnSpellThat.Click
        RollTheClip("C:\Soundboard\Cheryl\TIE INS\Could You Please Spell That Out.wav")
    End Sub

    Private Sub Button90_Click(sender As Object, e As EventArgs) Handles bthnHelo.Click
        RollTheClip("c:\soundboard\cheryl\INTRO\HELLO.wav")
    End Sub

    Private Sub VehicleInfo_Click(sender As Object, e As EventArgs) Handles VehicleInfo.Click

    End Sub
End Class