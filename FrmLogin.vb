Option Strict Off
Option Explicit On

Friend Class FrmLogin
	Inherits System.Windows.Forms.Form
    Dim updateLetter As Boolean = False
    Dim updateNumber As Boolean = False
    Dim indexLetter As Integer
    Dim indexNumber As Integer
    Dim screenControl As Integer
    Dim statusLogin As Boolean

    Private Sub FrmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        screenControl = 0
        txt_password.Focus()
        tmrSequence.Enabled = True
    End Sub

    Private Sub btn_login_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        Dim queryID As String = "SELECT * FROM [USER] WHERE USER_ID='" & txt_username.Text & "'"
        Dim dtID As DataTable = ConnectionDatabase.readData(queryID).Tables(0)

        If dtID.Rows.Count > 0 Then
            Dim queryPSWD As String = "SELECT * FROM [USER] WHERE USER_PASSWORD='" & txt_password.Text & "'"
            Dim dtPSWD As DataTable = ConnectionDatabase.readData(queryPSWD).Tables(0)

            If dtPSWD.Rows.Count > 0 Then
                Debug.WriteLine("LOGIN BERHASIL!")
                'GVL.Login = True
                statusLogin = True

            Else
                Debug.WriteLine("LOGIN GAGAL!, USERNAME ATAU PASSWORD SALAH!")
                'GVL.Login = False
                statusLogin = False
            End If
        Else
            Debug.WriteLine("LOGIN GAGAL!, USERNAME ATAU PASSWORD SALAH!")
            'GVL.Login = False
            statusLogin = False
        End If
        Me.Dispose()
        frmMain.Show()
    End Sub

    Private Sub tmrSequence_Tick(sender As Object, e As EventArgs) Handles tmrSequence.Tick
        tmrSequence.Enabled = False
        If updateLetter Then
            If screenControl = 0 Then
                Select Case indexLetter
                    Case 0
                        txt_username.Text = txt_username.Text & "A"
                    Case 1
                        txt_username.Text = txt_username.Text & "B"
                    Case 2
                        txt_username.Text = txt_username.Text & "C"
                    Case 3
                        txt_username.Text = txt_username.Text & "D"
                    Case 4
                        txt_username.Text = txt_username.Text & "E"
                    Case 5
                        txt_username.Text = txt_username.Text & "F"
                    Case 6
                        txt_username.Text = txt_username.Text & "G"
                    Case 7
                        txt_username.Text = txt_username.Text & "H"
                    Case 8
                        txt_username.Text = txt_username.Text & "I"
                    Case 9
                        txt_username.Text = txt_username.Text & "J"
                    Case 10
                        txt_username.Text = txt_username.Text & "K"
                    Case 11
                        txt_username.Text = txt_username.Text & "L"
                    Case 12
                        txt_username.Text = txt_username.Text & "M"
                    Case 13
                        txt_username.Text = txt_username.Text & "N"
                    Case 14
                        txt_username.Text = txt_username.Text & "O"
                    Case 15
                        txt_username.Text = txt_username.Text & "P"
                    Case 16
                        txt_username.Text = txt_username.Text & "Q"
                    Case 17
                        txt_username.Text = txt_username.Text & "R"
                    Case 18
                        txt_username.Text = txt_username.Text & "S"
                    Case 19
                        txt_username.Text = txt_username.Text & "T"
                    Case 20
                        txt_username.Text = txt_username.Text & "U"
                    Case 21
                        txt_username.Text = txt_username.Text & "V"
                    Case 22
                        txt_username.Text = txt_username.Text & "W"
                    Case 23
                        txt_username.Text = txt_username.Text & "X"
                    Case 24
                        txt_username.Text = txt_username.Text & "Y"
                    Case 25
                        txt_username.Text = txt_username.Text & "Z"
                    Case 26
                        txt_username.Text = txt_username.Text & " "
                    Case 27
                        If txt_username.Text.Length > 0 Then
                            txt_username.Text = txt_username.Text.Remove(txt_username.Text.Length - 1, 1)
                        End If
                End Select
                txt_username.Focus()
            Else
                Select Case indexLetter
                    Case 0
                        txt_password.Text = txt_password.Text & "A"
                    Case 1
                        txt_password.Text = txt_password.Text & "B"
                    Case 2
                        txt_password.Text = txt_password.Text & "C"
                    Case 3
                        txt_password.Text = txt_password.Text & "D"
                    Case 4
                        txt_password.Text = txt_password.Text & "E"
                    Case 5
                        txt_password.Text = txt_password.Text & "F"
                    Case 6
                        txt_password.Text = txt_password.Text & "G"
                    Case 7
                        txt_password.Text = txt_password.Text & "H"
                    Case 8
                        txt_password.Text = txt_password.Text & "I"
                    Case 9
                        txt_password.Text = txt_password.Text & "J"
                    Case 10
                        txt_password.Text = txt_password.Text & "K"
                    Case 11
                        txt_password.Text = txt_password.Text & "L"
                    Case 12
                        txt_password.Text = txt_password.Text & "M"
                    Case 13
                        txt_password.Text = txt_password.Text & "N"
                    Case 14
                        txt_password.Text = txt_password.Text & "O"
                    Case 15
                        txt_password.Text = txt_password.Text & "P"
                    Case 16
                        txt_password.Text = txt_password.Text & "Q"
                    Case 17
                        txt_password.Text = txt_password.Text & "R"
                    Case 18
                        txt_password.Text = txt_password.Text & "S"
                    Case 19
                        txt_password.Text = txt_password.Text & "T"
                    Case 20
                        txt_password.Text = txt_password.Text & "U"
                    Case 21
                        txt_password.Text = txt_password.Text & "V"
                    Case 22
                        txt_password.Text = txt_password.Text & "W"
                    Case 23
                        txt_password.Text = txt_password.Text & "X"
                    Case 24
                        txt_password.Text = txt_password.Text & "Y"
                    Case 25
                        txt_password.Text = txt_password.Text & "Z"
                    Case 26
                        txt_password.Text = txt_password.Text & " "
                    Case 27
                        If txt_password.Text.Length > 0 Then
                            txt_password.Text = txt_password.Text.Remove(txt_password.Text.Length - 1, 1)
                        End If
                End Select
                txt_password.Focus()
            End If
            updateLetter = False
        End If

        If updateNumber Then
            If screenControl = 0 Then
                txt_username.Text = txt_username.Text & indexNumber
            Else
                txt_password.Text = txt_password.Text & indexNumber
            End If
            updateNumber = False
        End If
        tmrSequence.Enabled = True
    End Sub

    Private Sub _cmdNumeric_1_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_1.Click
        indexNumber = 1
        updateNumber = True
    End Sub
    Private Sub _cmdNumeric_2_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_2.Click
        indexNumber = 2
        updateNumber = True
    End Sub
    Private Sub _cmdNumeric_3_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_3.Click
        indexNumber = 3
        updateNumber = True
    End Sub
    Private Sub _cmdNumeric_4_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_4.Click
        indexNumber = 4
        updateNumber = True
    End Sub
    Private Sub _cmdNumeric_5_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_5.Click
        indexNumber = 5
        updateNumber = True
    End Sub
    Private Sub _cmdNumeric_6_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_6.Click
        indexNumber = 6
        updateNumber = True
    End Sub
    Private Sub _cmdNumeric_7_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_7.Click
        indexNumber = 7
        updateNumber = True
    End Sub
    Private Sub _cmdNumeric_8_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_8.Click
        indexNumber = 8
        updateNumber = True
    End Sub
    Private Sub _cmdNumeric_9_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_9.Click
        indexNumber = 9
        updateNumber = True
    End Sub

    Private Sub _cmdNumeric_0_Click(sender As Object, e As EventArgs) Handles _cmdNumeric_0.Click
        indexNumber = 0
        updateNumber = True
    End Sub

    Private Sub _cmdKeyboard_0_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_0.Click
        indexLetter = 0
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_1_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_1.Click
        indexLetter = 1
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_2_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_2.Click
        indexLetter = 2
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_3_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_3.Click
        indexLetter = 3
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_4_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_4.Click
        indexLetter = 4
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_5_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_5.Click
        indexLetter = 5
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_6_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_6.Click
        indexLetter = 6
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_7_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_7.Click
        indexLetter = 7
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_8_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_8.Click
        indexLetter = 8
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_9_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_9.Click
        indexLetter = 9
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_10_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_10.Click
        indexLetter = 10
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_11_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_11.Click
        indexLetter = 11
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_12_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_12.Click
        indexLetter = 12
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_13_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_13.Click
        indexLetter = 13
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_14_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_14.Click
        indexLetter = 14
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_15_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_15.Click
        indexLetter = 15
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_16_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_16.Click
        indexLetter = 16
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_17_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_17.Click
        indexLetter = 17
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_18_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_18.Click
        indexLetter = 18
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_19_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_19.Click
        indexLetter = 19
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_20_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_20.Click
        indexLetter = 20
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_21_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_21.Click
        indexLetter = 21
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_22_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_22.Click
        indexLetter = 22
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_23_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_23.Click
        indexLetter = 23
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_24_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_24.Click
        indexLetter = 24
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_25_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_25.Click
        indexLetter = 25
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_26_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_26.Click
        indexLetter = 26
        updateLetter = True
    End Sub
    Private Sub _cmdKeyboard_27_Click(sender As Object, e As EventArgs) Handles _cmdKeyboard_27.Click
        indexLetter = 27
        updateLetter = True
    End Sub

    Private Sub btn_enter_Click(sender As Object, e As EventArgs) Handles btn_enter.Click
        If screenControl = 0 Then
            txt_password.Focus()
            screenControl = 1
        Else
            txt_username.Focus()
            screenControl = 0
        End If
    End Sub

    Private Sub FrmLogin_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Login = statusLogin
    End Sub
End Class