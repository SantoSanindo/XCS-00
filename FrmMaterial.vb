Public Class FrmMaterial
    Private Sub cmdback_Click(sender As Object, e As EventArgs) Handles cmdback.Click
        frmMain.Timer1.Enabled = True
        frmMain.Show()
        My.Forms.FrmMaterial.Dispose()
    End Sub

    Private Sub FrmMaterial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadrefCombo()
        LoadStnCombo()
    End Sub

    Private Sub LoadRefCombo()
        Dim query As String = "SELECT [ModelName] FROM [Parameter]"
        Dim ds As DataTable

        ds = ConnectionDatabase.readData(query).Tables(0)
        If ds.Rows.Count > 0 Then
            For index As Integer = 0 To ds.Rows.Count - 1
                Combo12.Items.Add(ds.Rows(index).Item("ModelName"))
            Next
        End If
    End Sub

    Private Sub LoadStnCombo()
        'Combo2.AddItem "SubAssy1"
        Combo2.Items.Add("Station1")
        Combo2.Items.Add("Station2")
        Combo2.Items.Add("Station3")
        Combo2.Items.Add("Station4")
        Combo2.Items.Add("Station5")
        Combo2.Items.Add("Station6")
        Combo2.Items.Add("SA_Connector")
    End Sub


    Private Sub Combo2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo2.SelectedIndexChanged
        Dim cbo() As ComboBox = {combo1_1, combo1_2, combo1_3, combo1_4, combo1_5, combo1_6, combo1_7, combo1_8, combo1_9, combo1_10, combo1_11, combo1_12, combo1_13, combo1_14, combo1_15, combo1_16, combo1_17, combo1_18, combo1_19, combo1_20, combo1_21, combo1_22, combo1_23, combo1_24, combo1_25, combo1_26, combo1_27, combo1_28, combo1_29, combo1_30, combo1_31, combo1_32, combo1_33, combo1_34, combo1_35, combo1_36, combo1_37, combo1_38, combo1_39, combo1_40, combo1_41, vvvvvvvvvvvv, combo1_43, combo1_44, combo1_45}
        Dim FNum As Integer
        Dim pos1, pos2 As Integer
        Dim Linestr As String
        Dim a As Integer
        Combo12.Text = ""
        'Debug.WriteLine(Combo12.Items.Count)
        For i As Integer = 1 To 45
            cbo(i - 1).Items.Clear()
        Next

        For i As Integer = 1 To 45
            FNum = FreeFile()

            FileOpen(FNum, INIMATERIALPATH & "Rack\" & Combo2.Text, OpenMode.Input)
            Do While Not EOF(FNum)
                Linestr = LineInput(FNum)
                cbo(i - 1).Items.Add(Linestr)
            Loop
            FileClose(FNum)
        Next
        a = 1
        FNum = FreeFile()
        FileOpen(FNum, INIMATERIALPATH & "Rack\" & Combo2.Text, OpenMode.Input)
        Do While Not EOF(FNum)
            Linestr = LineInput(FNum)
            cbo(a - 1).Text = Linestr
            a = a + 1
        Loop
        FileClose(FNum)
        Debug.WriteLine(cbo(0).Items.Count)
    End Sub

    Private Sub Combo12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo12.SelectedIndexChanged
        Dim cbo3() As ComboBox = {combo3_1, combo3_2, combo3_3, combo3_4, combo3_5, combo3_6, combo3_7, combo3_8, combo3_9, combo3_10, combo3_11, combo3_12, combo3_13, combo3_14, combo3_15, combo3_16, combo3_17, combo3_18, combo3_19, combo3_20, combo3_21, combo3_22, combo3_23, combo3_24, combo3_25, combo3_26, combo3_27, combo3_28, combo3_29, combo3_30, combo3_31, combo3_32, combo3_33, combo3_34, combo3_35, combo3_36, combo3_37, combo3_38, combo3_39, combo3_40, combo3_41, combo3_42, combo3_43, combo3_44, combo3_45}
        Dim cbo1() As ComboBox = {combo1_1, combo1_2, combo1_3, combo1_4, combo1_5, combo1_6, combo1_7, combo1_8, combo1_9, combo1_10, combo1_11, combo1_12, combo1_13, combo1_14, combo1_15, combo1_16, combo1_17, combo1_18, combo1_19, combo1_20, combo1_21, combo1_22, combo1_23, combo1_24, combo1_25, combo1_26, combo1_27, combo1_28, combo1_29, combo1_30, combo1_31, combo1_32, combo1_33, combo1_34, combo1_35, combo1_36, combo1_37, combo1_38, combo1_39, combo1_40, combo1_41, vvvvvvvvvvvv, combo1_43, combo1_44, combo1_45}
        Dim FNum As Integer
        Dim Linestr As String
        Dim a As Integer
        'Exit Sub
        If Dir(INIMATERIALPATH & Combo2.Text & "\" & Combo12.Text & ".Txt") = "" Then
            MsgBox("Unable to locate Material file")
            Exit Sub
        End If

        For i As Integer = 0 To 44
            cbo3(i).Text = "0"
        Next
        FNum = FreeFile()
        'Open INIMATERIALPATH & Combo2.Text & "\" & Combo12.Text & ".Txt" For Input As FNum
        FileOpen(FNum, INIMATERIALPATH & Combo2.Text & "\" & Combo12.Text & ".Txt", OpenMode.Input)
        Do While Not EOF(FNum)
            'Line Input #FNum, Linestr
            Linestr = LineInput(FNum)
            For i As Integer = 0 To 44
                If Linestr = cbo1(i).Text Then
                    cbo3(i).Text = 1
                End If

            Next
        Loop
        FileClose(FNum)
    End Sub

    Private Sub Cmd_Save_Click(sender As Object, e As EventArgs) Handles Cmd_Save.Click
        Dim cbo3() As ComboBox = {combo3_1, combo3_2, combo3_3, combo3_4, combo3_5, combo3_6, combo3_7, combo3_8, combo3_9, combo3_10, combo3_11, combo3_12, combo3_13, combo3_14, combo3_15, combo3_16, combo3_17, combo3_18, combo3_19, combo3_20, combo3_21, combo3_22, combo3_23, combo3_24, combo3_25, combo3_26, combo3_27, combo3_28, combo3_29, combo3_30, combo3_31, combo3_32, combo3_33, combo3_34, combo3_35, combo3_36, combo3_37, combo3_38, combo3_39, combo3_40, combo3_41, combo3_42, combo3_43, combo3_44, combo3_45}
        Dim cbo1() As ComboBox = {combo1_1, combo1_2, combo1_3, combo1_4, combo1_5, combo1_6, combo1_7, combo1_8, combo1_9, combo1_10, combo1_11, combo1_12, combo1_13, combo1_14, combo1_15, combo1_16, combo1_17, combo1_18, combo1_19, combo1_20, combo1_21, combo1_22, combo1_23, combo1_24, combo1_25, combo1_26, combo1_27, combo1_28, combo1_29, combo1_30, combo1_31, combo1_32, combo1_33, combo1_34, combo1_35, combo1_36, combo1_37, combo1_38, combo1_39, combo1_40, combo1_41, vvvvvvvvvvvv, combo1_43, combo1_44, combo1_45}

        Dim Fieldloc As Integer
        Dim FNum As Integer
        FNum = FreeFile()
        'Open INIMATERIALPATH & Combo2.Text & "\" & Combo12.Text & ".Txt" For Output As FNum
        FileOpen(FNum, INIMATERIALPATH & Combo2.Text & "\" & Combo12.Text & ".Txt", OpenMode.Output)
        For i As Integer = 0 To 44
            If cbo3(i).Text = "1" Then
                If cbo1(i).Text <> "" Then
                    PrintLine(FNum, cbo1(i).Text)
                End If
            End If
        Next
        FileClose(FNum)
    End Sub
End Class