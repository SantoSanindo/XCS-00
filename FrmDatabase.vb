Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks

Friend Class FrmDatabase
    Inherits System.Windows.Forms.Form

    Private Sub ReloadCombo()
        Dim queryRef As String = "SELECT [MODELNAME] FROM [PARAMETER]"
        Dim dtRef As DataTable = ConnectionDatabase.readData(queryRef).Tables(0)

        If dtRef.Rows.Count > 0 Then
            For index As Integer = 0 To dtRef.Rows.Count - 1
                Combo1.Items.Add(dtRef.Rows(index).Item("ModelName"))
            Next
        End If
    End Sub

    Private Sub FrmDatabase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReloadCombo()
    End Sub

    Private Sub Combo1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo1.SelectedIndexChanged
        Dim clean As Control
        Dim query As String = "SELECT * FROM [Parameter] WHERE ModelName = '" & Combo1.Text & "'"
        Dim dt As DataTable = ConnectionDatabase.readData(query).Tables(0)

        For i As Integer = 1 To 16
            clean = Me.Controls("_Txt_ConnNos_" & i)
            clean.Text = Nothing
        Next
        For j As Integer = 1 To 16
            clean = Me.Controls("_Txt_ContactNos_" & j)
            clean.Text = Nothing
        Next
        For k As Integer = 1 To 10
            clean = Me.Controls("_ColorCombo_" & k)
            clean.Text = Nothing
        Next
        If dt.Rows.Count > 0 Then
            Txt_ArticleNos.Text = dt.Rows(0).Item("ArticleNos")
            Combo2.Text = dt.Rows(0).Item("MaterialType")
            Combo3.Text = dt.Rows(0).Item("BodyType")
            Combo4.Text = dt.Rows(0).Item("TensionType")
            _Text1_3.Text = dt.Rows(0).Item("ElectroMagnetType")
            _Text1_1.Text = dt.Rows(0).Item("PrimaryPCBAType")
            _Text1_2.Text = dt.Rows(0).Item("SecondaryPCBAType")
            _Combo5_1.Text = dt.Rows(0).Item("Contact1Type")
            _Combo5_2.Text = dt.Rows(0).Item("Contact2Type")
            _Combo5_3.Text = dt.Rows(0).Item("Contact3Type")
            _Combo5_4.Text = dt.Rows(0).Item("Contact4Type")
            _Combo5_5.Text = dt.Rows(0).Item("Contact5Type")
            _Combo5_6.Text = dt.Rows(0).Item("Contact6Type")
            _Combo6_1.Text = dt.Rows(0).Item("Contact1_W_Key")
            _Combo6_2.Text = dt.Rows(0).Item("Contact2_W_Key")
            _Combo6_3.Text = dt.Rows(0).Item("Contact3_W_Key")
            _Combo6_4.Text = dt.Rows(0).Item("Contact4_W_Key")
            _Combo6_5.Text = dt.Rows(0).Item("Contact5_W_Key")
            _Combo6_6.Text = dt.Rows(0).Item("Contact6_W_Key")
            _Combo7_1.Text = dt.Rows(0).Item("Contact1_W_Key_Ten")
            _Combo7_2.Text = dt.Rows(0).Item("Contact2_W_Key_Ten")
            _Combo7_3.Text = dt.Rows(0).Item("Contact3_W_Key_Ten")
            _Combo7_4.Text = dt.Rows(0).Item("Contact4_W_Key_Ten")
            _Combo7_5.Text = dt.Rows(0).Item("Contact5_W_Key_Ten")
            _Combo7_6.Text = dt.Rows(0).Item("Contact6_W_Key_Ten")
            Combo10.Text = dt.Rows(0).Item("FunctionType")
            Combo8.Text = dt.Rows(0).Item("ButtonOpt")
            Combo16.Text = dt.Rows(0).Item("ConnectorType")
            Combo9.Text = dt.Rows(0).Item("MechanismScrew")
            Combo11.Text = dt.Rows(0).Item("Head2BodyScrew")
            Combo12.Text = dt.Rows(0).Item("BouchouCap")
            _Text2_0.Text = dt.Rows(0).Item("SycLL")
            _Text2_1.Text = dt.Rows(0).Item("SycUL")

            _Txt_ContactNos_1.Text = dt.Rows(0).Item("S11")
            _Txt_ContactNos_2.Text = dt.Rows(0).Item("S12")
            _Txt_ContactNos_3.Text = dt.Rows(0).Item("S21")
            _Txt_ContactNos_4.Text = dt.Rows(0).Item("S22")
            _Txt_ContactNos_5.Text = dt.Rows(0).Item("S31")
            _Txt_ContactNos_6.Text = dt.Rows(0).Item("S32")
            _Txt_ContactNos_7.Text = dt.Rows(0).Item("S41")
            _Txt_ContactNos_8.Text = dt.Rows(0).Item("S42")
            _Txt_ContactNos_9.Text = dt.Rows(0).Item("S51")
            _Txt_ContactNos_10.Text = dt.Rows(0).Item("S52")
            _Txt_ContactNos_11.Text = dt.Rows(0).Item("S61")
            _Txt_ContactNos_12.Text = dt.Rows(0).Item("S62")

            _Txt_ConnNos_1.Text = dt.Rows(0).Item("S11PN")
            _Txt_ConnNos_2.Text = dt.Rows(0).Item("S12PN")
            _Txt_ConnNos_3.Text = dt.Rows(0).Item("S21PN")
            _Txt_ConnNos_4.Text = dt.Rows(0).Item("S22PN")
            _Txt_ConnNos_5.Text = dt.Rows(0).Item("S31PN")
            _Txt_ConnNos_6.Text = dt.Rows(0).Item("S32PN")
            _Txt_ConnNos_7.Text = dt.Rows(0).Item("S41PN")
            _Txt_ConnNos_8.Text = dt.Rows(0).Item("S42PN")
            _Txt_ConnNos_9.Text = dt.Rows(0).Item("S51PN")
            _Txt_ConnNos_10.Text = dt.Rows(0).Item("S52PN")
            _Txt_ConnNos_11.Text = dt.Rows(0).Item("S61PN")
            _Txt_ConnNos_12.Text = dt.Rows(0).Item("S62PN")
            _Txt_ConnNos_13.Text = dt.Rows(0).Item("X1PN")
            _Txt_ConnNos_14.Text = dt.Rows(0).Item("X2PN")
            _Txt_ConnNos_15.Text = dt.Rows(0).Item("E1PN")
            _Txt_ConnNos_16.Text = dt.Rows(0).Item("E2PN")

            _ColorCombo_1.Text = dt.Rows(0).Item("S1CC")
            _ColorCombo_2.Text = dt.Rows(0).Item("S2CC")
            _ColorCombo_3.Text = dt.Rows(0).Item("S3CC")
            _ColorCombo_4.Text = dt.Rows(0).Item("S4CC")
            _ColorCombo_5.Text = dt.Rows(0).Item("S5CC")
            _ColorCombo_6.Text = dt.Rows(0).Item("S6CC")
            _ColorCombo_7.Text = dt.Rows(0).Item("X1CC")
            _ColorCombo_8.Text = dt.Rows(0).Item("X2CC")
            _ColorCombo_9.Text = dt.Rows(0).Item("E1CC")
            _ColorCombo_10.Text = dt.Rows(0).Item("E2CC")

            Debug.WriteLine(Combo1.Items.Count.ToString())
        End If



    End Sub

    Private Sub Cmd_Add_Click(sender As Object, e As EventArgs) Handles Cmd_Add.Click
        Dim statQuery1 As Boolean
        Dim statQuery2 As Boolean
        Dim inString As String = InputBox("Enter new model Name", "Adding New Reference")
        Dim FNum As Integer
        Dim dbNull As String = ""

        If inString <> "" Then
            If RefCheck(inString) Then
                MsgBox("Model Name already exist in Database. Please use another.")
                Exit Sub
            End If

            Combo1.Text = inString
            Dim query1 As String = "INSERT INTO [Parameter] ([ModelName],[ArticleNos],[MaterialType],[BodyType],[ConnectorType],[TensionType],[ElectroMagnetType],
                                [PrimaryPCBAType],[SecondaryPCBAType],[Contact1Type],[Contact2Type],[Contact3Type],[Contact4Type],[Contact5Type],[Contact6Type],[Contact1_W_Key],
                                [Contact2_W_Key],[Contact3_W_Key],[Contact4_W_Key],[Contact5_W_Key],[Contact6_W_Key],
                                [Contact1_W_Key_Ten],[Contact2_W_Key_Ten],[Contact3_W_Key_Ten],[Contact4_W_Key_Ten],[Contact5_W_Key_Ten],[Contact6_W_Key_Ten],
                                [FunctionType],[ButtonOpt],[S11],[S12],[S21],[S22],[S31],[S32],[S41],[S42],[S51],[S52],[S61],[S62],
                                [S11PN],[S12PN],[S21PN],[S22PN],[S31PN],[S32PN],[S41PN],[S42PN],[S51PN],[S52PN],[S61PN],[S62PN],
                                [X1PN],[X2PN],[E1PN],[E2PN],[S1CC],[S2CC],[S3CC],[S4CC],[S5CC],[S6CC],[X1CC],[X2CC],[E1CC],[E2CC],
                                [MechanismScrew],[Head2BodyScrew],[BouchouCap],[ProductVer],[SycLL],[SycUL])
                                VALUES (ISNULL('" & inString & "','" & dbNull & "'),ISNULL('" & Txt_ArticleNos.Text & "','" & dbNull & "'),ISNULL('" & Combo2.Text & "','" & dbNull & "'),ISNULL('" & Combo3.Text & "','" & dbNull & "'),ISNULL('" & Combo16.Text & "','" & dbNull & "'),ISNULL('" & Combo4.Text & "','" & dbNull & "'),ISNULL('" & _Text1_3.Text & "','" & dbNull & "'),
                                ISNULL('" & _Text1_1.Text & "','" & dbNull & "'),ISNULL('" & _Text1_2.Text & "','" & dbNull & "'),ISNULL('" & _Combo5_1.Text & "','" & dbNull & "'),ISNULL('" & _Combo5_2.Text & "','" & dbNull & "'),ISNULL('" & _Combo5_3.Text & "','" & dbNull & "'),ISNULL('" & _Combo5_4.Text & "','" & dbNull & "'),ISNULL('" & _Combo5_5.Text & "','" & dbNull & "'),ISNULL('" & _Combo5_6.Text & "','" & dbNull & "'),
                                ISNULL('" & _Combo6_1.Text & "','" & dbNull & "'),ISNULL('" & _Combo6_2.Text & "','" & dbNull & "'),ISNULL('" & _Combo6_3.Text & "','" & dbNull & "'),ISNULL('" & _Combo6_4.Text & "','" & dbNull & "'),ISNULL('" & _Combo6_5.Text & "','" & dbNull & "'),ISNULL('" & _Combo6_6.Text & "','" & dbNull & "'),ISNULL('" & _Combo7_1.Text & "','" & dbNull & "'),ISNULL('" & _Combo7_2.Text & "','" & dbNull & "'),
                                ISNULL('" & _Combo7_3.Text & "','" & dbNull & "'),ISNULL('" & _Combo7_4.Text & "','" & dbNull & "'),ISNULL('" & _Combo7_5.Text & "','" & dbNull & "'),ISNULL('" & _Combo7_6.Text & "','" & dbNull & "'),ISNULL('" & Combo10.Text & "','" & dbNull & "'),ISNULL('" & Combo8.Text & "','" & dbNull & "'),
                                ISNULL('" & _Txt_ContactNos_1.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_2.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_3.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_4.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_5.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_6.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_7.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_8.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_9.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_10.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_11.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ContactNos_12.Text & "','" & dbNull & "'),
                                ISNULL('" & _Txt_ConnNos_1.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_2.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_3.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_4.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_5.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_6.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_7.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_8.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_9.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_10.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_11.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_12.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_13.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_14.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_15.Text & "','" & dbNull & "'),ISNULL('" & _Txt_ConnNos_16.Text & "','" & dbNull & "'),
                                ISNULL('" & _ColorCombo_1.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_2.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_3.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_4.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_5.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_6.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_7.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_8.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_9.Text & "','" & dbNull & "'),ISNULL('" & _ColorCombo_10.Text & "','" & dbNull & "'),
                                '','','','','','')"

            If ConnectionDatabase.insertData(query1) Then
                Dim query2 As String = "INSERT INTO [Label] ([ModelName],[ArticleNos],
                                   [Group_Qty],[Group_Template],[Group_Img],[Unitary_Template],[Unitary_Img],[Schematic_Reference],[Schematic_Img],[Schematic_Template],[Product_Reference],[Product_Img],[Unitary_Reference],[Product_Template],[Group_Reference],[Unitary_Symbol],[Unitary_Tension],
                                   [Schematic_S4_1],[Schematic_S4_2],[Schematic_S5_1],[Schematic_S5_2],[Schematic_S6_1],[Schematic_S6_2],[Schematic_E2],[Schematic_E1],[Accessory1],[Accessory2],[Accessory3],[InstructionSheet],[Conn_Template],[Conn_Img]) 
                                   VALUES (ISNULL('" & inString & "','" & dbNull & "'),'" & Txt_ArticleNos.Text & "',
                                   '','','','','','','','','','','','','','','',
                                   '','','','','','','','','','','','','','')"

                If ConnectionDatabase.insertData(query2) Then
                    MsgBox("Success add database!")
                Else
                    MsgBox("Failed add database!")
                End If
            Else
                MsgBox("Failed add database!")
            End If

            FNum = FreeFile()
            FileOpen(FNum, INIMATERIALPATH & "Station1\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)
            FileOpen(FNum, INIMATERIALPATH & "Station2\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)
            FileOpen(FNum, INIMATERIALPATH & "Station3\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)
            FileOpen(FNum, INIMATERIALPATH & "Station4\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)
            FileOpen(FNum, INIMATERIALPATH & "Station5\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)
            FileOpen(FNum, INIMATERIALPATH & "Station6\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)
            FileOpen(FNum, INIMATERIALPATH & "SubAssy1\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)
            FileOpen(FNum, INIMATERIALPATH & "SubAssy2\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)
            FileOpen(FNum, INIMATERIALPATH & "SubAssy3\" & inString & ".Txt", OpenMode.Output)
            FileClose(FNum)

            ClearCombo()
            ReloadCombo()
        End If
    End Sub

    Private Sub ClearCombo()
        Combo1.Items.Clear()
    End Sub
    Private Function RefCheck(ByRef strName As String) As Boolean
        Dim query As String = "SELECT * FROM [Parameter] WHERE ModelName ='" & strName & "'"
        Dim ds As DataTable = ConnectionDatabase.readData(query).Tables(0)

        If ds.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Cmd_Back_Click(sender As Object, e As EventArgs) Handles Cmd_Back.Click
        BusyFlag = False
        frmMain.Timer1.Enabled = True
        frmMain.Show()
        My.Forms.FrmDatabase.Dispose()
    End Sub

    Private Sub Cmd_Save_Click(sender As Object, e As EventArgs) Handles Cmd_Save.Click
        Dim statusQuery1 As Boolean
        Dim query1 As String = "UPDATE [Parameter] SET [ArticleNos] = '" & Txt_ArticleNos.Text & "',[MaterialType] = '" & Combo2.Text & "',[BodyType] = '" & Combo3.Text & "', [FunctionType] = '" & Combo10.Text & "',[ConnectorType] = '" & Combo16.Text & "', [TensionType] = '" & Combo4.Text & "', [ElectroMagnetType] = '" & _Text1_3.Text & "', [PrimaryPCBAType] = '" & _Text1_1.Text & "', [SecondaryPCBAType] = '" & _Text1_2.Text & "',[Contact1Type] = '" & _Combo5_1.Text & "', [Contact2Type] = '" & _Combo5_2.Text & "', [Contact3Type] = '" & _Combo5_3.Text & "', [Contact4Type] = '" & _Combo5_4.Text & "', [Contact5Type] = '" & _Combo5_5.Text & "', [Contact6Type] = '" & _Combo5_6.Text & "', [Contact1_W_Key] = '" & _Combo6_1.Text & "', [Contact2_W_Key] = '" & _Combo6_2.Text & "',  [Contact3_W_Key] = '" & _Combo6_3.Text & "', [Contact4_W_Key] = '" & _Combo6_4.Text & "', [Contact5_W_Key] = '" & _Combo6_5.Text & "', [Contact6_W_Key] = '" & _Combo6_6.Text & "', 
                                [Contact1_W_Key_Ten] = '" & _Combo7_1.Text & "', [Contact2_W_Key_Ten] = '" & _Combo7_2.Text & "', [Contact3_W_Key_Ten] = '" & _Combo7_3.Text & "', [Contact4_W_Key_Ten] = '" & _Combo7_4.Text & "', [Contact5_W_Key_Ten] = '" & _Combo7_5.Text & "', [Contact6_W_Key_Ten] = '" & _Combo7_6.Text & "', 
                                [ButtonOpt] = '" & Combo8.Text & "',
                                [MechanismScrew] = '" & Combo9.Text & "', [Head2BodyScrew] = '" & Combo11.Text & "',[BouchouCap] = '" & Combo12.Text & "', [SycLL] = '" & _Text2_0.Text & "', [SycUL] = '" & _Text2_1.Text & "',
                                [S11] = '" & _Txt_ContactNos_1.Text & "', [S12] = '" & _Txt_ContactNos_2.Text & "',[S21] = '" & _Txt_ContactNos_3.Text & "', [S22] = '" & _Txt_ContactNos_4.Text & "', [S31] = '" & _Txt_ContactNos_5.Text & "',[S32] = '" & _Txt_ContactNos_6.Text & "', [S41] = '" & _Txt_ContactNos_7.Text & "',[S42] = '" & _Txt_ContactNos_8.Text & "', [S51] = '" & _Txt_ContactNos_9.Text & "', [S52] = '" & _Txt_ContactNos_10.Text & "',[S61] = '" & _Txt_ContactNos_11.Text & "', [S62] = '" & _Txt_ContactNos_12.Text & "',
                                [S11PN] = '" & _Txt_ConnNos_1.Text & "', [S12PN] = '" & _Txt_ConnNos_2.Text & "',[S21PN] = '" & _Txt_ConnNos_3.Text & "', [S22PN] = '" & _Txt_ConnNos_4.Text & "', [S31PN] = '" & _Txt_ConnNos_5.Text & "',[S32PN] = '" & _Txt_ConnNos_6.Text & "', [S41PN] = '" & _Txt_ConnNos_7.Text & "',[S42PN] = '" & _Txt_ConnNos_8.Text & "', [S51PN] = '" & _Txt_ConnNos_9.Text & "', [S52PN] = '" & _Txt_ConnNos_10.Text & "',[S61PN] = '" & _Txt_ConnNos_11.Text & "', [S62PN] = '" & _Txt_ConnNos_12.Text & "',
                                [X1PN] = '" & _Txt_ConnNos_13.Text & "', [X2PN] = '" & _Txt_ConnNos_14.Text & "',[E1PN] = '" & _Txt_ConnNos_15.Text & "', [E2PN] = '" & _Txt_ConnNos_16.Text & "',
                                [S1CC] = '" & _ColorCombo_1.Text & "', [S2CC] = '" & _ColorCombo_2.Text & "',[S3CC] = '" & _ColorCombo_3.Text & "', [S4CC] = '" & _ColorCombo_4.Text & "', [S5CC] = '" & _ColorCombo_5.Text & "',[S6CC] = '" & _ColorCombo_6.Text & "', [X1CC] = '" & _ColorCombo_7.Text & "',[X2CC] = '" & _ColorCombo_8.Text & "', [E1CC] = '" & _ColorCombo_9.Text & "', [E2CC] = '" & _ColorCombo_10.Text & "' Where [ModelName] = '" & Combo1.Text & "'"
        If ConnectionDatabase.updateData(query1) Then
            MsgBox("Success save & update database!")
        Else
            MsgBox("Failed save & update database!")
        End If
    End Sub
End Class