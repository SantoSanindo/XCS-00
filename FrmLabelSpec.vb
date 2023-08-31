Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class FrmLabelSpec
	Inherits System.Windows.Forms.Form

	Private Sub ReloadCombo()
		Dim query As String
		Dim ds As DataTable

		query = "SELECT [ModelName] FROM [Parameter]"
		ds = ConnectionDatabase.readData(query).Tables(0)
		If ds.Rows.Count > 0 Then
			For index As Integer = 0 To ds.Rows.Count - 1
				Combo1.Items.Add(ds.Rows(index).Item("ModelName"))
			Next
		End If
	End Sub
	Private Sub ClearCombo()
		Combo1.Items.Clear()
	End Sub
	Private Sub FrmLabelSpec_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		ReloadCombo()
	End Sub

	Private Sub Cmd_Back_Click(sender As Object, e As EventArgs) Handles Cmd_Back.Click
		BusyFlag = False
		frmMain.Timer1.Enabled = True
		frmMain.Show()
		My.Forms.FrmLabelSpec.Dispose()
	End Sub

	Private Sub ClearScreen()
		Dim clean As Control
		Text1.Text = ""
		_Text3_1.Text = ""
		_Text3_2.Text = ""
		_Text3_3.Text = ""

		For j As Integer = 1 To 3
			clean = Me.Controls("_Text6_" & j)
			clean.Text = ""
		Next
		For k As Integer = 1 To 5
			clean = Me.Controls("_Text7_" & k)
			clean.Text = ""
		Next

		For l As Integer = 1 To 4
			clean = Me.Controls("_Text8_" & l)
			clean.Text = ""
		Next
		'UPGRADE_NOTE: Object Image1.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Image1.Image = Nothing
		'UPGRADE_NOTE: Object Image2.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Image2.Image = Nothing
		'UPGRADE_NOTE: Object Image3.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Image3.Image = Nothing
	End Sub
	Private Sub Combo1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo1.SelectedIndexChanged
		Dim clean As Control
		Dim query As String
		Dim ds As DataTable
		For i As Integer = 1 To 3
			clean = Me.Controls("_Text3_" & i)
			clean.Text = ""
			clean = Me.Controls("_Text6_" & i)
			clean.Text = ""
		Next
		For j As Integer = 0 To 7
			clean = Me.Controls("_Text2_" & j)
			clean.Text = ""
		Next
		For k As Integer = 1 To 5
			clean = Me.Controls("_Text7_" & k)
			clean.Text = ""
		Next
		For l As Integer = 1 To 4
			clean = Me.Controls("_Text8_" & l)
			clean.Text = ""
		Next
		For m As Integer = 0 To 2
			clean = Me.Controls("_Text4_" & m)
			clean.Text = ""
		Next
		For n As Integer = 1 To 2
			clean = Me.Controls("_Text9_" & n)
			clean.Text = ""
		Next
		'UPGRADE_NOTE: Object Image1.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Image1.Image = Nothing
		'UPGRADE_NOTE: Object Image2.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Image2.Image = Nothing
		'UPGRADE_NOTE: Object Image3.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Image3.Image = Nothing
		'UPGRADE_NOTE: Object Image4.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Image4.Image = Nothing
		'UPGRADE_NOTE: Object Image5.Picture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Image5.Image = Nothing

		On Error Resume Next ' GoTo ErrorHandler
		ClearScreen()

		query = "SELECT * FROM [Label] WHERE [ModelName] = '" & Combo1.Text & "'"
		ds = ConnectionDatabase.readData(query).Tables(0)
		_Text3_1.Text = ds.Rows(0).Item("Schematic_Reference")
		_Text3_2.Text = ds.Rows(0).Item("Schematic_Template")
		_Text3_3.Text = ds.Rows(0).Item("Schematic_Img")
		Image1.Image = System.Drawing.Image.FromFile(INILABELPHOTOPATH & _Text3_3.Text)
		_Text2_0.Text = ds.Rows(0).Item("Schematic_S4_1")
		_Text2_1.Text = ds.Rows(0).Item("Schematic_S4_2")
		_Text2_2.Text = ds.Rows(0).Item("Schematic_S5_1")
		_Text2_3.Text = ds.Rows(0).Item("Schematic_S5_2")
		_Text2_4.Text = ds.Rows(0).Item("Schematic_S6_1")
		_Text2_5.Text = ds.Rows(0).Item("Schematic_S6_2")
		_Text2_6.Text = ds.Rows(0).Item("Schematic_E2")
		_Text2_7.Text = ds.Rows(0).Item("Schematic_E1")

		_Text6_1.Text = ds.Rows(0).Item("Product_Reference")
		_Text6_2.Text = ds.Rows(0).Item("Product_Template")
		_Text6_3.Text = ds.Rows(0).Item("Product_Img")
		Image2.Image = System.Drawing.Image.FromFile(INILABELPHOTOPATH & _Text6_3.Text)

		_Text7_1.Text = ds.Rows(0).Item("Unitary_Reference")
		_Text7_2.Text = ds.Rows(0).Item("Unitary_Template")
		_Text7_3.Text = ds.Rows(0).Item("Unitary_Img")
		_Text7_4.Text = ds.Rows(0).Item("Unitary_Symbol")
		_Text7_5.Text = ds.Rows(0).Item("Unitary_Tension")
		Image3.Image = System.Drawing.Image.FromFile(INILABELPHOTOPATH & _Text7_3.Text)

		_Text8_1.Text = ds.Rows(0).Item("Group_Reference")
		_Text8_2.Text = ds.Rows(0).Item("Group_Template")
		_Text8_3.Text = ds.Rows(0).Item("Group_Img")
		_Text8_4.Text = ds.Rows(0).Item("Group_Qty")
		Image4.Image = System.Drawing.Image.FromFile(INILABELPHOTOPATH & _Text8_3.Text)

		_Text4_0.Text = ds.Rows(0).Item("Accessory1")
		_Text4_1.Text = ds.Rows(0).Item("Accessory2")
		_Text4_2.Text = ds.Rows(0).Item("InstructionSheet")

		_Text9_1.Text = ds.Rows(0).Item("Conn_Template")
		_Text9_2.Text = ds.Rows(0).Item("Conn_Img")
		Image5.Image = System.Drawing.Image.FromFile(INILABELPHOTOPATH & _Text9_2.Text)

		query = "SELECT * FROM [Parameter] WHERE [ModelName] = '" & Combo1.Text & "'"
		ds = ConnectionDatabase.readData(query).Tables(0)
		Text1.Text = ds.Rows(0).Item("ArticleNos")
		Exit Sub

ErrorHandler:
		MsgBox("Missing Parameters")
	End Sub

	Private Sub Cmd_Save_Click(sender As Object, e As EventArgs) Handles Cmd_Save.Click
		Dim query As String
		Dim ds As DataTable
		On Error Resume Next
		query = "UPDATE [Label] SET [ArticleNos] = '" & Text1.Text & "',
				[Schematic_Reference] = '" & _Text3_1.Text & "',
				[Schematic_Img] = '" & _Text3_3.Text & "',
				[Schematic_Template] = '" & _Text3_2.Text & "',
				[Schematic_S4_1] = '" & _Text2_0.Text & "',
				[Schematic_S4_2] = '" & _Text2_1.Text & "',
				[Schematic_S5_1] = '" & _Text2_2.Text & "',
				[Schematic_S5_2] = '" & _Text2_3.Text & "',
				[Schematic_S6_1] = '" & _Text2_4.Text & "',
				[Schematic_S6_2] = '" & _Text2_5.Text & "',
				[Schematic_E2] = '" & _Text2_6.Text & "',
				[Schematic_E1] = '" & _Text2_7.Text & "',
				[Product_Reference] = '" & _Text6_1.Text & "',
				[Product_Template] = '" & _Text6_2.Text & "',
				[Product_Img] = '" & _Text6_3.Text & "',
				[Unitary_Reference] = '" & _Text7_1.Text & "',
				[Unitary_Template] = '" & _Text7_2.Text & "',
				[Unitary_Img] = '" & _Text7_3.Text & "',
				[Unitary_Symbol] = '" & _Text7_4.Text & "',
				[Unitary_Tension] = '" & _Text7_5.Text & "',
				[Group_Reference] = '" & _Text8_1.Text & "',
				[Group_Template] = '" & _Text8_2.Text & "',
				[Group_Img] = '" & _Text8_3.Text & "',
				[Group_Qty] = '" & _Text8_4.Text & "',
				[Accessory1] = '" & _Text4_0.Text & "',
				[Accessory2] = '" & _Text4_1.Text & "',
				[InstructionSheet] = '" & _Text4_2.Text & "',
				[Conn_Template] = '" & _Text9_1.Text & "',
				[Conn_Img] = '" & _Text9_2.Text & "'
				WHERE [ModelName] = '" & Combo1.Text & "'"
		If ConnectionDatabase.updateData(query) Then
			MsgBox("SUCCESS UPDATE DATABASE")
		Else
			MsgBox("FAILED UPDATE DATABASE")
		End If
		Exit Sub

ErrorHandler:
		MsgBox("Unable to save parameter")
	End Sub
End Class