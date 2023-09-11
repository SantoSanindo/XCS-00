Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class FrmRack
	Inherits System.Windows.Forms.Form

	Private Sub cmd_back_Click(sender As Object, e As EventArgs) Handles cmd_back.Click
        frmMain.Show()
        My.Forms.FrmRack.Dispose()
    End Sub

    Private Sub cmd_save_Click(sender As Object, e As EventArgs) Handles cmd_save.Click
		Dim i As Integer
		Dim FNum As Object

		'UPGRADE_WARNING: Couldn't resolve default property of object FNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FNum = FreeFile()
		'UPGRADE_WARNING: Couldn't resolve default property of object FNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(FNum, INIMATERIALPATH & "Rack\" & Label3.Text, OpenMode.Output)
		For i = 1 To 45
			Dim control As Control
			control = Me.Controls("_Txt_slot_" & i)
			If Control.Text = "" Then Exit For
			'UPGRADE_WARNING: Couldn't resolve default property of object FNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(FNum, Control.Text)
		Next
		'UPGRADE_WARNING: Couldn't resolve default property of object FNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileClose(FNum)
	End Sub
End Class