Option Strict Off
Option Explicit On
Friend Class FrmDebug
	Inherits System.Windows.Forms.Form
	Dim RCVBuffer As Object
	Private Sub FrmDebug_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MSComm1.Open()
    End Sub

    Private Sub MSComm1_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles MSComm1.DataReceived
        Dim TableArticle As Object
        Dim FileNum As Short
		Dim PSNData As String
		'UPGRADE_WARNING: Couldn't resolve default property of object RCVBuffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RCVBuffer = RCVBuffer & MSComm1.ReadExisting()
		'UPGRADE_WARNING: Couldn't resolve default property of object RCVBuffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If InStr(1, RCVBuffer, vbCrLf) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object RCVBuffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RCVBuffer = Mid(RCVBuffer, 1, InStr(1, RCVBuffer, vbCr) - 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object RCVBuffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(INIPSNFOLDERPATH & RCVBuffer & ".Txt") = "" Then
				MsgBox("PSN does not exist")
				'UPGRADE_WARNING: Couldn't resolve default property of object RCVBuffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RCVBuffer = ""
				Exit Sub
			End If

			FileNum = FreeFile()
			'UPGRADE_WARNING: Couldn't resolve default property of object TableArticle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(FileNum, INIPSNFOLDERPATH & TableArticle & ".txt", OpenMode.Input)
			Do While Not EOF(FileNum)
				PSNData = LineInput(FileNum)
				FrmDisplay.Text1.Text = FrmDisplay.Text1.Text & PSNData & vbCrLf
			Loop
			FileClose(FileNum)
			'UPGRADE_WARNING: Couldn't resolve default property of object RCVBuffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RCVBuffer = ""

		End If
	End Sub
End Class