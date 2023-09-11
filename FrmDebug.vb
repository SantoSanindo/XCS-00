Option Strict Off
Option Explicit On
Friend Class FrmDebug
	Inherits System.Windows.Forms.Form
	Dim RCVBuffer As Object
	Private Sub FrmDebug_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MSComm1.Open()
    End Sub

	Private Sub MSComm1_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles MSComm1.DataReceived
		Dim FileNum As Short
		Dim PSNData As String
		RCVBuffer = MSComm1.ReadExisting()
		If InStr(1, RCVBuffer, vbCrLf) <> 0 Then
			Me.Invoke(Sub()
						  RCVBuffer = Mid(RCVBuffer, 1, InStr(1, RCVBuffer, vbCr) - 1)
						  If Dir(INIPSNFOLDERPATH & RCVBuffer & ".Txt") = "" Then
							  MsgBox("PSN does not exist")
							  RCVBuffer = ""
							  Exit Sub
						  End If

						  FileNum = FreeFile()
						  FileOpen(FileNum, INIPSNFOLDERPATH & RCVBuffer & ".txt", OpenMode.Input)
						  Do While Not EOF(FileNum)
							  PSNData = LineInput(FileNum)
							  FrmDisplay.Text1.Text = FrmDisplay.Text1.Text & PSNData & vbCrLf
						  Loop
						  FileClose(FileNum)
						  RCVBuffer = ""
					  End Sub)
		End If
	End Sub
End Class