Option Strict Off
Option Explicit On
Friend Class FrmMsg
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim RCVBuffer As Object
		Dim Action As Object
		'FrmMain.Txt_WONOS = ""
		'FrmMain.Txt_WOMODEL = ""
		''FrmMain.Txt_WOQTY = ""
		'FrmMain.Txt_CSUnit = "" '
		'UPGRADE_WARNING: Couldn't resolve default property of object Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Action = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object RCVBuffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RCVBuffer = ""
		Me.Close()
	End Sub
End Class