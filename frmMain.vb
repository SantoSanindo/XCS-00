Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Text
Imports VB = Microsoft.VisualBasic
Public Class frmMain
    Dim xNew As Integer = 3
    Dim BusyFlag As Boolean
    Dim TimerCount As Integer
    Dim BusyTimerInProgress As Boolean
    Dim ChangeRecorded As Boolean
    Dim StationConnectionStatus(10) As Integer
    Dim SelectedMode As Integer

    Private Const ICMP_SUCCESS As Integer = 0
    Private Const ICMP_STATUS_BUFFER_TO_SMALL As Short = 11001 'Buffer Too Small
    Private Const ICMP_STATUS_DESTINATION_NET_UNREACH As Short = 11002 'Destination Net Unreachable
    Private Const ICMP_STATUS_DESTINATION_HOST_UNREACH As Short = 11003 'Destination Host Unreachable
    Private Const ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH As Short = 11004 'Destination Protocol Unreachable
    Private Const ICMP_STATUS_DESTINATION_PORT_UNREACH As Short = 11005 'Destination Port Unreachable
    Private Const ICMP_STATUS_NO_RESOURCE As Short = 11006 'No Resources
    Private Const ICMP_STATUS_BAD_OPTION As Short = 11007 'Bad Option
    Private Const ICMP_STATUS_HARDWARE_ERROR As Short = 11008 'Hardware Error
    Private Const ICMP_STATUS_LARGE_PACKET As Short = 11009 'Packet Too Big
    Private Const ICMP_STATUS_REQUEST_TIMED_OUT As Short = 11010 'Request Timed Out
    Private Const ICMP_STATUS_BAD_REQUEST As Short = 11011 'Bad Request
    Private Const ICMP_STATUS_BAD_ROUTE As Short = 11012 'Bad Route
    Private Const ICMP_STATUS_TTL_EXPIRED_TRANSIT As Short = 11013 'TimeToLive Expired Transit
    Private Const ICMP_STATUS_TTL_EXPIRED_REASSEMBLY As Short = 11014 'TimeToLive Expired Reassembly
    Private Const ICMP_STATUS_PARAMETER As Short = 11015 'Parameter Problem
    Private Const ICMP_STATUS_SOURCE_QUENCH As Short = 11016 'Source Quench
    Private Const ICMP_STATUS_OPTION_TOO_BIG As Short = 11017 'Option Too Big
    Private Const ICMP_STATUS_BAD_DESTINATION As Short = 11018 'Bad Destination
    Private Const ICMP_STATUS_NEGOTIATING_IPSEC As Short = 11032 'Negotiating IPSEC
    Private Const ICMP_STATUS_GENERAL_FAILURE As Short = 11050 'General Failure


    Private Const WINSOCK_ERROR As String = "Windows Sockets not responding correctly."
    Private Const INADDR_NONE As Integer = &HFFFFFFFF
    Private Const WSA_SUCCESS As Short = 0
    Private Const WS_VERSION_REQD As Integer = &H101

    'Clean up sockets.
    'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512

    Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Integer

    'Open the socket connection.
    'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
    'UPGRADE_WARNING: Structure WSADATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, ByRef lpWSADATA As WSADATA) As Integer

    'Create a handle on which Internet Control Message Protocol (ICMP) requests can be issued.
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpcreatefile.asp
    Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Integer

    'Convert a string that contains an (Ipv4) Internet Protocol dotted address into a correct address.
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winsock/wsapiref_4esy.asp
    Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Integer

    'Close an Internet Control Message Protocol (ICMP) handle that IcmpCreateFile opens.
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpclosehandle.asp

    Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Integer) As Integer

    'Information about the Windows Sockets implementation
    'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
    Private Structure WSADATA
        Dim wVersion As Short
        Dim wHighVersion As Short
        <VBFixedArray(256)> Dim szDescription() As Byte
        <VBFixedArray(128)> Dim szSystemStatus() As Byte
        Dim iMaxSockets As Integer
        Dim iMaxUDPDG As Integer
        Dim lpVendorInfo As Integer

        'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
        Public Sub Initialize()
            ReDim szDescription(256)
            ReDim szSystemStatus(128)
        End Sub
    End Structure

    'Send an Internet Control Message Protocol (ICMP) echo request, and then return one or more replies.
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIcmpSendEcho.asp
    'UPGRADE_WARNING: Structure ICMP_ECHO_REPLY may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Integer, ByVal DestinationAddress As Integer, ByVal RequestData As String, ByVal RequestSize As Integer, ByVal RequestOptions As Integer, ByRef ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Integer, ByVal Timeout As Integer) As Integer

    'This structure describes the options that will be included in the header of an IP packet.
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIP_OPTION_INFORMATION.asp
    Private Structure IP_OPTION_INFORMATION
        Dim Ttl As Byte
        Dim Tos As Byte
        Dim Flags As Byte
        Dim OptionsSize As Byte
        Dim OptionsData As Integer
    End Structure

    'This structure describes the data that is returned in response to an echo request.
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmp_echo_reply.asp
    Private Structure ICMP_ECHO_REPLY
        Dim address As Integer
        Dim Status As Integer
        Dim RoundTripTime As Integer
        Dim DataSize As Integer
        Dim Reserved As Short
        Dim ptrData As Integer
        Dim Options As IP_OPTION_INFORMATION
        'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
        <VBFixedString(250), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=250)> Public Data() As Char
    End Structure
    Dim CSAction As Integer 'For Change Series action
    Dim WOBuffer As String
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Frame3.Visible = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Frame3.Visible = False
    End Sub

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

    Private Sub LoadDistrupWOData()
        Dim N As Object
        Dim i As Object
        Dim query As String
        Dim ds As DataTable

        query = "SELECT * FROM [ONGOING]"
        ds = ConnectionDatabase.readData(query).Tables(0)
        ' Loop through the rows in the DataTable and add them to the DataGridView
        For rowIndex As Integer = 0 To ds.Rows.Count - 1
            For columnIndex As Integer = 0 To ds.Columns.Count - 1
                'Set the value of the cell in the DataGridView
                If ds.Rows(rowIndex)(columnIndex).ToString() IsNot Nothing Then
                    MSFlexGrid2.Rows(rowIndex).Cells(columnIndex).Value = ds.Rows(rowIndex)(columnIndex).ToString()
                End If
            Next
        Next
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim directoryPath As String
        Dim fullPath As String = System.AppDomain.CurrentDomain.BaseDirectory
        Dim projectFolder As String = fullPath.Replace("\bin\", "")
        Cmd_RFID.Enabled = False
        Cmd_Tag.Enabled = False
        Cmd_Parameters.Enabled = False
        Cmd_Parts.Enabled = False
        Cmd_Login.Text = "Login"

        TimerCount = 1
        BusyFlag = False
        BusyTimerInProgress = False
        ChangeRecorded = False


        If Dir(projectFolder & "\Config\Config.INI") = "" Then
            MsgBox("Config.INI is missing")
            End
        End If
        ReadINI(projectFolder & "\Config\Config.INI")
        directoryPath = INIPSNFOLDERPATH
        Dim files As String() = Directory.GetFiles(directoryPath)
        For Each file As String In files
            'Debug.WriteLine(file)
            File1.Items.Add(Path.GetFileName(file))
        Next

        CheckConnections(TimerCount)
        LoadPSNTable()
        LoadSTNTable()
        LoadWOTable()
        'RFID_Comm.Open()
        StationConnectionStatus(1) = -1
        StationConnectionStatus(2) = -1
        StationConnectionStatus(3) = -1
        StationConnectionStatus(4) = -1
        StationConnectionStatus(5) = -1
        StationConnectionStatus(6) = -1
        StationConnectionStatus(7) = -1
        StationConnectionStatus(8) = -1
        StationConnectionStatus(9) = -1
        StationConnectionStatus(10) = -1

        Frame2.Location = New Point(3, 3)
        Frame3.Location = New Point(20000, 120)
        Frame4.Location = New Point(20000, 120)
        Frame5.Location = New Point(20000, 120)
        Frame6.Location = New Point(20000, 120)
        Frame7.Location = New Point(20000, 120)
        Frame8.Location = New Point(20000, 120)
        Frame9.Location = New Point(20000, 120)
        'BarcodeScan_Comm.Open()
        SelectedMode = 2
        TimerCount = 3
    End Sub

    Private Sub LoadPSNTable()
        MSFlexGrid1.Columns.Add("Column1", "S/No")
        MSFlexGrid1.Columns.Add("Column2", "Product PSN")
        MSFlexGrid1.Columns.Add("Column3", "Main PCBA Assy")
        MSFlexGrid1.Columns.Add("Column4", "ElectroMagnet Assy")
        MSFlexGrid1.Columns.Add("Column5", "Body Assy")
        MSFlexGrid1.Columns.Add("Column6", "Head2Body Assy")
        MSFlexGrid1.Columns.Add("Column7", "Tester")
        MSFlexGrid1.Columns.Add("Column8", "Bouchou Assy")
        MSFlexGrid1.Columns.Add("Column9", "Cover Assy")
        MSFlexGrid1.Columns.Add("Column10", "IP67")
        MSFlexGrid1.Columns.Add("Column11", "Connector Test")
        MSFlexGrid1.Columns.Add("Column12", "IP66")
        MSFlexGrid1.Columns.Add("Column13", "Packaging")
        MSFlexGrid1.Columns.Add("Column14", "Date Packed")

        MSFlexGrid1.Columns("Column1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column3").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column4").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column5").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column6").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column7").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column8").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column9").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column10").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column11").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column12").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column13").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid1.Columns("Column14").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    Private Sub LoadDistrupTable()
        MSFlexGrid2.Columns.Add("Column0", "S/No")
        MSFlexGrid2.Columns.Add("Column1", "Date Created")
        MSFlexGrid2.Columns.Add("Column2", "WO Nos")
        MSFlexGrid2.Columns.Add("Column3", "Sub Assy1")
        MSFlexGrid2.Columns.Add("Column4", "Sub Assy2")
        MSFlexGrid2.Columns.Add("Column5", "Main PCBA Assy")
        MSFlexGrid2.Columns.Add("Column6", "ElectroMagnet Assy")
        MSFlexGrid2.Columns.Add("Column7", "Body Assy")
        MSFlexGrid2.Columns.Add("Column8", "Head2BodyAssy")
        MSFlexGrid2.Columns.Add("Column9", "Tester")
        MSFlexGrid2.Columns.Add("Column10", "Cover Assy")
        MSFlexGrid2.Columns.Add("Column11", "Packaging")
        MSFlexGrid2.Columns.Add("Column12", "Comments")
        MSFlexGrid2.Columns("Column0").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column3").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column4").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column5").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column6").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column7").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column8").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column9").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column10").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column11").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        MSFlexGrid2.Columns("Column12").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter


        For i As Integer = 1 To 99
            MSFlexGrid2.Rows.Add()
        Next
    End Sub
    Private Sub LoadSTNTable()
        ConnectTable.Columns.Add("Column1", "S/NO")
        For i As Integer = 1 To 10
            ConnectTable.Rows(ConnectTable.Rows.Add()).Cells("Column1").Value = i
        Next
        ConnectTable.Columns.Add("Column2", "Bench")
        ConnectTable.Columns.Add("Column3", "Description")
        ConnectTable.Rows(0).Cells(1).Value = "XCS - 7"
        ConnectTable.Rows(0).Cells(2).Value = "Sub Assy 1"
        ConnectTable.Rows(1).Cells(1).Value = "XCS - 3"
        ConnectTable.Rows(1).Cells(2).Value = "Sub Assy 2"
        ConnectTable.Rows(2).Cells(1).Value = "XCS - 8"
        ConnectTable.Rows(2).Cells(2).Value = "Main PCBA Assy"
        ConnectTable.Rows(3).Cells(1).Value = "XCS - 9"
        ConnectTable.Rows(3).Cells(2).Value = "ElectroMagnet Assy"
        ConnectTable.Rows(4).Cells(1).Value = "XCS - 11"
        ConnectTable.Rows(4).Cells(2).Value = "Body Assy"
        ConnectTable.Rows(5).Cells(1).Value = "XCS - 12"
        ConnectTable.Rows(5).Cells(2).Value = "Head2Body Assy"
        ConnectTable.Rows(6).Cells(1).Value = "XCS - 13"
        ConnectTable.Rows(6).Cells(2).Value = "Tester"
        ConnectTable.Rows(7).Cells(1).Value = "XCS - 15"
        ConnectTable.Rows(7).Cells(2).Value = "Station#5"
        ConnectTable.Rows(8).Cells(1).Value = "XCS - 17"
        ConnectTable.Rows(8).Cells(2).Value = "Packaging"
        ConnectTable.Rows(9).Cells(1).Value = "XCS - 18"
        ConnectTable.Rows(9).Cells(2).Value = "Connector Tester"

        ConnectTable.Columns.Add("Column4", "Connect")
        ConnectTable.Columns.Add("Column5", "WO Nos")
        ConnectTable.Columns.Add("Column6", "WO Model")
        ConnectTable.Columns.Add("Column7", "WO Qty")
        ConnectTable.Columns.Add("Column8", "Tag Nos")
        ConnectTable.Columns.Add("Column9", "Output")

    End Sub

    Private Sub LoadWOTable()
        WOGrid.Columns.Add("Column1", "S/No")
        For i As Integer = 1 To 29
            WOGrid.Rows(WOGrid.Rows.Add()).Cells("Column1").Value = i
        Next
        WOGrid.Columns.Add("Coloumn2", "WO Nos")
        WOGrid.Columns.Add("Coloumn3", "WO Model")
        WOGrid.Columns.Add("Coloumn4", "WO Qty")
        WOGrid.Columns.Add("Coloumn5", "PF")
        WOGrid.Columns.Add("Coloumn6", "Tag Nos")
        WOGrid.Columns.Add("Coloumn7", "Date Created")
        WOGrid.Columns.Add("Coloumn8", "Date Closed")
        WOGrid.Columns.Add("Coloumn9", "Status")
    End Sub
    Private Sub LoadPSNData()
        Dim PSNCount As Integer
        Dim selectedItemName As String
        If Text6.Text = "" Then
            File1.Refresh()
            PSNCount = File1.Items.Count
            'MSFlexGrid1.Rows = PSNCount + 1
            MSFlexGrid1.Rows.Clear()

            For i As Integer = 1 To PSNCount
                Debug.WriteLine(i)
                MSFlexGrid1.Rows.Add()
                File1.SelectedIndex = i - 1
                selectedItemName = Path.GetFileNameWithoutExtension(File1.SelectedItem.ToString())
                'MSFlexGrid1.Col = 1
                'MSFlexGrid1.Row = i
                'MSFlexGrid1.CellAlignment = 1
                'MSFlexGrid1.Text = Mid(File1.SelectedItem(i - 1), 1, InStr(File1.SelectedItem(i - 1), ".") - 1)
                '        MSFlexGrid1.Rows(i).Cells(1).Value = Mid(File1.SelectedItem(i - 1).ToString(), 1, InStr(File1.SelectedItem(i - 1).ToString(), ".") - 1)
                MSFlexGrid1.Rows(i - 1).Cells(1).Value = selectedItemName
                If LOADPSNFILE(MSFlexGrid1.Rows(i - 1).Cells(1).Value) Then
                    'MSFlexGrid1.Col = 2
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = PSNFileInfo.MainPCBA
                    MSFlexGrid1.Rows(i - 1).Cells(2).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(i - 1).Cells(2).Value = PSNFileInfo.MainPCBA

                    'MSFlexGrid1.Col = 3
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = PSNFileInfo.ElectroMagnet
                    MSFlexGrid1.Rows(i - 1).Cells(3).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(i - 1).Cells(3).Value = PSNFileInfo.ElectroMagnet

                    If PSNFileInfo.BodyAssyStatus = "PASS" Then
                        'MSFlexGrid1.Col = 4
                        'MSFlexGrid1.CellBackColor = vbGreen
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(4).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(i - 1).Cells(4).Value = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(4).Style.BackColor = Color.Green
                    Else
                        'MSFlexGrid1.Col = 4
                        If PSNFileInfo.BodyAssyStatus = "FAIL" Then
                            'MSFlexGrid1.CellBackColor = vbRed
                            'MSFlexGrid1.Text = "FAIL"
                            'MSFlexGrid1.CellAlignment = 3
                            MSFlexGrid1.Rows(i - 1).Cells(4).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(4).Value = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(4).Style.BackColor = Color.Red
                        Else
                            MSFlexGrid1.Rows(i - 1).Cells(4).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(4).Value = ""
                            MSFlexGrid1.Rows(i - 1).Cells(4).Style.BackColor = Color.White
                        End If
                    End If

                    If PSNFileInfo.ScrewStnStatus = "PASS" Then
                        'MSFlexGrid1.Col = 5
                        'MSFlexGrid1.CellBackColor = vbGreen
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(5).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(i - 1).Cells(5).Value = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(5).Style.BackColor = Color.Green
                    Else
                        'MSFlexGrid1.Col = 5
                        If PSNFileInfo.ScrewStnStatus = "FAIL" Then
                            'MSFlexGrid1.CellBackColor = vbRed
                            'MSFlexGrid1.CellAlignment = 3
                            'MSFlexGrid1.Text = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(5).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(5).Value = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(5).Style.BackColor = Color.Red
                        Else
                            MSFlexGrid1.Rows(i - 1).Cells(5).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(5).Value = ""
                            MSFlexGrid1.Rows(i - 1).Cells(5).Style.BackColor = Color.White
                        End If
                    End If
                    If PSNFileInfo.FTStatus = "PASS" Then
                        'MSFlexGrid1.Col = 6
                        'MSFlexGrid1.CellBackColor = vbGreen
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(6).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(i - 1).Cells(6).Value = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(6).Style.BackColor = Color.Green
                    Else
                        'MSFlexGrid1.Col = 6
                        If PSNFileInfo.FTStatus = "FAIL" Then
                            MSFlexGrid1.Rows(i - 1).Cells(6).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(6).Value = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(6).Style.BackColor = Color.Red
                        Else
                            MSFlexGrid1.Rows(i - 1).Cells(6).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(6).Value = ""
                            MSFlexGrid1.Rows(i - 1).Cells(6).Style.BackColor = Color.White
                        End If
                    End If

                    'MSFlexGrid1.Col = 7
                    'MSFlexGrid1.CellBackColor = vbBlue
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "-"
                    MSFlexGrid1.Rows(i - 1).Cells(7).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(i - 1).Cells(7).Value = "-"
                    MSFlexGrid1.Rows(i - 1).Cells(7).Style.BackColor = Color.Blue

                    If PSNFileInfo.Stn5Status = "PASS" Then
                        'MSFlexGrid1.Col = 8
                        'MSFlexGrid1.CellBackColor = vbGreen
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(8).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(i - 1).Cells(8).Value = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(8).Style.BackColor = Color.Green
                    Else
                        'MSFlexGrid1.Col = 8
                        If PSNFileInfo.Stn5Status = "FAIL" Then
                            MSFlexGrid1.Rows(i - 1).Cells(8).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(8).Value = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(8).Style.BackColor = Color.Red
                        Else
                            MSFlexGrid1.Rows(i - 1).Cells(8).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(8).Value = ""
                            MSFlexGrid1.Rows(i - 1).Cells(8).Style.BackColor = Color.White
                        End If
                    End If

                    If PSNFileInfo.VacuumStatus = "PASS" Then
                        MSFlexGrid1.Rows(i - 1).Cells(9).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(i - 1).Cells(9).Value = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(9).Style.BackColor = Color.Green
                    Else
                        'MSFlexGrid1.Col = 9
                        If PSNFileInfo.VacuumStatus = "FAIL" Then
                            MSFlexGrid1.Rows(i - 1).Cells(9).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(9).Value = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(9).Style.BackColor = Color.Red
                        Else
                            'MSFlexGrid1.CellBackColor = vbWhite
                            'MSFlexGrid1.Text = ""
                            MSFlexGrid1.Rows(i - 1).Cells(9).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(9).Value = ""
                            MSFlexGrid1.Rows(i - 1).Cells(9).Style.BackColor = Color.White
                        End If
                    End If

                    If PSNFileInfo.ConnTestStatus = "PASS" Then
                        'MSFlexGrid1.Col = 10
                        'MSFlexGrid1.CellBackColor = vbGreen
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(10).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(i - 1).Cells(10).Value = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(10).Style.BackColor = Color.Green
                    Else
                        'MSFlexGrid1.Col = 10
                        If PSNFileInfo.ConnTestStatus = "FAIL" Then
                            'MSFlexGrid1.CellBackColor = vbRed
                            'MSFlexGrid1.CellAlignment = 3
                            'MSFlexGrid1.Text = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(10).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(10).Value = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(10).Style.BackColor = Color.Red
                        Else
                            'MSFlexGrid1.CellBackColor = vbWhite
                            'MSFlexGrid1.Text = ""
                            MSFlexGrid1.Rows(i - 1).Cells(10).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(10).Value = ""
                            MSFlexGrid1.Rows(i - 1).Cells(10).Style.BackColor = Color.White
                        End If
                    End If

                    If PSNFileInfo.Vacuum2Status = "PASS" Then
                        'MSFlexGrid1.Col = 11
                        'MSFlexGrid1.CellBackColor = vbGreen
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(11).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(i - 1).Cells(11).Value = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(11).Style.BackColor = Color.Green
                    Else
                        'MSFlexGrid1.Col = 11
                        If PSNFileInfo.Vacuum2Status = "FAIL" Then
                            'MSFlexGrid1.CellBackColor = vbRed
                            'MSFlexGrid1.CellAlignment = 3
                            'MSFlexGrid1.Text = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(11).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(11).Value = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(11).Style.BackColor = Color.Red
                        Else
                            'MSFlexGrid1.CellBackColor = vbWhite
                            'MSFlexGrid1.Text = ""
                            MSFlexGrid1.Rows(i - 1).Cells(11).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(11).Value = ""
                            MSFlexGrid1.Rows(i - 1).Cells(11).Style.BackColor = Color.White
                        End If
                    End If

                    If PSNFileInfo.PackagingStatus = "PASS" Then
                        'MSFlexGrid1.Col = 12
                        'MSFlexGrid1.CellBackColor = vbGreen
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(12).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(i - 1).Cells(12).Value = "PASS"
                        MSFlexGrid1.Rows(i - 1).Cells(12).Style.BackColor = Color.Green
                    Else
                        'MSFlexGrid1.Col = 12
                        If PSNFileInfo.PackagingStatus = "FAIL" Then
                            'MSFlexGrid1.CellBackColor = vbRed
                            'MSFlexGrid1.CellAlignment = 3
                            'MSFlexGrid1.Text = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(12).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(12).Value = "FAIL"
                            MSFlexGrid1.Rows(i - 1).Cells(12).Style.BackColor = Color.Red
                        Else
                            'MSFlexGrid1.CellBackColor = vbWhite
                            'MSFlexGrid1.Text = ""
                            MSFlexGrid1.Rows(i - 1).Cells(12).Style.Alignment = DataGridViewContentAlignment.TopCenter
                            MSFlexGrid1.Rows(i - 1).Cells(12).Value = ""
                            MSFlexGrid1.Rows(i - 1).Cells(12).Style.BackColor = Color.White
                        End If
                    End If
                    'MSFlexGrid1.Col = 13
                    'MSFlexGrid1.Text = PSNFileInfo.DateCompleted
                    MSFlexGrid1.Rows(i - 1).Cells(13).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(i - 1).Cells(13).Value = PSNFileInfo.DateCompleted
                End If
            Next
        Else
            'MSFlexGrid1.Rows = 2
            MSFlexGrid1.Rows.Clear()
            MSFlexGrid1.Rows.Add()
            If LOADPSNFILE(Text6.Text) Then

                'MSFlexGrid1.Col = 1
                'MSFlexGrid1.Row = 1
                'MSFlexGrid1.CellAlignment = 1
                'MSFlexGrid1.Text = Text6
                MSFlexGrid1.Rows(0).Cells(1).Style.Alignment = DataGridViewContentAlignment.TopLeft
                MSFlexGrid1.Rows(0).Cells(1).Value = Text6.Text

                'MSFlexGrid1.Col = 2
                'MSFlexGrid1.CellAlignment = 3
                'MSFlexGrid1.Text = PSNFileInfo.MainPCBA
                MSFlexGrid1.Rows(0).Cells(2).Style.Alignment = DataGridViewContentAlignment.TopLeft
                MSFlexGrid1.Rows(0).Cells(2).Value = PSNFileInfo.MainPCBA

                'MSFlexGrid1.Col = 3
                'MSFlexGrid1.CellAlignment = 3
                'MSFlexGrid1.Text = PSNFileInfo.ElectroMagnet

                MSFlexGrid1.Rows(0).Cells(3).Style.Alignment = DataGridViewContentAlignment.TopLeft
                MSFlexGrid1.Rows(0).Cells(3).Value = PSNFileInfo.ElectroMagnet

                If PSNFileInfo.BodyAssyStatus = "PASS" Then
                    'MSFlexGrid1.Col = 4
                    'MSFlexGrid1.CellBackColor = vbGreen
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "PASS"
                    MSFlexGrid1.Rows(0).Cells(4).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(0).Cells(4).Value = "PASS"
                    MSFlexGrid1.Rows(0).Cells(4).Style.BackColor = Color.Green
                Else
                    'MSFlexGrid1.Col = 4
                    If PSNFileInfo.BodyAssyStatus = "FAIL" Then
                        'MSFlexGrid1.CellBackColor = vbRed
                        'MSFlexGrid1.Text = "FAIL"
                        'MSFlexGrid1.CellAlignment = 3
                        MSFlexGrid1.Rows(0).Cells(4).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(4).Value = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(4).Style.BackColor = Color.Red
                    Else
                        'MSFlexGrid1.CellBackColor = vbWhite
                        'MSFlexGrid1.Text = ""
                        MSFlexGrid1.Rows(0).Cells(4).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(4).Value = ""
                        MSFlexGrid1.Rows(0).Cells(4).Style.BackColor = Color.White
                    End If
                End If

                If PSNFileInfo.ScrewStnStatus = "PASS" Then
                    'MSFlexGrid1.Col = 5
                    'MSFlexGrid1.CellBackColor = vbGreen
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "PASS"
                    MSFlexGrid1.Rows(0).Cells(5).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(0).Cells(5).Value = "PASS"
                    MSFlexGrid1.Rows(0).Cells(5).Style.BackColor = Color.Green
                Else
                    'MSFlexGrid1.Col = 5
                    If PSNFileInfo.ScrewStnStatus = "FAIL" Then
                        'MSFlexGrid1.CellBackColor = vbRed
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(5).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(5).Value = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(5).Style.BackColor = Color.Red
                    Else
                        'MSFlexGrid1.CellBackColor = vbWhite
                        'MSFlexGrid1.Text = ""
                        MSFlexGrid1.Rows(0).Cells(5).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(5).Value = ""
                        MSFlexGrid1.Rows(0).Cells(5).Style.BackColor = Color.White
                    End If
                End If
                If PSNFileInfo.FTStatus = "PASS" Then
                    'MSFlexGrid1.Col = 6
                    'MSFlexGrid1.CellBackColor = vbGreen
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "PASS"
                    MSFlexGrid1.Rows(0).Cells(6).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(0).Cells(6).Value = "PASS"
                    MSFlexGrid1.Rows(0).Cells(6).Style.BackColor = Color.Green
                Else
                    'MSFlexGrid1.Col = 6
                    If PSNFileInfo.FTStatus = "FAIL" Then
                        'MSFlexGrid1.CellBackColor = vbRed
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(6).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(6).Value = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(6).Style.BackColor = Color.Red
                    Else
                        'MSFlexGrid1.CellBackColor = vbWhite
                        'MSFlexGrid1.Text = ""
                        MSFlexGrid1.Rows(0).Cells(6).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(6).Value = ""
                        MSFlexGrid1.Rows(0).Cells(6).Style.BackColor = Color.White
                    End If
                End If

                'MSFlexGrid1.Col = 7
                'MSFlexGrid1.CellBackColor = vbBlue
                'MSFlexGrid1.CellAlignment = 3
                'MSFlexGrid1.Text = "-"
                MSFlexGrid1.Rows(0).Cells(7).Style.Alignment = DataGridViewContentAlignment.TopCenter
                MSFlexGrid1.Rows(0).Cells(7).Value = "-"
                MSFlexGrid1.Rows(0).Cells(7).Style.BackColor = Color.Blue

                If PSNFileInfo.Stn5Status = "PASS" Then
                    'MSFlexGrid1.Col = 8
                    'MSFlexGrid1.CellBackColor = vbGreen
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "PASS"
                    MSFlexGrid1.Rows(0).Cells(8).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(0).Cells(8).Value = "PASS"
                    MSFlexGrid1.Rows(0).Cells(8).Style.BackColor = Color.Green
                Else
                    'MSFlexGrid1.Col = 8
                    If PSNFileInfo.Stn5Status = "FAIL" Then
                        'MSFlexGrid1.CellBackColor = vbRed
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(8).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(8).Value = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(8).Style.BackColor = Color.Red
                    Else
                        'MSFlexGrid1.CellBackColor = vbWhite
                        'MSFlexGrid1.Text = ""
                        MSFlexGrid1.Rows(0).Cells(8).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(8).Value = ""
                        MSFlexGrid1.Rows(0).Cells(8).Style.BackColor = Color.White
                    End If
                End If

                If PSNFileInfo.VacuumStatus = "PASS" Then
                    'MSFlexGrid1.Col = 9
                    'MSFlexGrid1.CellBackColor = vbGreen
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "PASS"
                    MSFlexGrid1.Rows(0).Cells(9).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(0).Cells(9).Value = "PASS"
                    MSFlexGrid1.Rows(0).Cells(9).Style.BackColor = Color.Green
                Else
                    'MSFlexGrid1.Col = 9
                    If PSNFileInfo.PackagingStatus = "FAIL" Then
                        'MSFlexGrid1.CellBackColor = vbRed
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(9).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(9).Value = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(9).Style.BackColor = Color.Red
                    Else
                        MSFlexGrid1.Rows(0).Cells(9).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(9).Value = ""
                        MSFlexGrid1.Rows(0).Cells(9).Style.BackColor = Color.White
                    End If
                End If

                If PSNFileInfo.ConnTestStatus = "PASS" Then
                    'MSFlexGrid1.Col = 10
                    'MSFlexGrid1.CellBackColor = vbGreen
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "PASS"
                    MSFlexGrid1.Rows(0).Cells(10).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(0).Cells(10).Value = "PASS"
                    MSFlexGrid1.Rows(0).Cells(10).Style.BackColor = Color.Green
                Else
                    'MSFlexGrid1.Col = 10
                    If PSNFileInfo.ConnTestStatus = "FAIL" Then
                        'MSFlexGrid1.CellBackColor = vbRed
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(10).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(10).Value = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(10).Style.BackColor = Color.Red
                    Else
                        MSFlexGrid1.Rows(0).Cells(10).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(10).Value = ""
                        MSFlexGrid1.Rows(0).Cells(10).Style.BackColor = Color.White
                    End If
                End If

                If PSNFileInfo.Vacuum2Status = "PASS" Then
                    'MSFlexGrid1.Col = 11
                    'MSFlexGrid1.CellBackColor = vbGreen
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "PASS"
                    MSFlexGrid1.Rows(0).Cells(11).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(0).Cells(11).Value = "PASS"
                    MSFlexGrid1.Rows(0).Cells(11).Style.BackColor = Color.Green
                Else
                    'MSFlexGrid1.Col = 11
                    If PSNFileInfo.Vacuum2Status = "FAIL" Then
                        'MSFlexGrid1.CellBackColor = vbRed
                        'MSFlexGrid1.CellAlignment = 3
                        'MSFlexGrid1.Text = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(11).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(11).Value = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(11).Style.BackColor = Color.Red
                    Else
                        MSFlexGrid1.Rows(0).Cells(11).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(11).Value = ""
                        MSFlexGrid1.Rows(0).Cells(11).Style.BackColor = Color.White
                    End If
                End If

                If PSNFileInfo.PackagingStatus = "PASS" Then
                    'MSFlexGrid1.Col = 12
                    'MSFlexGrid1.CellBackColor = vbGreen
                    'MSFlexGrid1.CellAlignment = 3
                    'MSFlexGrid1.Text = "PASS"
                    MSFlexGrid1.Rows(0).Cells(12).Style.Alignment = DataGridViewContentAlignment.TopCenter
                    MSFlexGrid1.Rows(0).Cells(12).Value = "PASS"
                    MSFlexGrid1.Rows(0).Cells(12).Style.BackColor = Color.Green
                Else
                    'MSFlexGrid1.Col = 12
                    If PSNFileInfo.PackagingStatus = "FAIL" Then
                        MSFlexGrid1.Rows(0).Cells(12).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(12).Value = "FAIL"
                        MSFlexGrid1.Rows(0).Cells(12).Style.BackColor = Color.Red
                    Else
                        MSFlexGrid1.Rows(0).Cells(12).Style.Alignment = DataGridViewContentAlignment.TopCenter
                        MSFlexGrid1.Rows(0).Cells(12).Value = ""
                        MSFlexGrid1.Rows(0).Cells(12).Style.BackColor = Color.White
                    End If
                End If
                'MSFlexGrid1.Col = 13
                'MSFlexGrid1.Text = PSNFileInfo.DateCompleted
                MSFlexGrid1.Rows(0).Cells(13).Style.Alignment = DataGridViewContentAlignment.TopCenter
                MSFlexGrid1.Rows(0).Cells(13).Value = PSNFileInfo.DateCompleted
            End If
        End If
    End Sub

    Public Function LOADPSNFILE(ProductPSN As String) As Boolean

        Dim ItemStr As String
        Dim SectionHeading As String
        Dim pos1, pos2, pos3 As Integer
        ClearPSNVar()
        FNum = FreeFile()

        If String.IsNullOrEmpty(INIPSNFOLDERPATH & ProductPSN & ".Txt") Then
            'SetDefaultINIValues
            'WriteINI
            Exit Function
        End If


        FileOpen(FNum, INIPSNFOLDERPATH & ProductPSN & ".Txt", OpenMode.Input)
        Do While Not EOF(FNum)
            Try
                Linestr = LineInput(FNum)
                'Debug.WriteLine(Linestr)
                'Check for Section heading
                If InStr(Linestr, "[") > 0 And InStr(Linestr, "]") > 0 Then
                    pos1 = InStr(Linestr, "[")
                    pos2 = InStr(Linestr, "]")
                    pos3 = InStr(Linestr, ":")

                    SectionHeading = Mid(Linestr, pos1 + 1, pos2 - pos1 - 1)

                    Select Case UCase(SectionHeading)
                        Case "MODEL"
                            PSNFileInfo.ModelName = Trim(Mid(Linestr, pos3 + 1))

                        Case "DATE CREATED"
                            PSNFileInfo.DateCreated = Trim(Mid(Linestr, pos3 + 1))

                        Case "DATE COMPLETED"
                            PSNFileInfo.DateCompleted = Trim(Mid(Linestr, pos3 + 1))

                        Case "OPERATOR ID"
                            PSNFileInfo.OperatorID = Trim(Mid(Linestr, pos3 + 1))

                        Case "WORK ORDER NO"
                            PSNFileInfo.WONos = Trim(Mid(Linestr, pos3 + 1))

                        Case "MAIN PCBA S/N"
                            PSNFileInfo.MainPCBA = Trim(Mid(Linestr, pos3 + 1))

                        Case "SECONDARY PCBA S/N"
                            PSNFileInfo.SecondaryPCBA = Trim(Mid(Linestr, pos3 + 1))

                        Case "ELECTROMAGNET S/N"
                            PSNFileInfo.ElectroMagnet = Trim(Mid(Linestr, pos3 + 1))

                        Case "PSN"
                            PSNFileInfo.PSN = Trim(Mid(Linestr, pos3 + 1))

                        Case "BODY ASSY STATION CHECK IN DATE"
                            PSNFileInfo.BodyAssyCheckIn = Trim(Mid(Linestr, pos3 + 1))

                        Case "BODY ASSY STATION CHECK OUT DATE"
                            PSNFileInfo.BodyAssyCheckOut = Trim(Mid(Linestr, pos3 + 1))

                        Case "BODY ASSY STATION STATUS"
                            PSNFileInfo.BodyAssyStatus = Trim(Mid(Linestr, pos3 + 1))

                        Case "SCREWING STATION CHECK IN DATE"
                            PSNFileInfo.ScrewStnCheckIn = Trim(Mid(Linestr, pos3 + 1))

                        Case "SCREWING STATION CHECK OUT DATE"
                            PSNFileInfo.ScrewStnCheckOut = Trim(Mid(Linestr, pos3 + 1))

                        Case "SCREWING STATION STATUS"
                            PSNFileInfo.ScrewStnStatus = Trim(Mid(Linestr, pos3 + 1))

                        Case "FINAL TEST CHECK IN DATE"
                            PSNFileInfo.FTCheckIn = Trim(Mid(Linestr, pos3 + 1))

                        Case "FINAL TEST CHECK OUT DATE"
                            PSNFileInfo.FTCheckOut = Trim(Mid(Linestr, pos3 + 1))

                        Case "FINAL TEST STATUS"
                            PSNFileInfo.FTStatus = Trim(Mid(Linestr, pos3 + 1))

                        Case "STATION 5 CHECK IN DATE"
                            PSNFileInfo.Stn5CheckIn = Trim(Mid(Linestr, pos3 + 1))

                        Case "STATION 5 CHECK OUT DATE"
                            PSNFileInfo.Stn5CheckOut = Trim(Mid(Linestr, pos3 + 1))

                        Case "STATION 5 STATUS"
                            PSNFileInfo.Stn5Status = Trim(Mid(Linestr, pos3 + 1))

                        Case "VACUUM CHECK IN DATE"
                            PSNFileInfo.VacuumCheckIn = Trim(Mid(Linestr, pos3 + 1))

                        Case "VACUUM CHECK OUT DATE"
                            PSNFileInfo.VacummCheckOut = Trim(Mid(Linestr, pos3 + 1))

                        Case "VACUUM STATUS"
                            PSNFileInfo.VacuumStatus = Trim(Mid(Linestr, pos3 + 1))

                        Case "CONNECTOR TEST CHECK IN DATE"
                            PSNFileInfo.ConnTestCheckIn = Trim(Mid(Linestr, pos3 + 1))

                        Case "CONNECTOR TEST CHECK OUT DATE"
                            PSNFileInfo.ConnTestCheckOut = Trim(Mid(Linestr, pos3 + 1))

                        Case "CONNECTOR TEST STATUS"
                            PSNFileInfo.ConnTestStatus = Trim(Mid(Linestr, pos3 + 1))

                        Case "VACUUM #2 CHECK IN DATE"
                            PSNFileInfo.Vacuum2CheckIn = Trim(Mid(Linestr, pos3 + 1))

                        Case "VACUUM #2 CHECK OUT DATE"
                            PSNFileInfo.Vacumm2CheckOut = Trim(Mid(Linestr, pos3 + 1))

                        Case "VACUUM #2 STATUS"
                            PSNFileInfo.Vacuum2Status = Trim(Mid(Linestr, pos3 + 1))

                        Case "PACKAGING CHECK IN DATE"
                            PSNFileInfo.PackagingCheckIn = Trim(Mid(Linestr, pos3 + 1))

                        Case "PACKAGING CHECK OUT DATE"
                            PSNFileInfo.PackagingCheckOut = Trim(Mid(Linestr, pos3 + 1))

                        Case "PACKAGING STATUS"
                            PSNFileInfo.PackagingStatus = Trim(Mid(Linestr, pos3 + 1))

                        Case "DEBUG STATION #10 STATUS"
                            PSNFileInfo.DebugStatus = Trim(Mid(Linestr, pos3 + 1))

                        Case "DEBUG COMMENTS"
                            PSNFileInfo.DebugComment = Trim(Mid(Linestr, pos3 + 1))

                        Case "DEBUG TECHNICIANS ID"
                            PSNFileInfo.DebugTechnican = Trim(Mid(Linestr, pos3 + 1))

                        Case "DEBUG DATE REPAIRED"
                            PSNFileInfo.RepairDate = Trim(Mid(Linestr, pos3 + 1))

                    End Select
                End If
            Catch ex As Exception
                Return False
                Exit Do
            End Try

        Loop
        FileClose(FNum)
        Return True
        Exit Function
    End Function


    Public Sub CheckConnections(TimerIndex As Integer)
        Dim Reply As ICMP_ECHO_REPLY
        If BusyFlag = False Then
            Select Case TimerIndex
                Case 1 : StationConnectionStatus(1) = ping(INISTNPCIP(1), Reply) 'SA1
                Case 2 : StationConnectionStatus(2) = ping(INISTNPCIP(2), Reply) 'SA2
                Case 3 : StationConnectionStatus(3) = ping(INISTNPCIP(3), Reply) 'Main PCBA
                Case 4 : StationConnectionStatus(4) = ping(INISTNPCIP(4), Reply) 'Electromagnet & Locking PCBA
                Case 5 : StationConnectionStatus(5) = ping(INISTNPCIP(5), Reply) 'PCBA to Body
                Case 6 : StationConnectionStatus(6) = ping(INISTNPCIP(6), Reply) 'Screwing head to Body
                Case 7 : StationConnectionStatus(7) = ping(INISTNPCIP(7), Reply) 'Tester
                Case 8 : StationConnectionStatus(8) = ping(INISTNPCIP(8), Reply) 'Cover Assy
                Case 9 : StationConnectionStatus(9) = ping(INISTNPCIP(9), Reply) 'Packaging
                Case 10 : StationConnectionStatus(10) = ping(INISTNPCIP(10), Reply) 'Continuity Tester
            End Select
        End If
    End Sub

    Private Function ping(sAddress As String, Reply As ICMP_ECHO_REPLY) As Long

        Dim hIcmp As Long
        Dim lAddress As Long
        Dim lTimeOut As Long
        Dim StringToSend As String

        'Short string of data to send
        StringToSend = "hello"

        'ICMP (ping) timeout
        lTimeOut = 1 'ms

        'Convert string address to a long representation.
        lAddress = inet_addr(sAddress)

        If (lAddress <> -1) And (lAddress <> 0) Then

            'Create the handle for ICMP requests.
            hIcmp = IcmpCreateFile()

            If hIcmp Then
                'Ping the destination IP address.
                Call IcmpSendEcho(hIcmp, lAddress, StringToSend, Len(StringToSend), 0, Reply, Len(Reply), lTimeOut)

                'Reply status
                ping = Reply.status

                'Close the Icmp handle.
                IcmpCloseHandle(hIcmp)
            Else
                Debug.Print("failure opening icmp handle.")
                ping = -1
            End If
        Else
            ping = -1
        End If

    End Function
    Private Sub LoginToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Cmd_Login.Click
        If Cmd_Login.Text = "Login" Then
            Me.Hide()
            FrmLogin.ShowDialog()
            If GVL.Login Then
                Cmd_RFID.Enabled = True
                Cmd_Tag.Enabled = True
                Cmd_Parameters.Enabled = True
                Cmd_Parts.Enabled = True
                Cmd_Login.Text = "LogOff"
            Else
                Cmd_RFID.Enabled = False
                Cmd_Tag.Enabled = False
                Cmd_Parameters.Enabled = False
                Cmd_Parts.Enabled = False
            End If
        Else
            Cmd_RFID.Enabled = False
            Cmd_Tag.Enabled = False
            Cmd_Parameters.Enabled = False
            Cmd_Parts.Enabled = False
            Cmd_Login.Text = "Login"
        End If
    End Sub

    Private Sub Cmd_TestSpec_Click(sender As Object, e As EventArgs) Handles Cmd_TestSpec.Click
        Timer1.Enabled = False
        Me.Hide()
        FrmDatabase.ShowDialog()
    End Sub

    Private Sub BarcodeScan_Comm_DataReceived(sender As Object, e As Ports.SerialDataReceivedEventArgs) Handles BarcodeScan_Comm.DataReceived
        WOBuffer = BarcodeScan_Comm.ReadExisting()

        If InStr(1, WOBuffer, vbCrLf) <> 0 Then
            Me.Invoke(Sub()
                          WOBuffer = Mid(WOBuffer, 1, WOBuffer.IndexOf(vbCr) - 1) 'Trim off VBCRLF
                          If SelectedMode = 1 Then 'WO Entry
                              If CSAction = 0 Then
                                  WOBuffer = Mid(WOBuffer, 2, 1)
                                  CSInfo.CSWOPFMODE = WOBuffer
                                  _Text1_3.Text = CSInfo.CSWOPFMODE
                                  Label2.Text = "Please scan the WO Nos Barcode..."
                                  WOBuffer = ""
                                  If CSInfo.CSWOPFMODE = "C" Then
                                      _Label1_13.Visible = True
                                      _Text1_4.Visible = True
                                  ElseIf CSInfo.CSWOPFMODE = "S" Then
                                      _Label1_13.Visible = False
                                      _Text1_4.Visible = False
                                  End If
                                  CSAction = 1
                                  Exit Sub
                              ElseIf CSAction = 1 Then
                                  CSInfo.CSWONOS = WOBuffer
                                  _Text1_0.Text = CSInfo.CSWONOS
                                  If WONosCheck(WOBuffer) Then
                                      Label2.Text = "WO already exist. Cannot be enter again" & vbCrLf & "Please start over again"
                                      WOBuffer = ""
                                      CSAction = 0
                                      _Text1_0.Text = ""
                                      _Text1_1.Text = ""
                                      _Text1_2.Text = ""
                                      _Text1_3.Text = ""
                                      _Text1_4.Text = ""
                                      SelectedMode = 1
                                      CSAction = 0
                                      Exit Sub
                                  End If
                                  Label2.Text = "Please scan the Reference Barcode..."
                                  WOBuffer = ""
                                  CSAction = 2
                                  Exit Sub
                              ElseIf CSAction = 2 Then
                                  WOBuffer = Mid(WOBuffer, 2)
                                  If Not WORefCheck(WOBuffer) Then
                                      Label2.Text = "Invalid Reference. Reference Name not exist in database" & vbCrLf & "Please start over again"
                                      WOBuffer = ""
                                      CSAction = 0
                                      _Text1_0.Text = ""
                                      _Text1_1.Text = ""
                                      _Text1_2.Text = ""
                                      _Text1_3.Text = ""
                                      _Text1_4.Text = ""
                                      SelectedMode = 1
                                      CSAction = 0
                                      Exit Sub
                                  End If
                                  CSInfo.CSWOMODEL = WOBuffer
                                  _Text1_1.Text = CSInfo.CSWOMODEL
                                  Label2.Text = "Please scan the Quantity Barcode..."
                                  WOBuffer = ""
                                  CSAction = 3
                                  Exit Sub
                              ElseIf CSAction = 3 Then
                                  If Not IsNumeric(WOBuffer) Then
                                      Label2.Text = "Invalid Barcode Quantity" & vbCrLf & "Please start over again"
                                      CSAction = 0
                                      _Text1_0.Text = ""
                                      _Text1_1.Text = ""
                                      _Text1_2.Text = ""
                                      _Text1_3.Text = ""
                                      _Text1_4.Text = ""
                                      WOBuffer = ""
                                      SelectedMode = 1
                                      CSAction = 0
                                      Exit Sub
                                  End If
                                  CSInfo.CSWOQTY = WOBuffer
                                  _Text1_2.Text = CSInfo.CSWOQTY
                                  WOBuffer = ""
                                  If CSInfo.CSWOPFMODE = "S" Then
                                      GoTo PrgTag
                                  Else
                                      CSAction = 4
                                  End If
                                  Exit Sub
                              ElseIf CSAction = 4 Then
                                  CSInfo.CSWOLC = WOBuffer
                                  _Text1_4.Text = WOBuffer
                                  WOBuffer = ""
                                  GoTo PrgTag
                              End If

PrgTag:
                              Label2.Text = "Please wait while Verifying the Tag..."
                              Dim Rdtagcount As Long
                              Dim Rdtagnos As String
                              Dim Rdtagref As String
                              Dim RdtagWOnos As String
                              Dim Rdtagqty As String
                              Dim Rdtagop As String

                              Rdtagcount = CLng(RD_MULTI_RFID("004C", 3)) 'Tag Life Cycle
                              Rdtagnos = RD_MULTI_RFID("0040", 3)
                              Rdtagref = RD_MULTI_RFID("0014", 10)
                              RdtagWOnos = RD_MULTI_RFID("0000", 10)
                              Rdtagqty = RD_MULTI_RFID("0028", 10)
                              Rdtagop = RD_MULTI_RFID("0046", 3) 'WO output counter at packaging

                              Label5.Text = Rdtagcount
                              Label2.Text = "Please wait while programming the Tag..."
                              WOBuffer = ""
                              'Program RFID tag
                              'BarcodeScan_Comm.PortOpen = False

                              'Label2.Caption = "Checking RFID Tag Write Cycle..."
                              'If Rdtagcount < CDbl(CSInfo.CSWOQTY) Then
                              '    Label2.Caption = "Write Cycle unable to support WO quantity. Use another Tag and try again"
                              '    CSAction = 0
                              '    Text1(0).Text = ""
                              '    Text1(1).Text = ""
                              '    Text1(2).Text = ""
                              '    Text1(3).Text = ""
                              '    Text4(4).Text = ""
                              '    WOBuffer = ""
                              '    Exit Sub
                              'End If
                              Label2.Text = "Writing Work Order Number..."
                              If Not Wr_Tag(CSInfo.CSWONOS, "0000") Then
                                  Label2.Text = "Unable to write to address 0000H. Verify and try again from step 1."
                                  CSAction = 0
                                  _Text1_0.Text = ""
                                  _Text1_1.Text = ""
                                  _Text1_2.Text = ""
                                  _Text1_3.Text = ""
                                  _Text4_4.Text = ""
                                  WOBuffer = ""
                                  SelectedMode = 1
                                  CSAction = 0
                                  Exit Sub
                              End If
                              Label2.Text = "Writing Work Order Reference..."
                              If Not Wr_Tag(CSInfo.CSWOMODEL, "0014") Then
                                  Label2.Text = "Unable to write to address 0014H. Verify and try again from Step 1."
                                  CSAction = 0
                                  _Text1_0.Text = ""
                                  _Text1_1.Text = ""
                                  _Text1_2.Text = ""
                                  _Text1_3.Text = ""
                                  _Text4_4.Text = ""
                                  WOBuffer = ""
                                  SelectedMode = 1
                                  CSAction = 0
                                  Exit Sub
                              End If
                              Label2.Text = "Writing Work Order Quantity..."
                              If Not Wr_Tag(CSInfo.CSWOQTY, "0028") Then
                                  Label2.Text = "Unable to write to address 0028H. Verify and try again from Step 1."
                                  CSAction = 0
                                  _Text1_0.Text = ""
                                  _Text1_1.Text = ""
                                  _Text1_2.Text = ""
                                  WOBuffer = ""
                                  SelectedMode = 1
                                  CSAction = 0
                                  Exit Sub
                              End If
                              Label2.Text = "Writing WO PF Mode..."
                              If Not Wr_Tag(CSInfo.CSWOPFMODE, "0052") Then
                                  Label2.Text = "Unable to write to address 0052H. Verify and try again from Step 1."
                                  CSAction = 0
                                  _Text1_0.Text = ""
                                  _Text1_1.Text = ""
                                  _Text1_2.Text = ""
                                  _Text1_3.Text = ""
                                  _Text4_4.Text = ""
                                  WOBuffer = ""
                                  SelectedMode = 1
                                  CSAction = 0
                                  Exit Sub
                              End If
                              Label2.Text = "Reseting packaging counter..."
                              If Not Wr_Tag("000000", "0046") Then
                                  Label2.Text = "Unable to write to address 0046H. Verify and try again from Step 1."
                                  CSAction = 0
                                  _Text1_0.Text = ""
                                  _Text1_1.Text = ""
                                  _Text1_2.Text = ""
                                  _Text1_3.Text = ""
                                  _Text4_4.Text = ""
                                  WOBuffer = ""
                                  SelectedMode = 1
                                  CSAction = 0
                                  Exit Sub
                              End If
                              Label2.Text = "Incrementing Tag cycle..."
                              If Not Wr_Tag(CDbl(Rdtagcount) + 1, "004C") Then
                                  Label2.Text = "Unable to write to address 0046H. Verify and try again from Step 1."
                                  CSAction = 0
                                  _Text1_0.Text = ""
                                  _Text1_1.Text = ""
                                  _Text1_2.Text = ""
                                  _Text1_3.Text = ""
                                  _Text4_4.Text = ""
                                  WOBuffer = ""
                                  SelectedMode = 1
                                  CSAction = 0
                                  Exit Sub
                              End If
                              Label2.Text = "RFID Tag programming completed."
                              Label5.Text = CDbl(Rdtagcount) + 1


                              'InString = InputBox("Enter new model Name", "Adding New Reference")
                              Dim query As String
                              Dim timesNow As String = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
                              If CSInfo.CSWOPFMODE = "PFC" Then
                                  query = "INSERT INTO [CSUNIT] ([UNITNOS],[WONOS],[WOMODELNAME],[WOQTY],[PFMODE],[LOGISTICC],[STATUS],[DATECREATED]) 
                            VALUES (ISNULL('" & Rdtagnos & "','" & CSInfo.CSWONOS & "','" & CSInfo.CSWOMODEL & "','" & CSInfo.CSWOQTY & "',
                            '" & CSInfo.CSWOPFMODE & "','" & CSInfo.CSWOLC & "','" & "OPEN" & "','" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "')"
                              Else
                                  query = "INSERT INTO [CSUNIT] ([UNITNOS],[WONOS],[WOMODELNAME],[WOQTY],[PFMODE],[LOGISTICC],[STATUS],[DATECREATED]) 
                            VALUES (ISNULL('" & Rdtagnos & "','" & CSInfo.CSWONOS & "','" & CSInfo.CSWOMODEL & "','" & CSInfo.CSWOQTY & "',
                             '" & CSInfo.CSWOPFMODE & "','','" & "OPEN" & "','" & timesNow & "')"
                              End If

                              If ConnectionDatabase.insertData(query) Then
                                  MsgBox("Success add database!")
                              Else
                                  MsgBox("Failed add database!")
                              End If
                              'Creating PSN Folder for this WO
                              'MkDir INIPSNFOLDERPATH & CSInfo.CSWONOS
                              CSAction = 0
                              WOBuffer = ""
                              'RFID_Comm.PortOpen = False
                              Exit Sub
                          ElseIf SelectedMode = 2 Then
                              Text6.Text = WOBuffer
                              If File.Exists(INIPSNFOLDERPATH & Text6.Text & ".Txt") Then
                                  MSFlexGrid1.Rows.Clear()
                                  MSFlexGrid1.Columns.Clear()
                                  LoadPSNTable()
                                  Text6.Text = WOBuffer
                                  WOBuffer = ""

                              End If
                              WOBuffer = ""
                          ElseIf SelectedMode = 3 Then
                              Text7.Text = WOBuffer
                              WOBuffer = ""
                          Else
                              WOBuffer = ""
                          End If
                      End Sub)

        End If
    End Sub

    Private Function WONosCheck(strName As String) As Boolean
        Dim query As String = "SELECT * FROM [CSUNIT] WHERE WONOS = '" & strName & "'"
        Dim dt As DataTable = ConnectionDatabase.readData(query).Tables(0)
        If dt.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function WORefCheck(strName As String) As Boolean
        Dim query As String = "SELECT * FROM [Parameter] WHERE ModelName = '" & strName & "'"
        Dim dt As DataTable = ConnectionDatabase.readData(query).Tables(0)
        If dt.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        '===== additional Check box for refresh grid, prevent program to busy checking file and soon=====
        '===== 20141027 by Ari Indra A =====
        Timer1.Enabled = False
        If PingFGStation.Checked Then
            Shape1.BackColor = Color.Red
            CheckConnections(TimerCount)
        End If

        If (TimerCount = 11) Then
            'LoadConnectData (TimerCount)
            'LoadCSData
            'UpdateCSdata
            'LoadFridgeData
            TimerCount = 3
        Else
            LoadConnectData(TimerCount)
            If RefreshFGWO.Checked Then
                Shape2.BackColor = Color.Red
                WOGrid.DataSource = Nothing
                'LoadWOTable()

                WOGrid.Rows.Clear()
                WOGrid.Columns.Clear()
                LoadWOData()
            End If
            If RefreshFGPSN.Checked Then
                shape3.BackColor = Color.Red
                LoadPSNData()
                RefreshFGPSN.Checked = False
            End If
            TimerCount = TimerCount + 1
        End If
        Shape1.BackColor = Color.Green
        Shape2.BackColor = Color.Green
        shape3.BackColor = Color.Green
        'Debug.Write("TICK=")
        'Debug.WriteLine(TimerCount)
        Timer1.Enabled = True
    End Sub

    'All stations will need to log an individual file known as Status#n.Txt
    'Inside this file, station will update the data after the process.
    'Data include WO Nos, WO Model, WO QTY, Actual O/P
    Public Sub LoadConnectData(TimerIndex As Integer)
        Dim i As Object
        Dim pos1, pos2, pos3, pos4 As Integer
        Dim FileNum As Integer
        Dim statuscode
        Dim yap
        Dim rowsIndex As Integer = 0
        On Error Resume Next
        'For i = 1 To 9
        'yap = Dir(INISTNPATH(i))
        'Debug.Write("LOADCONNECTDATA= ")
        'Debug.WriteLine(rowsIndex)
        'ConnectTable.Row = 1
        If (StationConnectionStatus(TimerIndex) <> 0 And StationConnectionStatus(TimerIndex) <> -1) Then 'Dir(INISTNPATH(i)) = "" Then
            'ConnectTable.Row = TimerIndex
            'ConnectTable.Col = 3
            'ConnectTable.CellAlignment = 4
            'ConnectTable.CellBackColor = vbRed
            'ConnectTable.Text = "NO"
            rowsIndex = TimerIndex - 1
            ConnectTable.Rows(rowsIndex).Cells(3).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ConnectTable.Rows(rowsIndex).Cells(3).Style.BackColor = Color.Red
            ConnectTable.Rows(rowsIndex).Cells(3).Value = "NO"

            'ConnectTable.Col = 4
            'ConnectTable.Text = ""
            ConnectTable.Rows(rowsIndex).Cells(4).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ConnectTable.Rows(rowsIndex).Cells(4).Value = ""

            'ConnectTable.Col = 5
            'ConnectTable.Text = ""
            ConnectTable.Rows(rowsIndex).Cells(5).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ConnectTable.Rows(rowsIndex).Cells(5).Value = ""
            'ConnectTable.Col = 6
            'ConnectTable.Text = ""
            ConnectTable.Rows(rowsIndex).Cells(6).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ConnectTable.Rows(rowsIndex).Cells(6).Value = ""
            'GoTo Skip
        ElseIf StationConnectionStatus(TimerIndex) <> -1 Then
            'ConnectTable.Row = TimerIndex
            'ConnectTable.Col = 3
            'ConnectTable.CellAlignment = 4
            'ConnectTable.CellBackColor = vbGreen
            'ConnectTable.Text = "YES"
            rowsIndex = TimerIndex - 1
            ConnectTable.Rows(rowsIndex).Cells(3).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ConnectTable.Rows(rowsIndex).Cells(3).Style.BackColor = Color.Green
            ConnectTable.Rows(rowsIndex).Cells(3).Value = "YES"
        End If
        FileNum = FreeFile()
        'Open INISTNPATH(TimerIndex) For Input As #FileNum
        'Line Input #FileNum, statuscode
        'Close #FileNum

        FileOpen(FileNum, INISTNPATH(TimerIndex), OpenMode.Input)
        statuscode = LineInput(FileNum)
        FileClose(FileNum)

        pos1 = InStr(1, statuscode, ",")
        pos2 = InStr(pos1 + 1, statuscode, ",")
        pos3 = InStr(pos2 + 1, statuscode, ",")
        pos4 = InStr(pos3 + 1, statuscode, ",")

        'ConnectTable.Col = 4
        'ConnectTable.CellAlignment = 4
        'ConnectTable.Text = Mid(statuscode, 1, pos1 - 1)
        'ServerMonitorOP.WONos(i) = Mid(statuscode, 1, pos1 - 1)
        ConnectTable.Rows(rowsIndex).Cells(4).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        ConnectTable.Rows(rowsIndex).Cells(4).Value = Mid(statuscode, 1, pos1 - 1)
        'ServerMonitorOP.WONos(i) = Mid(statuscode, 1, pos1 - 1)

        'ConnectTable.Col = 5
        'ConnectTable.CellAlignment = 4
        'ConnectTable.Text = Mid(statuscode, pos1 + 1, (pos2 - pos1) - 1)
        ConnectTable.Rows(rowsIndex).Cells(5).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        ConnectTable.Rows(rowsIndex).Cells(5).Value = Mid(statuscode, pos1 + 1, (pos2 - pos1) - 1)

        'ConnectTable.Col = 6
        'ConnectTable.CellAlignment = 4
        'ConnectTable.Text = Mid(statuscode, pos2 + 1, (pos3 - pos2) - 1)
        ConnectTable.Rows(rowsIndex).Cells(6).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        ConnectTable.Rows(rowsIndex).Cells(6).Value = Mid(statuscode, pos2 + 1, (pos3 - pos2) - 1)

        'ConnectTable.Col = 7
        'ConnectTable.CellAlignment = 4
        'ConnectTable.Text = Mid(statuscode, pos3 + 1, (pos4 - pos3) - 1)
        ConnectTable.Rows(rowsIndex).Cells(7).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        ConnectTable.Rows(rowsIndex).Cells(7).Value = Mid(statuscode, pos3 + 1, (pos4 - pos3) - 1)

        'ConnectTable.Col = 8
        'ConnectTable.CellAlignment = 4
        'ConnectTable.Text = Mid(statuscode, pos4 + 1)

        ConnectTable.Rows(rowsIndex).Cells(8).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        ConnectTable.Rows(rowsIndex).Cells(8).Value = Mid(statuscode, pos4 + 1)
        'ServerMonitorOP.OutputQty(i) = Mid(statuscode, pos4 + 1)
Skip:
        'Next
    End Sub

    Private Sub LoadWOData()
        Dim RSCount As Integer
        Dim query As String
        Dim ds As DataTable

        On Error Resume Next
        query = "SELECT [WONOS],[WOMODELNAME],[WOQTY],[PFMODE],[UNITNOS],[DATECREATED],[DATECLOSED],[STATUS] FROM [CSUNIT]"
        ds = ConnectionDatabase.readData(query).Tables(0)
        WOGrid.DataSource = ds
    End Sub

    Private Sub Image1_Click(sender As Object, e As EventArgs) Handles Image1.Click
        Text6.Text = ""
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If RFID_Comm.IsOpen Then
            RFID_Comm.Close()
        End If
        If BarcodeScan_Comm.IsOpen Then
            BarcodeScan_Comm.Close()
        End If
        End
    End Sub

    Private Sub MonitoringToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MonitoringToolStripMenuItem.Click
        SelectedMode = 2
        Timer1.Enabled = True
        Frame5.Top = 20000
        Frame5.Left = 120
        Frame3.Top = 20000
        Frame3.Left = 120
        Frame4.Top = 20000
        Frame4.Left = 120
        Frame2.Top = 3
        Frame2.Left = 3
        Frame6.Top = 20000
        Frame6.Left = 120
        Frame7.Left = 120
        Frame7.Top = 20000
        Frame8.Left = 120
        Frame8.Top = 20000
        Frame9.Left = 120
        Frame9.Top = 20000
    End Sub

    Private Sub WOEntryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WOEntryToolStripMenuItem.Click
        SelectedMode = 1
        Timer1.Enabled = False
        Frame2.Top = 20000
        Frame2.Left = 120
        Frame5.Top = 20000
        Frame5.Left = 120
        Frame4.Top = 20000
        Frame4.Left = 120
        Frame3.Top = 3
        Frame3.Left = 3
        Frame6.Left = 120
        Frame6.Top = 20000
        Frame7.Left = 120
        Frame7.Top = 20000
        Frame8.Left = 120
        Frame8.Top = 20000
        Frame9.Left = 120
        Frame9.Top = 20000

        Label2.Text = "Place a Change Seies Dummy Unit on the Sensor and scan the Work Order Number..."
        'BarcodeScan_Comm.PortOpen = True
        'If RFID_Comm.PortOpen = False Then
        '    RFID_Comm.PortOpen = True
        'End If
        WOBuffer = ""
        CSAction = 0
        _Text1_0.Text = ""
        _Text1_1.Text = ""
        _Text1_2.Text = ""
        _Text1_3.Text = ""
        _Text1_4.Text = ""
        SelectedMode = 1
        CSAction = 0
        'Label2.Caption = "Place a Change Seies Dummy Unit on the Sensor and scan the Work Order Number..."
    End Sub

    Private Sub MSFlexGrid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles MSFlexGrid1.CellContentClick
        Dim TableRow As Integer
        Dim TableArticle As String
        Dim TableModel As String
        Dim PSNData As String
        Dim FileNum As Integer

        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso MSFlexGrid1(e.ColumnIndex, e.RowIndex).Value IsNot Nothing Then
            TableArticle = MSFlexGrid1(1, e.RowIndex).Value.ToString()
            'MessageBox.Show("Clicked Cell Value: " & TableArticle)

            FrmDisplay.Show()
            FrmDisplay.Label1.Text = TableModel
            FileNum = FreeFile()
            FileOpen(FileNum, INIPSNFOLDERPATH & TableArticle & ".Txt", OpenMode.Input)
            Do While Not EOF(FileNum)
                PSNData = LineInput(FileNum)
                'Debug.WriteLine(PSNData)
                FrmDisplay.Text1.Text = FrmDisplay.Text1.Text + PSNData + vbCrLf
            Loop
            FileClose(FileNum)
        End If
    End Sub

    Private Sub _Text1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles _Text1_3.KeyPress, _Text1_0.KeyPress, _Text1_1.KeyPress, _Text1_2.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim Rdtagcount As String
        Dim Rdtagnos As String
        Dim Rdtagref As String
        Dim RdtagWOnos As String
        Dim Rdtagqty As String
        Dim Rdtagop As String
        Dim dbNull As String = ""

        If _Text1_3.Focused Then
            If KeyAscii = 13 Then
                If _Text1_3.Text = "C" Then
                    _Label1_13.Visible = True
                    _Text1_4.Visible = True
                End If
            End If
        ElseIf _Text1_2.Focused Then
            If KeyAscii = 13 Then
                Label2.Text = "Please wait while Verifying the Tag..."

                Rdtagcount = CStr(CInt(RD_MULTI_RFID("004C", 3))) 'Tag Life Cycle
                Rdtagnos = RD_MULTI_RFID("0040", 3)
                Rdtagref = RD_MULTI_RFID("0014", 10)
                RdtagWOnos = RD_MULTI_RFID("0000", 10)
                Rdtagqty = RD_MULTI_RFID("0028", 10)
                Rdtagop = RD_MULTI_RFID("0046", 3) 'WO output counter at packaging
                CSInfo.CSWONOS = _Text1_0.Text
                CSInfo.CSWOMODEL = _Text1_1.Text
                CSInfo.CSWOQTY = _Text1_2.Text
                CSInfo.CSWOPFMODE = _Text1_3.Text
                Label5.Text = Rdtagcount
                Label2.Text = "Please wait while programming the Tag..."
                WOBuffer = ""
                'Program RFID tag
                Label2.Text = "Checking RFID Tag Write Cycle..."
                'If Rdtagcount < CSInfo.CSWOQTY Then
                '    Label2.Caption = "Write Cycle unable to support WO quantity. Use another Tag and try again"
                '_Text1_0.Text = ""
                '_Text1_1.Text = ""
                '_Text1_2.Text = ""
                '_Text1_3.Text = ""
                '_Text4_4.Text = ""
                '    Exit Sub
                'End If
                Label2.Text = "Writing Work Order Number..."
                If Not Clear_Tag("0000", 10) Then MsgBox("Unable to Clear Tag Address &H0000")
                If Not Wr_Tag(CSInfo.CSWONOS, "0000") Then
                    Label2.Text = "Unable to write to address 0000H. Verify and try again from step 1."
                    _Text1_0.Text = ""
                    _Text1_1.Text = ""
                    _Text1_2.Text = ""
                    _Text1_3.Text = ""
                    _Text4_4.Text = ""
                    GoTo EventExitSub
                End If
                Label2.Text = "Writing Work Order Reference..."
                If Not Clear_Tag("0014", 10) Then MsgBox("Unable to Clear Tag Address &H0014")
                If Not Wr_Tag(CSInfo.CSWOMODEL, "0014") Then
                    Label2.Text = "Unable to write to address 0014H. Verify and try again from Step 1."
                    _Text1_0.Text = ""
                    _Text1_1.Text = ""
                    _Text1_2.Text = ""
                    _Text1_3.Text = ""
                    _Text4_4.Text = ""
                    GoTo EventExitSub
                End If
                Label2.Text = "Writing Work Order Quantity..."
                If Not Clear_Tag("0028", 10) Then MsgBox("Unable to Clear Tag Address &H0028")
                If Not Wr_Tag(CSInfo.CSWOQTY, "0028") Then
                    Label2.Text = "Unable to write to address 0028H. Verify and try again from Step 1."
                    _Text1_0.Text = ""
                    _Text1_1.Text = ""
                    _Text1_2.Text = ""
                    GoTo EventExitSub
                End If
                Label2.Text = "Writing WO PF Mode..."
                If Not Wr_Tag(CSInfo.CSWOPFMODE, "0052") Then
                    Label2.Text = "Unable to write to address 0052H. Verify and try again from Step 1."
                    _Text1_0.Text = ""
                    _Text1_1.Text = ""
                    _Text1_2.Text = ""
                    _Text1_3.Text = ""
                    _Text4_4.Text = ""
                    GoTo EventExitSub
                End If
                Label2.Text = "Reseting packaging counter..."
                If Not Wr_Tag("000000", "0046") Then
                    Label2.Text = "Unable to write to address 0046H. Verify and try again from Step 1."
                    _Text1_0.Text = ""
                    _Text1_1.Text = ""
                    _Text1_2.Text = ""
                    _Text1_3.Text = ""
                    _Text4_4.Text = ""
                    GoTo EventExitSub
                End If
                Label2.Text = "Incrementing Tag cycle..."
                Rdtagcount = CStr(CDbl(Rdtagcount) + 1)
                Rdtagcount = Mid("000000", 1, 6 - Len(Rdtagcount)) & Rdtagcount
                If Not Wr_Tag(Rdtagcount, "004C") Then
                    Label2.Text = "Unable to write to address 0046H. Verify and try again from Step 1."
                    _Text1_0.Text = ""
                    _Text1_1.Text = ""
                    _Text1_2.Text = ""
                    _Text1_3.Text = ""
                    _Text4_4.Text = ""
                    GoTo EventExitSub
                End If
                Label2.Text = "RFID Tag programming completed."
                Label5.Text = CStr(CDbl(Rdtagcount) + 1)

                If CSInfo.CSWOPFMODE = "PFC" Then
                    Dim Query As String = "INSERT INTO [CSUNIT] ([UNITNOS],[WONOS],[WOMODELNAME],[WOQTY],[PFMODE],[LOGISTICC],[STATUS],[DATECREATED],[DATECLOSED])
                                           VALUE (ISNULL('" & Rdtagnos & "','" & dbNull & "'),ISNULL('" & CSInfo.CSWONOS & "','" & dbNull & "'),ISNULL('" & CSInfo.CSWOMODEL & "','" & dbNull & "'),
                                           ISNULL('" & CSInfo.CSWOQTY & "','" & dbNull & "'),ISNULL('" & CSInfo.CSWOPFMODE & "','" & dbNull & "'),ISNULL('" & CSInfo.CSWOLC & "','" & dbNull & "'),
                                          ISNULL('OPEN','" & dbNull & "'),ISNULL('" & Today & "," & TimeOfDay & "','" & dbNull & "'),"
                    If ConnectionDatabase.insertData(Query) Then
                        MsgBox("Success add WO!")
                    Else
                        MsgBox("Failed add WO!")
                    End If
                Else
                    Dim Query As String = "INSERT INTO [CSUNIT] ([UNITNOS],[WONOS],[WOMODELNAME],[WOQTY],[PFMODE],[LOGISTICC],[STATUS],[DATECREATED],[DATECLOSED])
                                           VALUE (ISNULL('" & Rdtagnos & "','" & dbNull & "'),ISNULL('" & CSInfo.CSWONOS & "','" & dbNull & "'),ISNULL('" & CSInfo.CSWOMODEL & "','" & dbNull & "'),
                                           ISNULL('" & CSInfo.CSWOQTY & "','" & dbNull & "'),ISNULL('" & CSInfo.CSWOPFMODE & "','" & dbNull & "'),ISNULL('','" & dbNull & "'),
                                          ISNULL('OPEN','" & dbNull & "'),ISNULL('" & Today & "," & TimeOfDay & "','" & dbNull & "'),"
                    If ConnectionDatabase.insertData(Query) Then
                        MsgBox("Success add WO!")
                    Else
                        MsgBox("Failed add WO!")
                    End If
                End If


                'Creating PSN Folder for this WO
                'MkDir INIPSNFOLDERPATH & CSInfo.CSWONOS
                'RFID_Comm.PortOpen = False
                GoTo EventExitSub
            End If
        End If
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Label6.Text = "Reading information from tag..."
        If VB.Left(RD_MULTI_RFID("0040", 3), 3) = "TAG" Then
            Label6.Text = "Reading tag information..."
            _Text2_4.Text = RD_MULTI_RFID("0040", 3)
            _Text2_0.Text = RD_MULTI_RFID("0000", 10)
            _Text2_1.Text = RD_MULTI_RFID("0014", 10)
            _Text2_2.Text = RD_MULTI_RFID("0028", 10)
            _Text2_3.Text = RD_MULTI_RFID("0046", 3)
            _Text2_5.Text = RD_MULTI_RFID("0052", 3)
            Timer2.Enabled = False
            Label6.Text = ""
            Cmd_WOClosure.Enabled = True
        End If
        Label6.Text = ""
    End Sub

    Private Sub Cmd_ForceClose_Click(sender As Object, e As EventArgs) Handles Cmd_ForceClose.Click
        Dim DBDel As dao.Database
        'UPGRADE_WARNING: Arrays in structure RSDel may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim RSDel As dao.Recordset
        Dim FileNum As Short
        Dim Tagnos As String
        Dim TagWOnos As String
        Dim TagWOref As String
        Dim TagWOQty As String
        Dim TagWOOutput As String
        Dim TagPFMode As String
        Dim query As String
        Dim ds As DataTable

        Cmd_WOClosure.Enabled = False
        Cmd_ForceClose.Enabled = False
        Tagnos = _Text2_4.Text
        TagWOnos = _Text2_0.Text
        TagWOref = _Text2_1.Text
        TagWOQty = _Text2_2.Text
        TagWOOutput = _Text2_3.Text
        TagPFMode = _Text2_5.Text

        'Erasing all data from Tag
        Label6.Text = "Please wait, Clearing tag info..."
        If Not Clear_Tag("0000", 10) Then 'WO Nos
            MsgBox("Unable to Clear Tag Address - &H0000")
            Cmd_WOClosure.Enabled = True
            Cmd_ForceClose.Enabled = True
            Exit Sub
        End If
        If Not Clear_Tag("0014", 10) Then 'WO Ref
            MsgBox("Unable to Clear Tag Address - &H0014")
            Cmd_ForceClose.Enabled = True
            Cmd_WOClosure.Enabled = True
            Exit Sub
        End If
        If Not Clear_Tag("0028", 10) Then 'WO Quantity
            MsgBox("Unable to Clear Tag Address - &H0028")
            Cmd_ForceClose.Enabled = True
            Cmd_WOClosure.Enabled = True
            Exit Sub
        End If
        If Not Clear_Tag("0046", 3) Then 'Packaging Counter
            MsgBox("Unable to Clear Tag Address - &H0046")
            Cmd_ForceClose.Enabled = True
            Cmd_WOClosure.Enabled = True
            Exit Sub
        End If
        If Not Clear_Tag("0052", 10) Then 'PF Mode
            MsgBox("Unable to Clear Tag Address - &H0052")
            Cmd_ForceClose.Enabled = True
            Cmd_WOClosure.Enabled = True
            Exit Sub
        End If

        Label6.Text = "Please wait, deleting Work Order info in server"

        query = "UPDATE [CSUNIT] SET [STATUS] = 'FORCED',[DATECLOSED] = '" & Today & "," & TimeOfDay & "' Where [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.updateData(query) Then
            'MsgBox("Success save & update database!")
        Else
            'MsgBox("Failed save & update database!")
        End If
        On Error Resume Next

        query = "DELETE FROM [STN1] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success delete database!")
        Else
            'MsgBox("Failed delete database!")
        End If

        query = "DELETE FROM [STN2] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success delete database!")
        Else
            'MsgBox("Failed delete database!")
        End If

        query = "DELETE FROM [STN3] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success save & delete database!")
        Else
            'MsgBox("Failed save & delete database!")
        End If

        query = "DELETE FROM [STN4] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success save & delete database!")
        Else
            'MsgBox("Failed save & delete database!")
        End If

        query = "DELETE FROM [STN5] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success save & delete database!")
        Else
            'MsgBox("Failed save & delete database!")
        End If

        query = "DELETE FROM [TESTER] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success save & delete database!")
        Else
            'MsgBox("Failed save & delete database!")
        End If


        'Log data to CSV file
        FileNum = FreeFile()
        FileOpen(FileNum, INILOGPATH & "DeletedWO" & Year(Today) & ".CSV", OpenMode.Append)
        PrintLine(FileNum, _Text2_4.Text & "," & _Text2_0.Text & "," & _Text2_1.Text & "," & _Text2_2.Text & "," & _Text2_3.Text & "," & _Text2_5.Text)
        FileClose(FileNum)
        Label6.Text = "Work Order - " & TagWOnos & "Forced to Close"
        Cmd_ForceClose.Enabled = True
        Cmd_WOClosure.Enabled = True
    End Sub

    Private Sub Cmd_WOClosure_Click(sender As Object, e As EventArgs) Handles Cmd_WOClosure.Click
        Dim query As String
        Dim ds As DataTable
        Dim FileNum As Short
        Dim Tagnos As String
        Dim TagWOnos As String
        Dim TagWOref As String
        Dim TagWOQty As String
        Dim TagWOOutput As String
        Dim TagPFMode As String

        Cmd_ForceClose.Enabled = False
        Cmd_WOClosure.Enabled = False
        Tagnos = _Text2_4.Text
        TagWOnos = _Text2_0.Text
        TagWOref = _Text2_1.Text
        TagWOQty = _Text2_2.Text
        TagWOOutput = _Text2_3.Text
        TagPFMode = _Text2_5.Text

        'Check if WO Qty  = Packaging Output Qty
        If CDbl(_Text2_2.Text) <> CDbl(_Text2_3.Text) Then
            MsgBox("Unable to close WO due to wrong quantity")
            Cmd_WOClosure.Enabled = True
            Exit Sub
        End If
        Label6.Text = "Please wait, Clearing tag info..."
        If Not Clear_Tag("0000", 10) Then
            MsgBox("Unable to Clear Tag Address - &H0000")
            Cmd_WOClosure.Enabled = True
            Cmd_ForceClose.Enabled = True
            Exit Sub
        End If
        If Not Clear_Tag("0014", 10) Then
            MsgBox("Unable to Clear Tag Address - &H0014")
            Cmd_WOClosure.Enabled = True
            Cmd_ForceClose.Enabled = True
            Exit Sub
        End If
        If Not Clear_Tag("0028", 10) Then
            MsgBox("Unable to Clear Tag Address - &H0028")
            Cmd_WOClosure.Enabled = True
            Cmd_ForceClose.Enabled = True
            Exit Sub
        End If
        If Not Clear_Tag("0046", 3) Then
            MsgBox("Unable to Clear Tag Address - &H0046")
            Cmd_WOClosure.Enabled = True
            Cmd_ForceClose.Enabled = True
            Exit Sub
        End If
        If Not Clear_Tag("0052", 10) Then
            MsgBox("Unable to Clear Tag Address - &H0052")
            Cmd_WOClosure.Enabled = True
            Cmd_ForceClose.Enabled = True
            Exit Sub
        End If
        Cmd_WOClosure.Enabled = True
        Cmd_ForceClose.Enabled = True
        Label6.Text = "Please wait, Closing Work Order info in server"

        query = "UPDATE [CSUNIT] SET [STATUS] = 'CLOSED',[DATECLOSED] = '" & Today & "," & TimeOfDay & "' Where [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.updateData(query) Then
            'MsgBox("Success save & update database!")
        Else
            'MsgBox("Failed save & update database!")
        End If

        On Error Resume Next

        query = "DELETE FROM [STN1] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success delete database!")
        Else
            'MsgBox("Failed delete database!")
        End If

        query = "DELETE FROM [STN2] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success delete database!")
        Else
            'MsgBox("Failed delete database!")
        End If

        query = "DELETE FROM [STN3] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success save & delete database!")
        Else
            'MsgBox("Failed save & delete database!")
        End If

        query = "DELETE FROM [STN4] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success save & delete database!")
        Else
            'MsgBox("Failed save & delete database!")
        End If

        query = "DELETE FROM [STN5] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success save & delete database!")
        Else
            'MsgBox("Failed save & delete database!")
        End If

        query = "DELETE FROM [TESTER] WHERE [WONOS] = '" & _Text2_0.Text & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success save & delete database!")
        Else
            'MsgBox("Failed save & delete database!")
        End If

        FileNum = FreeFile()
        FileOpen(FileNum, INILOGPATH & Year(Today) & ".CSV", OpenMode.Append)
        PrintLine(FileNum, _Text2_4.Text & "," & _Text2_0.Text & "," & _Text2_1.Text & "," & _Text2_2.Text & "," & _Text2_3.Text & "," & _Text2_5.Text)
        FileClose(FileNum)
        Label6.Text = "Work Order - " & TagWOnos & " Closed"
    End Sub

    Private Sub WOClosureToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WOClosureToolStripMenuItem.Click
        SelectedMode = 0
        Timer1.Enabled = False
        Frame3.Top = 20000
        Frame3.Left = 120
        Frame5.Top = 20000
        Frame5.Left = 120
        Frame4.Top = 3
        Frame4.Left = 3
        Frame2.Top = 20000
        Frame2.Left = 120
        Frame6.Top = 20000
        Frame6.Left = 120
        Frame7.Left = 120
        Frame7.Top = 20000
        Frame8.Left = 120
        Frame8.Top = 20000
        Frame9.Left = 120
        Frame9.Top = 20000

        'RFID_Comm.PortOpen = True
        Timer2.Enabled = True
    End Sub

    Private Sub WOMasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WOMasterToolStripMenuItem.Click
        SelectedMode = 0
        ReloadCombo()
        Frame2.Top = 20000
        Frame2.Left = 120
        Frame5.Top = 3
        Frame5.Left = 3
        Frame4.Top = 20000
        Frame4.Left = 120
        Frame3.Top = 20000
        Frame3.Left = 120
        Frame6.Top = 20000
        Frame6.Left = 120
        Frame7.Left = 120
        Frame7.Top = 20000
        Frame8.Left = 120
        Frame8.Top = 20000
        Frame9.Left = 120
        Frame9.Top = 20000

        'RFID_Comm.PortOpen = True
    End Sub

    Private Sub Cmd_ProgramMaster_Click(sender As Object, e As EventArgs) Handles Cmd_ProgramMaster.Click
        If Combo1.Text = "" Then
            MsgBox("Please select a reference")
            Exit Sub
        End If

        If IsNumeric(Text3.Text) = False Then
            MsgBox("Invalid Quantity")
            Exit Sub
        End If

        'Check if the tag on the sensor is a defined Master Tag
        If RD_MULTI_RFID("0000", 10) <> "MASTER" Then
            MsgBox("This Tag is not a Master Tag. Unable to program.")
            Exit Sub
        End If
        'If Not Wr_Tag("Master", "0000") Then
        '    MsgBox "Unable to write to Master RFID WONos"
        '    Exit Sub
        'End If
        'Label2.Caption = "Writing Work Order Qty"
        If Not Wr_Tag((Text3.Text), "0028") Then
            MsgBox("Unable to write to Master RFID Qty")
            Exit Sub
        End If
        'Label2.Caption = "Writing Work Order reference"
        If Not Wr_Tag((Combo1.Text), "0014") Then
            MsgBox("Unable to write to Master RFID Reference")
            Exit Sub
        End If
        'If Not Wr_Tag("TAG999", "0040") Then
        '    MsgBox "Unable to write to Master RFID Tag Name"
        '    Exit Sub
        'End If
        Label3.Text = "Master programming completed"
        ClearCombo()
        'RFID_Comm.PortOpen = False
    End Sub

    Private Sub WODistrupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WODistrupToolStripMenuItem.Click
        Dim N As Object
        Dim i As Object
        SelectedMode = 0
        Timer1.Enabled = False
        Frame2.Top = 20000
        Frame2.Left = 120
        Frame5.Top = 20000
        Frame5.Left = 120
        Frame4.Top = 20000
        Frame4.Left = 120
        Frame3.Top = 20000
        Frame3.Left = 120
        Frame6.Left = 120
        Frame6.Top = 20000
        Frame7.Left = 3
        Frame7.Top = 3
        Frame8.Left = 120
        Frame8.Top = 20000
        Frame9.Left = 120
        Frame9.Top = 20000

        Dim FileNum As Short
        Dim tempcode As String
        Dim pos2, pos1, pos3 As Object
        Dim pos4 As Short
        Dim Stnwonos(9) As String
        Dim Stndata(9) As String
        Dim CommentInput As String

        Combo3.Items.Clear()
        On Error Resume Next
        'load the list of WO nos that are current in the assy line
        FileNum = FreeFile()

        For i = 3 To 9
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(FileNum, INISTNPATH(i), OpenMode.Input)
            tempcode = LineInput(FileNum)
            FileClose(FileNum)
            'UPGRADE_WARNING: Couldn't resolve default property of object pos1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pos1 = InStr(1, tempcode, ",")
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Stnwonos(i) = Mid(tempcode, 1, pos1 - 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For N = 1 To i
                'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Stnwonos(i) = Stnwonos(N - 1) Then GoTo SkipAdd
            Next
            'to screen out already distrupted WO
            Combo3.Items.Add(Stnwonos(i))
SkipAdd:

        Next
        MSFlexGrid2.Rows.Clear()
        MSFlexGrid2.Columns.Clear()
        LoadDistrupTable()
        LoadDistrupWOData()
    End Sub

    Private Sub ManualOverwriteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ManualOverwriteToolStripMenuItem.Click
        Timer1.Enabled = False
        Frame3.Top = 20000
        Frame3.Left = 120
        Frame5.Top = 20000
        Frame5.Left = 120
        Frame4.Top = 20000
        Frame4.Left = 120
        Frame2.Top = 20000
        Frame2.Left = 120
        Frame6.Top = 3
        Frame6.Left = 3
        Frame7.Left = 120
        Frame7.Top = 20000
        Frame8.Left = 120
        Frame8.Top = 20000
        Frame9.Left = 120
        Frame9.Top = 20000
    End Sub

    Private Sub Cmd_Distrip_Click(sender As Object, e As EventArgs) Handles Cmd_Distrip.Click
        Dim i As Object
        Dim DB As dao.Database
        'UPGRADE_WARNING: Arrays in structure RS may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim RS As dao.Recordset
        Dim FileNum As Short
        Dim tempcode As String
        Dim pos2, pos1, pos3 As Object
        Dim pos4 As Short
        Dim Stnwonos(9) As String
        Dim Stnref(9) As String
        Dim Stnwoqty(9) As String
        Dim Stnop(9) As String
        Dim Stndata(9) As String
        Dim CommentInput As String
        Dim DBNull As String = ""
        Dim query As String
        Dim ds As DataTable
        FileNum = FreeFile()

        For i = 3 To 9
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(FileNum, INISTNPATH(i), OpenMode.Input)
            tempcode = LineInput(FileNum)
            FileClose(FileNum)
            'UPGRADE_WARNING: Couldn't resolve default property of object pos1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pos1 = InStr(1, tempcode, ",")
            'UPGRADE_WARNING: Couldn't resolve default property of object pos1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pos2 = InStr(pos1 + 1, tempcode, ",")
            'UPGRADE_WARNING: Couldn't resolve default property of object pos2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pos3 = InStr(pos2 + 1, tempcode, ",")
            'pos4 = InStr(pos3 + 1, tempcode, ",")

            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Stnwonos(i) = Mid(tempcode, 1, pos1 - 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Stnref(i) = Mid(tempcode, pos1 + 1, (pos2 - pos1) - 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Stnwoqty(i) = Mid(tempcode, pos2 + 1, (pos3 - pos2) - 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Stnop(i) = Mid(tempcode, pos3 + 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Stnwonos(i) = Combo3.Text Then 'check if current station is working on this WO
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Stndata(i) = tempcode
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Stndata(i) = "-"
            End If
        Next

        CommentInput = InputBox("Please enter comment")

        query = "INSERT INTO [ONGOING] ([REMARK],[WONos],[Date],[SA1WONos],[SA2WONos],[Stn1WONos],[Stn2WONos],
                 [Stn3WONos],[Stn4WONos],[Stn5WONos],[Stn6WONos],[TesterWONos])
                 VALUES (ISNULL('" & CommentInput & "','" & DBNull & "'),ISNULL('" & Combo3.Text & "','" & DBNull & "'),ISNULL('" & Today & "," & TimeOfDay & "','" & DBNull & "'),ISNULL('" & Stndata(1) & "','" & DBNull & "'),ISNULL('" & Stndata(2) & "','" & DBNull & "'),ISNULL('" & Stndata(3) & "','" & DBNull & "'),ISNULL('" & Stndata(4) & "','" & DBNull & "'),
                 ISNULL('" & Stndata(5) & "','" & DBNull & "'),ISNULL('" & Stndata(6) & "','" & DBNull & "'),ISNULL('" & Stndata(8) & "','" & DBNull & "'),ISNULL('" & Stndata(9) & "','" & DBNull & "'),ISNULL('" & Stndata(7) & "','" & DBNull & "'))"

        If ConnectionDatabase.insertData(query) Then
            'MsgBox("Succes add database!")
        Else
            'MsgBox("Failed add database!")
        End If

        MSFlexGrid2.Rows.Clear()
        MSFlexGrid2.Columns.Clear()
        LoadDistrupTable()
        LoadDistrupWOData()

        'query = "UPDATE [CSUNIT] SET [STATUS] = 'DISTRUP' WHERE [WONOS] = '" & Combo3.Text & "' "
        'If ConnectionDatabase.updateData(query) Then
        '    'MsgBox("Success save & update database!")
        'Else
        '    'MsgBox("Failed save & update database!")
        'End If
    End Sub

    Private Sub Cmd_Relaunch_Click(sender As Object, e As EventArgs) Handles Cmd_Relaunch.Click
        Dim selectrows As DataGridViewSelectedRowCollection
        Dim selectrow As DataGridViewRow
        Dim FileNum As Short
        Dim inputstr As String
        Dim SelectedWONos As String

        'LAter to include a check if the Station completed the current WO QTY before allowing Re-introduce

        'UPGRADE_WARNING: Couldn't resolve default property of object selectrow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        selectrows = MSFlexGrid2.SelectedRows
        selectrow = selectrows(0)
        SelectedWONos = selectrow.Cells(2).Value
        MsgBox(SelectedWONos)
        'Main PCBA Assy
        FileNum = FreeFile()
        inputstr = selectrow.Cells(5).Value
        If inputstr <> "-" Then
            FileOpen(FileNum, INISTNPATH(3), OpenMode.Output)
            PrintLine(FileNum, inputstr)
            FileClose(FileNum)
        End If
        'ElectroMagnet and Locking PCBA Assy
        FileNum = FreeFile()
        inputstr = selectrow.Cells(6).Value
        If inputstr <> "-" Then
            FileOpen(FileNum, INISTNPATH(4), OpenMode.Output)
            PrintLine(FileNum, inputstr)
            FileClose(FileNum)
        End If
        'PCBA to Body Assy
        FileNum = FreeFile()
        inputstr = selectrow.Cells(7).Value
        If inputstr <> "-" Then
            FileOpen(FileNum, INISTNPATH(5), OpenMode.Output)
            PrintLine(FileNum, inputstr)
            FileClose(FileNum)
        End If
        'Head to Body Screwing
        FileNum = FreeFile()
        inputstr = selectrow.Cells(8).Value
        If inputstr <> "-" Then
            FileOpen(FileNum, INISTNPATH(6), OpenMode.Output)
            PrintLine(FileNum, inputstr)
            FileClose(FileNum)
        End If
        'Tester
        FileNum = FreeFile()
        inputstr = selectrow.Cells(9).Value
        If inputstr <> "-" Then
            FileOpen(FileNum, INISTNPATH(7), OpenMode.Output)
            PrintLine(FileNum, inputstr)
            FileClose(FileNum)
        End If
        'Station 5 -Cover Assy
        FileNum = FreeFile()
        inputstr = selectrow.Cells(10).Value
        If inputstr <> "-" Then
            FileOpen(FileNum, INISTNPATH(8), OpenMode.Output)
            PrintLine(FileNum, inputstr)
            FileClose(FileNum)
        End If
        'Packaging
        FileNum = FreeFile()
        inputstr = selectrow.Cells(11).Value
        If inputstr <> "-" Then
            FileOpen(FileNum, INISTNPATH(9), OpenMode.Output)
            PrintLine(FileNum, inputstr)
            FileClose(FileNum)
        End If

        Dim DComment As String
        Dim query As String
        DComment = selectrow.Cells(1).Value
        MsgBox(DComment)

        query = "DELETE FROM [ONGOING] WHERE [Date] = '" & DComment & "'"
        If ConnectionDatabase.deleteData(query) Then
            'MsgBox("Success delete from database")
        Else
            'MsgBox("Failed delete from database")
        End If

        query = "UPDATE [CSUNIT] SET [STATUS] = 'OPEN' WHERE [WONOS] = '" & SelectedWONos & "'"

        If ConnectionDatabase.updateData(query) Then
            'MsgBox("Success update from database")
        Else
            'MsgBox("Failed update from database")
        End If

        MSFlexGrid2.Rows.Clear()
        MSFlexGrid2.Columns.Clear()
        LoadDistrupTable()
        LoadDistrupWOData()
    End Sub

    Private Sub Cmd_readTag_Click(sender As Object, e As EventArgs) Handles Cmd_readTag.Click
        Dim clean As Control
        Cmd_readTag.Enabled = False
        For i As Integer = 0 To 7
            clean = Me.Controls("_Text4_" & i)
            clean.Text = ""
        Next
        _Text4_0.Text = RD_MULTI_RFID("0052", 3)
        _Text4_1.Text = RD_MULTI_RFID("0000", 10)
        _Text4_2.Text = RD_MULTI_RFID("0014", 10)
        _Text4_3.Text = RD_MULTI_RFID("0028", 10)
        _Text4_4.Text = RD_MULTI_RFID("0046", 3) 'WO output counter at packaging
        _Text4_6.Text = RD_MULTI_RFID("004C", 3) 'Tag Life Cycle
        _Text4_7.Text = RD_MULTI_RFID("0040", 3)
        Cmd_readTag.Enabled = True
    End Sub

    Private Sub Cmd_ManualWrite_Click(sender As Object, e As EventArgs) Handles Cmd_ManualWrite.Click
        Cmd_ManualWrite.Enabled = False
        Select Case Combo2.Text
            Case "WO Nos"
                If Not Clear_Tag("0000", 10) Then MsgBox("Unable to Clear Tag Address &H0000")
                If Text5.Text = "" Then Exit Sub
                If Len(Text5.Text) > 20 Then Exit Sub
                If Not Wr_Tag((Text5.Text), "0000") Then
                    MsgBox("Unable to Write Tag Address &H0000")
                End If
            Case "WO Reference"
                If Not Clear_Tag("0014", 10) Then MsgBox("Unable to Clear Tag Address &H0014")
                If Text5.Text = "" Then Exit Sub
                If Len(Text5.Text) > 20 Then Exit Sub
                If Not Wr_Tag((Text5.Text), "0014") Then
                    MsgBox("Unable to Write Tag Address &H0014")
                End If

            Case "WO Quantity"
                If Not Clear_Tag("0028", 10) Then MsgBox("Unable to Clear Tag Address &H0028")
                If Text5.Text = "" Then Exit Sub
                If Len(Text5.Text) > 20 Then Exit Sub
                If Not Wr_Tag((Text5.Text), "0028") Then
                    MsgBox("Unable to Write Tag Address &H0028")
                End If

            Case "Tag ID"
                If Not Clear_Tag("0040", 3) Then MsgBox("Unable to Clear Tag Address &H0040")
                If Text5.Text = "" Then Exit Sub
                If Len(Text5.Text) > 6 Then Exit Sub
                If Not Wr_Tag((Text5.Text), "0040") Then
                    MsgBox("Unable to Write Tag Address &H0040")
                End If

            Case "WO Output"
                If Not Clear_Tag("0046", 3) Then MsgBox("Unable to Clear Tag Address &H0046")
                If Text5.Text = "" Then Exit Sub
                If Len(Text5.Text) > 6 Then Exit Sub
                If Not Wr_Tag((Text5.Text), "0046") Then
                    MsgBox("Unable to Write Tag Address &H0046")
                End If

            Case "Tag Cycle"
                If Not Clear_Tag("004C", 3) Then MsgBox("Unable to Clear Tag Address &H004C")
                If Text5.Text = "" Then Exit Sub
                If Len(Text5.Text) > 6 Then Exit Sub
                If Not Wr_Tag((Text5.Text), "004C") Then
                    MsgBox("Unable to Write Tag Address &H004C")
                End If

            Case "PF Mode"
                If Not Clear_Tag("0052", 2) Then MsgBox("Unable to Clear Tag Address &H0052")
                If Text5.Text = "" Then Exit Sub
                If Len(Text5.Text) > 4 Then Exit Sub
                If Not Wr_Tag((Text5.Text), "0052") Then
                    MsgBox("Unable to Write Tag Address &H0052")
                End If

        End Select
        Cmd_ManualWrite.Enabled = True
    End Sub

    Private Sub LabelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LabelToolStripMenuItem.Click
        Timer1.Enabled = False
        Me.Hide()
        FrmLabelSpec.ShowDialog()
    End Sub

    Private Sub Cmd_Rack_Click(sender As Object, e As EventArgs) Handles SubAssy1ToolStripMenuItem.Click, SubAssy2ToolStripMenuItem.Click, SubAssy3ToolStripMenuItem.Click, Station1ToolStripMenuItem.Click, Station2ToolStripMenuItem.Click, Station3ToolStripMenuItem.Click, Station4ToolStripMenuItem.Click, Station5ToolStripMenuItem.Click, Station6ToolStripMenuItem.Click, SubAssyConnectorToolStripMenuItem.Click
        Dim selectedMenuItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim temp As String
        Dim Linestr As String
        Dim FNum As Short
        Dim i As Integer
        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        If selectedMenuItem.Text = "SubAssy1" Then
            temp = "SubAssy1"
        ElseIf selectedMenuItem.Text = "SubAssy2" Then
            temp = "SubAssy2"
        ElseIf selectedMenuItem.Text = "SubAssy3" Then
            temp = "SubAssy3"
        ElseIf selectedMenuItem.Text = "Station1" Then
            temp = "Station1"
        ElseIf selectedMenuItem.Text = "Station2" Then
            temp = "Station2"
        ElseIf selectedMenuItem.Text = "Station3" Then
            temp = "Station3"
        ElseIf selectedMenuItem.Text = "Station4" Then
            temp = "Station4"
        ElseIf selectedMenuItem.Text = "Station5" Then
            temp = "Station5"
        ElseIf selectedMenuItem.Text = "Station6" Then
            temp = "Station6"
        ElseIf selectedMenuItem.Text = "Sub Assy - Connector" Then
            temp = "SA_Connector"
        End If
        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir(INIMATERIALPATH & "Rack\" & temp) = "" Then
            MsgBox("Unable to locate file")
            Exit Sub
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = 1
        FNum = FreeFile()
        FileOpen(FNum, INIMATERIALPATH & "Rack\" & temp, OpenMode.Input)
        Do While Not EOF(FNum)
            Linestr = LineInput(FNum)
            FrmRack.Txt_slot(i).Text = Linestr
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            i = i + 1
        Loop
        FileClose(FNum)
        FrmRack.Label3.Text = temp
        Me.Hide()
        FrmRack.ShowDialog()
    End Sub

    Private Sub DataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataToolStripMenuItem.Click
        BarcodeScan_Comm.Close()
        Me.Hide()
        FrmDebug.ShowDialog()
    End Sub

    Private Sub Cmd_Fstn_Click(sender As Object, e As EventArgs) Handles SubAssy1ToolStripMenuItem1.Click, SubAssy2ToolStripMenuItem1.Click, SubAssy3ToolStripMenuItem1.Click, Station1ToolStripMenuItem1.Click, Station2ToolStripMenuItem1.Click, Station3ToolStripMenuItem1.Click, Station4ToolStripMenuItem1.Click, Station5ToolStripMenuItem1.Click, Station6ToolStripMenuItem1.Click, FailureMsgToolStripMenuItem.Click
        Dim Index As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim i As Integer
        Dim FNum As Short
        Dim Linestr As String
        Dim pos As String

        Timer1.Enabled = False
        On Error Resume Next
        Frame5.Top = 20000
        Frame5.Left = 120
        Frame3.Top = 20000
        Frame3.Left = 120
        Frame4.Top = 20000
        Frame4.Left = 120
        Frame2.Top = 20000
        Frame2.Left = 120
        Frame6.Top = 20000
        Frame6.Left = 120
        Frame7.Left = 120
        Frame7.Top = 20000
        Frame8.Left = 3
        Frame8.Top = 3
        Frame9.Left = 120
        Frame9.Top = 20000

        MSFlexGrid3.Rows.Clear()
        MSFlexGrid3.Columns.Clear()
        LoadFailureTable()

        Select Case Index.Text
            Case "SubAssy1"
                Label8.Text = "SubAssy1"
            Case "SubAssy2"
                Label8.Text = "SubAssy2"
            Case "SubAssy3"
                Label8.Text = "SubAssy3"
            Case "Station1"
                Label8.Text = "Station1"
            Case "Station2"
                Label8.Text = "Station2"
            Case "Station3"
                Label8.Text = "Station3"
            Case "Station4"
                Label8.Text = "Station4"
            Case "Station5"
                Label8.Text = "Station5"
            Case "Station6"
                Label8.Text = "Station6"
        End Select
        i = 0

        FNum = FreeFile()
        If File.Exists(INIFAILCODEPATH & Label8.Text & ".Txt") Then
            'Debug.WriteLine("FILE EXISTS")
            FileOpen(FNum, INIFAILCODEPATH & Label8.Text & ".Txt", OpenMode.Input)
            Do While Not EOF(FNum)
                Linestr = LineInput(FNum)

                pos = InStr(1, Linestr, ",")
                MSFlexGrid3.Rows(i).Cells(1).Value = Mid(Linestr, 1, CDbl(pos) - 1).ToString()
                MSFlexGrid3.Rows(i).Cells(2).Value = Mid(Linestr, CDbl(pos) + 1).ToString()
                i += 1
            Loop
            FileClose(FNum)
            Cmd_Save.Enabled = True
        Else
            MsgBox("FILE NOT EXISTS")
            Cmd_Save.Enabled = False
        End If


    End Sub

    Private Sub LoadFailureTable()
        MSFlexGrid3.Columns.Add("Column0", "S/No")
        MSFlexGrid3.Columns.Add("Column1", "Fail Code")
        MSFlexGrid3.Columns.Add("Column2", "Descriptions")
        For i As Integer = 1 To 20
            MSFlexGrid3.Rows.Add()
        Next
    End Sub

    Private Sub MSFlexGrid3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles MSFlexGrid3.CellContentClick
        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 Then
            FrmEnter.Show()
            If EntryCode = "" Then Exit Sub
            MSFlexGrid3.Text = EntryCode
        End If

    End Sub

    Private Sub Cmd_Save_Click(sender As Object, e As EventArgs) Handles Cmd_Save.Click
        Dim N As Object
        Dim i As Object
        Dim Linestr(100) As String
        Dim FileNum As Short
        Dim ErrorNos(100) As String
        Dim ErrorDescrip(100) As String

        For i = 0 To 100

            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ErrorNos(i) = MSFlexGrid3.Rows(i).Cells(1).Value
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If ErrorNos(i) = "" Then Exit For
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ErrorDescrip(i) = MSFlexGrid3.Rows(i).Cells(2).Value
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Linestr(i) = ErrorNos(i) & "," & ErrorDescrip(i)
        Next
        FileNum = FreeFile()
        FileOpen(FileNum, INIFAILCODEPATH & Label8.Text & ".Txt", OpenMode.Output)
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For N = 0 To i - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(FileNum, Linestr(N))
        Next
        FileClose(FileNum)
        MsgBox("DONE!")
    End Sub

    Private Sub PSNToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PSNToolStripMenuItem.Click
        Frame2.Top = 20000
        Frame2.Left = 120
        Frame5.Top = 20000
        Frame5.Left = 120
        Frame4.Top = 20000
        Frame4.Left = 120
        Frame3.Top = 20000
        Frame3.Left = 120
        Frame6.Left = 120
        Frame6.Top = 20000
        Frame7.Left = 120
        Frame7.Top = 20000
        Frame8.Left = 120
        Frame8.Top = 20000
        Frame9.Left = 3
        Frame9.Top = 3
        SelectedMode = 3
    End Sub

    Private Sub MaterialToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaterialToolStripMenuItem.Click
        Timer1.Enabled = False
        Me.Hide()
        FrmMaterial.ShowDialog()
    End Sub

    Private Sub TimerBusy_Tick(sender As Object, e As EventArgs) Handles TimerBusy.Tick

    End Sub
End Class