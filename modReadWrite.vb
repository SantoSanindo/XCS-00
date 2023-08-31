Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Public Module modReadWrite
    Public INICSINFOPATH As String 'Path to ChangeSerialInfo.mdb
    Public INIDATABASEPATH As String 'Path to Model.mdb
    Public INIUSERPATH As String 'Path to User.mdb
    Public INISTNPCIP(11) As String
    Public INISTNPLCIP(11)
    Public INISTNPATH(10) As String 'Path to all station Status#n.txt
    Public INISTNINFO As String
    Public INIPSNFOLDERPATH As String 'Path to all PSN files
    Public INIPSNACHIEVEPATH As String 'Backup path to all PSN files
    Public INIMATERIALPATH As String 'Path to Material files
    Public INILOGPATH As String 'Path to all closed WO
    Public INIDISTRUPPATH As String
    Public INILABELPHOTOPATH As String
    Public INIFAILCODEPATH As String
    Public FNum As Integer
    Public Linestr As String
    Public Sub ReadINI(Filename As String)
        Dim itemStr As String
        Dim SectionHeading As String
        Dim pos As Integer

        FNum = FreeFile()

        'If Dir(Filename) Then
        'SetDefaultINIValues
        'WriteINI
        'End If

        FileOpen(FNum, Filename, OpenMode.Input)

        While Not EOF(FNum)
            Linestr = LineInput(FNum)

            If Left(Linestr, 1) = "[" Then
                SectionHeading = Mid(Linestr, 2, Len(Linestr) - 2)

            Else
                If InStr(Linestr, "=") > 0 Then
                    pos = InStr(Linestr, "=")
                    itemStr = Left$(Linestr, pos - 1)

                    Select Case UCase(SectionHeading)
                        Case "DATABASE PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INIDATABASEPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "USER PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INIUSERPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "PSN PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INIPSNFOLDERPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "ACHIEVE PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INIPSNACHIEVEPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "MATERIAL PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INIMATERIALPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "WOLOG PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INILOGPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "DISTRUPT PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INIDISTRUPPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "LABEL PHOTO PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INILABELPHOTOPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "FAILURE CODE PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INIFAILCODEPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "SA1 PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(1) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "SA2 PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(2) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN1 PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(3) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN2 PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(4) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN3 PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(5) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN4 PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(6) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "TESTER PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(7) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN5 PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(8) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN6 PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(9) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN CONNECTOR PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPATH(10) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "CSUNIT PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INICSINFOPATH = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN1 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(1) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN1 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(1) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN2 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(2) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN2 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(2) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN3 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(3) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN3 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(3) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN4 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(4) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN4 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(4) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN5 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(5) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN5 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(5) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN6 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(6) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN6 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(6) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN7 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(7) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN7 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(7) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN8 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(8) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN8 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(8) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN9 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(9) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN9 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(9) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN10 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(10) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN10 PLC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPLCIP(10) = Mid$(Linestr, pos + 1)
                            End Select

                        Case "STN11 PC IP PATH"
                            Select Case UCase(itemStr)
                                Case "PATH" : INISTNPCIP(11) = Mid$(Linestr, pos + 1)
                            End Select
                    End Select
                End If
            End If

        End While
        FileClose(FNum)
    End Sub

    Public Sub ClearPSNVar()
        PSNFileInfo.BodyAssyCheckIn = ""
        PSNFileInfo.BodyAssyCheckOut = ""
        PSNFileInfo.BodyAssyStatus = ""
        PSNFileInfo.ConnTestCheckIn = ""
        PSNFileInfo.ConnTestCheckOut = ""
        PSNFileInfo.ConnTestStatus = ""
        PSNFileInfo.DateCompleted = ""
        PSNFileInfo.DateCreated = ""
        PSNFileInfo.DebugComment = ""
        PSNFileInfo.DebugStatus = ""
        PSNFileInfo.DebugTechnican = ""
        PSNFileInfo.ElectroMagnet = ""
        PSNFileInfo.FTCheckIn = ""
        PSNFileInfo.FTCheckOut = ""
        PSNFileInfo.MainPCBA = ""
        PSNFileInfo.FTStatus = ""
        PSNFileInfo.ModelName = ""
        PSNFileInfo.OperatorID = ""
        PSNFileInfo.PackagingCheckIn = ""
        PSNFileInfo.PackagingCheckOut = ""
        PSNFileInfo.PackagingStatus = ""
        PSNFileInfo.PSN = ""
        PSNFileInfo.RepairDate = ""
        PSNFileInfo.ScrewStnCheckIn = ""
        PSNFileInfo.ScrewStnCheckOut = ""
        PSNFileInfo.ScrewStnStatus = ""
        PSNFileInfo.SecondaryPCBA = ""
        PSNFileInfo.Stn5CheckIn = ""
        PSNFileInfo.Stn5CheckOut = ""
        PSNFileInfo.Stn5Status = ""
        PSNFileInfo.Vacumm2CheckOut = ""
        PSNFileInfo.VacummCheckOut = ""
        PSNFileInfo.Vacuum2CheckIn = ""
        PSNFileInfo.VacuumCheckIn = ""
        PSNFileInfo.VacuumStatus = ""
        PSNFileInfo.Vacuum2Status = ""
        PSNFileInfo.WONos = ""

    End Sub

End Module
