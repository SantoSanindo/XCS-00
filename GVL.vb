Module GVL
    Public CSInfo As ChangeSeriesInfo
    Public Login As Boolean
    Public BusyFlag As Boolean
    Public Parameter As ControlSpec
    Public ServerMonitorOP As Stnstatus
    Public PSNFileInfo As PSNText
    Public EntryCode As String



    Public Structure ControlSpec
        Dim UnitTagNos As String
        Dim UnitPartNos As String
        Dim UnitModel As String
        Dim UnitType As String
        Dim UnitFunction As String
        Dim UnitTension As String
        Dim UnitContact1_WO_Trig As String
        Dim UnitContact2_WO_Trig As String
        Dim UnitContact3_WO_Trig As String
        Dim UnitContact4_WO_Trig As String
        Dim UnitContact5_WO_Trig As String
        Dim UnitContact6_WO_Trig As String
        Dim UnitContact_WO_Trig As Long
        Dim UnitContact1_W_Key As String
        Dim UnitContact2_W_Key As String
        Dim UnitContact3_W_Key As String
        Dim UnitContact4_W_Key As String
        Dim UnitContact5_W_Key As String
        Dim UnitContact6_W_Key As String
        Dim UnitContact_W_Key As Long

        Dim UnitContact1_W_Key_Ten As String
        Dim UnitContact2_W_Key_Ten As String
        Dim UnitContact3_W_Key_Ten As String
        Dim UnitContact4_W_Key_Ten As String
        Dim UnitContact5_W_Key_Ten As String
        Dim UnitContact6_W_Key_Ten As String
        Dim UnitContact_W_Key_Ten As Long
        Dim UnitLabelTemplate As String
        Dim UnitLabelPhoto As String
        Dim UnitSycUL As Double
        Dim UnitSycLL As Double
    End Structure

    Public Structure ChangeSeriesInfo
        Dim CSWONOS As String
        Dim CSWOMODEL As String
        Dim CSWOQTY As String
        Dim CSWOPFMODE As String 'PFC/PFS
        Dim CSWOLC As String 'Logistic Center
    End Structure

    Public Structure Stnstatus
        Dim OutputQty As String()
        Dim WONos As String()
        Dim CSUnit As String()
        Dim WOQTY As String()
    End Structure

    Structure PSNText
        Dim ModelName As String
        Dim DateCreated As String
        Dim DateCompleted As String
        Dim OperatorID As String
        Dim WONos As String
        Dim MainPCBA As String
        Dim SecondaryPCBA As String
        Dim ElectroMagnet As String
        Dim PSN As String
        Dim BodyAssyCheckIn As String
        Dim BodyAssyCheckOut As String
        Dim BodyAssyStatus As String
        Dim ScrewStnCheckIn As String
        Dim ScrewStnCheckOut As String
        Dim ScrewStnStatus As String
        Dim FTCheckIn As String
        Dim FTCheckOut As String
        Dim FTStatus As String
        Dim Stn5CheckIn As String
        Dim Stn5CheckOut As String
        Dim Stn5Status As String
        Dim VacuumCheckIn As String
        Dim VacummCheckOut As String
        Dim VacuumStatus As String
        Dim ConnTestCheckIn As String
        Dim ConnTestCheckOut As String
        Dim ConnTestStatus As String
        Dim Vacuum2CheckIn As String
        Dim Vacumm2CheckOut As String
        Dim Vacuum2Status As String
        Dim PackagingCheckIn As String
        Dim PackagingCheckOut As String
        Dim PackagingStatus As String
        Dim DebugStatus As String
        Dim DebugComment As String
        Dim DebugTechnican As String
        Dim RepairDate As String
    End Structure

End Module
