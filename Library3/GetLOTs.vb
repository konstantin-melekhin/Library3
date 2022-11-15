Imports System.Data.SqlClient

Public Module GetLots
    'запрос получения доступных ErrorCodes
    'Public ErrorCodesList As String = "use FAS SELECT ID, Category+Code as ErrorCode, [Description] FROM [FAS].[dbo].[M_ErrorCode]"
    Private SQL As String

    Public Sub AddToOperLogFasStart(PCBID As String, LineID As Integer, StationID As Integer, IDApp As Integer, UserID As Integer)
        SQL = "use fas insert into [FAS].[dbo].[FAS_OperationLog] ([PCBID], [ProductionAreaID], [StationID], [ApplicationId],
        [StateCodeDate], [StateCodeByID]) values
        (" & PCBID & "," & LineID & "," & StationID & "," & IDApp & ", CURRENT_TIMESTAMP, " & UserID & ")"
        RunCommand(SQL)
    End Sub

    Public Sub AddToOperLogFasUpload(PCBID As String, LineID As Integer, StationID As Integer, IDApp As Integer, UserID As Integer,
                            SN As String, RePrint As Boolean, ReUpload As Boolean, OldLabelDate As String, SCID As String)
        SQL = "use fas insert into [FAS].[dbo].[FAS_OperationLog] ([PCBID], [ProductionAreaID], [StationID], [ApplicationId],
        [StateCodeDate], [StateCodeByID], [SerialNumber], [RePrint], [ReUpload], [OldLabelDate], [SmartCardId]) values
        (" & PCBID & "," & LineID & "," & StationID & "," & IDApp & ", CURRENT_TIMESTAMP, " & UserID & ", " & SN & ", 
        '" & RePrint & "', '" & ReUpload & "', " & OldLabelDate & ", " & SCID & ")"
        RunCommand(SQL)
    End Sub

    Public Sub AddToOperLogFasDisassembly(PCBID As String, LineID As Integer, StationID As Integer, IDApp As Integer, UserID As Integer, SN As String)
        SQL = "use fas insert into [FAS].[dbo].[FAS_OperationLog] ([PCBID], [ProductionAreaID], [StationID], [ApplicationId],
        [StateCodeDate], [StateCodeByID], [SerialNumber]) values
        (" & PCBID & "," & LineID & "," & StationID & "," & IDApp & ", CURRENT_TIMESTAMP, " & UserID & ", " & SN & ")"
        RunCommand(SQL)
    End Sub

    Public Sub AddToOperLogFasEnd(PCBID As String, LineID As Integer, StationID As Integer, IDApp As Integer, UserID As Integer, SN As String)
        SQL = "use fas insert into [FAS].[dbo].[FAS_OperationLog] ([PCBID], [ProductionAreaID], [StationID], [ApplicationId],
        [StateCodeDate], [StateCodeByID], [SerialNumber]) values
        (" & PCBID & "," & LineID & "," & StationID & "," & IDApp & ", CURRENT_TIMESTAMP, " & UserID & ", " & SN & ")"
        RunCommand(SQL)
    End Sub

    Public Function GetLotList_ContractStation(DG_LotList As DataGridView) As DataGridView
        SQL = "use fas
        SELECT [Specification],[FullLOTCode],M.ModelName,[ID]
        FROM [FAS].[dbo].[Contract_LOT] as L
        left join FAS_Models as M On m.ModelID = L.ModelID
        where L.IsActive = 1 and id > 20053
        order by id desc"
        LoadGridFromDB(DG_LotList, SQL)
        Return DG_LotList
    End Function

    Public Function GetLotList_ContractStation(DG_LotList As DataGridView, CustamerID As Integer) As DataGridView
        SQL = $"use fas
        SELECT [Specification],[FullLOTCode],M.ModelName,[ID]
        FROM [FAS].[dbo].[Contract_LOT] as L
        left join FAS_Models as M On m.ModelID = L.ModelID
        where [СustomersID] = {CustamerID} and L.IsActive = 1 and id > 20053
        order by id desc"
        LoadGridFromDB(DG_LotList, SQL)
        Return DG_LotList
    End Function
    Public Function GetLotList_ContractStation(DG_LotList As DataGridView, CustamerID As Integer, DS As DataSet) As DataGridView
        SQL = $"use fas
        SELECT [Specification],[FullLOTCode],M.ModelName,[ID]
        FROM [FAS].[dbo].[Contract_LOT] as L
        left join FAS_Models as M On m.ModelID = L.ModelID
        where [СustomersID] = {CustamerID} and L.IsActive = 1 and id > 20053
        order by id desc"
        LoadGridFromDB2(DG_LotList, SQL, DS)
        Return DG_LotList
    End Function

    Public Function GetCurrentContractLot(LOTID As Integer) As ArrayList
        SQL = $"USE FAS SELECT 
        m.ModelName,[FullLOTCode],[CheckFormatSN_SMT],[SMTNumberFormat],[SMTRangeChecked],[SMTStartRange],[SMTEndRange]
        ,[CheckFormatSN_FAS],[FASNumberFormat],[FASRangeChecked],[FASStartRange],[FASEndRange]
        ,[SingleSN],[ParseLog],[StepSequence]
        ,[BoxCapacity],[PalletCapacity],[LiterIndex],[HexFasSN],[FASNumberFormat2],[СustomersID]
        FROM [FAS].[dbo].[Contract_LOT] as L
        left join FAS_Models as M On m.ModelID = L.ModelID
        where IsActive = 1 and id > 20053 and ID = {LOTID}
        order by id desc"
        Return SelectListString(SQL)
    End Function



    Public Function GetLotList_SDTV(DG_LotList As DataGridView) As DataGridView
        SQL = "use fas
        SELECT  [Full_LOT_Name],[LOT],[Specification],M.ModelName,[LOTID]
--,[ModelID],[IsActive],[IsHDCPUpload],[IsCertUpload]
--,[IsMACUpload],[ModelCheck],[SWRead],[SWGS1Read],[Manufacture],[Operator],[MarketID],[PTID]
FROM [FAS].[dbo].[SDTV_LOT] as L
left join FAS_Models as M On m.ModelID = L.ModelID
where IsActive = 1
order by L.lotid desc"
        LoadGridFromDB(DG_LotList, SQL)
        Return DG_LotList
    End Function

    Public Function GetCurrent_SDTV(LOTID As Integer) As ArrayList
        SQL = "use fas
        SELECT  [Full_LOT_Name],[LOT],[Specification],M.ModelName,[LOTID]
,l.[ModelID],[IsActive],[IsHDCPUpload],[IsCertUpload]
,[IsMACUpload],[ModelCheck],[SWRead],[SWGS1Read],[Manufacture],[Operator],[MarketID],[PTID]
FROM [FAS].[dbo].[SDTV_LOT] as L
left join FAS_Models as M On m.ModelID = L.ModelID
where LOTID = " & LOTID & " and IsActive = 1
order by L.lotid desc"
        Return SelectListString(SQL)
    End Function


End Module
