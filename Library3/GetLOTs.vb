Imports System.Data.SqlClient

Public Module GetLots
    'запрос получения доступных ErrorCodes
    'Public ErrorCodesList As String = "use FAS SELECT ID, Category+Code as ErrorCode, [Description] FROM [FAS].[dbo].[M_ErrorCode]"
    Private SQL As String
    'Функция вывода информации в Лэйбл

    Public Sub PrintLabel(Lab As Label, Message As String, Col As Color)
        Lab.Text = Message
        Lab.Location = New Point(13, 349)
        Lab.Font = New Font("Microsoft Sans Serif", 20, FontStyle.Bold)
        Lab.ForeColor = Col
        Lab.Visible = True
    End Sub
    Public Sub PrintLabel(Lab As Label, Message As String, x As Integer, y As Integer, Col As Color)
        Lab.Text = Message
        Lab.Location = New Point(x, y)
        Lab.Font = New Font("Microsoft Sans Serif", 20, FontStyle.Bold)
        Lab.ForeColor = Col
        Lab.Visible = True
    End Sub
    Public Sub PrintLabel(Lab As Label, Message As String, x As Integer, y As Integer, Col As Color, Font As Font)
        Lab.Text = Message
        Lab.Location = New Point(x, y)
        Lab.Font = Font 'Font("Microsoft Sans Serif", 20, FontStyle.Bold)
        Lab.ForeColor = Col
        Lab.Visible = True
    End Sub
    'запрос обновления списка лотов для LOTManagment спутник
    Public LotList_LOTManagment As String = " use fas
         SELECT LOTCode as LOT
        ,FULL_LOT_Code as 'FULL LOT'
        ,FAS_Models.ModelName as MODEL
        ,L.CreateDate  as 'CREATE DATE'
        ,(select count (*) from FAS_SerialNumbers where FAS_SerialNumbers.IsUsed = 0 and FAS_SerialNumbers.LOTID = L.LOTID) as Ready
        ,(select count (*) from FAS_SerialNumbers where FAS_SerialNumbers.IsUsed = 1 and FAS_SerialNumbers.LOTID = L.LOTID) as Used
        ,(select count (*) from FAS_SerialNumbers where FAS_SerialNumbers.LOTID = L.LOTID) as InLOT
        ,L.LOTID
        FROM [FAS].[dbo].[FAS_GS_LOTs] as L
        left join FAS_Models ON L.ModelID = FAS_Models.ModelID
        where isactive = 1"
    'запрос информация о лоте для LOTManagment спутник
    'переделать 
    Public GetLotInfo As String = " Use FAS
        SELECT LOTID,LOTCode,FULL_LOT_Code,M.ModelName,Specification,Manufacturer,Operator,MarketID,PTID,BoxCapacity,PalletCapacity,
        LiterIndex,IsHDCPUpload,IsCertUpload,IsMACUpload,W.Scenario,Lb.Scenario,CreateDate,u.UserName,[ModelCheck],[SWRead],[SWGS1Read]
        FROM [FAS].[dbo].[FAS_GS_LOTs] as L
        left join M_Models as M on L.ModelID = M.ModelID
        left join M_WorkingScenario as W on w.WSID = L.WorkingScenarioID
        left join M_LabelScenario as Lb on Lb.ID = L.LabelScenarioID
        left join M_Users as U on U.UserID = L.CreateByID
        where lotid = "

    'запрос обновления списка лотов для FAS Start Station спутник
    Public Function GetLotList_FASStart_GS(DG_LotList As DataGridView) As DataGridView
        SQL = "Use FAS
        SELECT LOTCode as LOT
        ,FULL_LOT_Code as 'FULL LOT'
        ,M.ModelName as MODEL
        ,(select count (*) from FAS_SerialNumbers as SN where SN.LOTID = L.LOTID) as InLOT
        ,(select count (*) from FAS_SerialNumbers as SN where SN.IsUsed = 0 and SN.LOTID = L.LOTID) as Ready
        ,(select count (*) from FAS_SerialNumbers as SN where SN.IsUsed = 1 and SN.LOTID = L.LOTID) as Used
        ,LOTID
        ,Scenario
        FROM [FAS].[dbo].[FAS_GS_LOTs] as L
        left join FAS_Models as M ON L.ModelID = M.ModelID
        left join FAS_LabelScenario as Lab ON L.LabelScenarioID = Lab.ID
        where L.IsActive = 1
        order by LOT desc"
        LoadGridFromDB(DG_LotList, SQL)
        Return DG_LotList
    End Function

    'запрос обновления списка лотов для Upload Station спутник
    Public Function GetLotList_UpStation_GS(DG_LotList As DataGridView) As DataGridView
        SQL = "Use FAS
             SELECT LOTCode as LOT
            ,FULL_LOT_Code as 'FULL LOT'
            ,M.ModelName as MODEL
            --,FORMAT(M_LOTs.CreateDate,'dd.MM.yyyy HH:mm:ss', 'de-de') as 'CREATE DATE'
            ,(select count (*) from FAS_SerialNumbers AS SN where SN.LOTID = L.LOTID) as InLOT
            ,(select count (*) from FAS_SerialNumbers AS SN where SN.IsUploaded = 0 and SN.LOTID = L.LOTID and IsUsed = 1) as Ready
            ,(select count (*) from FAS_SerialNumbers AS SN where SN.IsUploaded = 1 and SN.LOTID = L.LOTID and IsUsed = 1) as Used
            ,LOTID
            ,[IsHDCPUpload]
            ,[IsCertUpload]
            ,[IsMACUpload]
            ,ModelCheck
            ,SWRead
            ,SWGS1Read
            ,[LabelScenarioID]
            ,PTID
            FROM [FAS].[dbo].[FAS_GS_LOTs] as L
            left join FAS_Models as M ON L.ModelID = M.ModelID
            where L.IsActive = 1"
        LoadGridFromDB(DG_LotList, SQL)
        Return DG_LotList
    End Function

    Public Function GetLotList_Disassembly_GS(DG_LotList As DataGridView) As DataGridView
        SQL = "Use FAS
             SELECT LOTCode as LOT
            ,FULL_LOT_Code as 'FULL LOT'
            ,M.ModelName as MODEL
            ,(select count (*) from FAS_SerialNumbers as SN where SN.LOTID = L.LOTID) as InLOT
            ,(select count (*) from FAS_SerialNumbers as SN where SN.IsUsed = 0 and SN.LOTID = L.LOTID) as Ready
            ,(select count (*) from FAS_SerialNumbers as SN where SN.IsUsed = 1 and SN.LOTID = L.LOTID) as Used
            ,LOTID
            FROM [FAS].[dbo].[FAS_GS_LOTs] as L
            left join FAS_Models as M ON L.ModelID = M.ModelID
            where IsActive = 1"
        LoadGridFromDB(DG_LotList, SQL)
        Return DG_LotList
    End Function


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

    'запрос обновления списка лотов для FAS END
    Public Function GetLotList_FASEND_GS(DG_LotList As DataGridView) As DataGridView
        SQL = "Use FAS
        SELECT LOTCode as LOT
        ,FULL_LOT_Code as 'FULL LOT'
        ,LiterIndex
        ,M.ModelName as MODEL
        ,[BoxCapacity]
        ,[PalletCapacity]
        ,(select count (*) from FAS_SerialNumbers as SN where SN.LOTID = L.LOTID) as InLOT
        ,(select count (*) from FAS_SerialNumbers as SN where SN.IsUploaded = 1 and SN.LOTID = L.LOTID and SN.IsPacked = 0) as Ready
        ,(select count (*) from FAS_SerialNumbers as SN where SN.IsPacked = 1 and SN.LOTID = L.LOTID) as Used
        ,LOTID
        FROM [FAS].[dbo].[FAS_GS_LOTs] as L
        left join FAS_Models as M ON L.ModelID = M.ModelID
        where L.IsActive = 1"
        LoadGridFromDB(DG_LotList, SQL)
        Return DG_LotList
    End Function


    ' 'запрос обновления списка лотов для Cadena
    ' Public LotList_LotCadena As String = "use fas
    '      SELECT [LOTCode] as LOT      
    '   ,[FullLOTCode] as 'Full LOT'
    '   ,LiterIndex
    ',M.ModelName  as Model
    '   ,[BoxCapacity]
    '   ,[PalletCapacity]
    '   ,(select count (*) from M_CadenaID where M_CadenaID.LOTID =L.ID) as InLOT
    '   ,(select count (*) from M_CadenaID where M_CadenaID.IsUsed = 1 and M_CadenaID.LOTID = L.ID and M_CadenaID.IsPacked = 0) as Ready
    ',(select count (*) from M_CadenaID where M_CadenaID.IsPacked = 1 and M_CadenaID.LOTID = L.ID) as Used
    '   ,ID
    '   FROM [FAS].[dbo].[M_LOT_Cadena] as L
    '   left join M_Models as M ON L.ModelID = M.ModelID
    '   where isactiv = 1"

    'getlot for Scanning station
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
    Public Function GetCurrentContractLot(LOTID As Integer) As ArrayList
        'SQL = "USE FAS
        'SELECT m.ModelName,[FullLOTCode],[SMTNumberFormat],[SMTRangeChecked],[SMTStartRange],[SMTEndRange],[ParseLog],[StepSequence]
        ',l.BoxCapacity, l.PalletCapacity, l.LiterIndex
        'FROM [FAS].[dbo].[Contract_LOT] as L
        'left join FAS_Models as M On m.ModelID = L.ModelID
        'where IsActive = 1 and id > 20053 and ID = " & LOTID & "
        'order by id desc"


        SQL = "USE FAS SELECT m.ModelName,[FullLOTCode]
        ,[CheckFormatSN_SMT],[SMTNumberFormat],[SMTRangeChecked],[SMTStartRange],[SMTEndRange]
        ,[CheckFormatSN_FAS],[FASNumberFormat],[FASRangeChecked],[FASStartRange],[FASEndRange]
        ,[SingleSN],[ParseLog],[StepSequence]
        ,[BoxCapacity],[PalletCapacity],[LiterIndex],[HexFasSN]
        FROM [FAS].[dbo].[Contract_LOT] as L
        left join FAS_Models as M On m.ModelID = L.ModelID
        where IsActive = 1 and id > 20053 and ID = " & LOTID & "
        order by id desc"
        Return SelectListString(SQL)
    End Function



End Module
