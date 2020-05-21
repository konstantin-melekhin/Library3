Public Module ContractModule
    'запрос обновления списка лотов для FAS END
    Public ContractLotList As String = "Use FAS
        SELECT L.LOTCode as LOT
	    ,L.FullLOTCode as 'FULL LOT'
	    ,M.ModelName as MODEL
		,l.BoxCapacity
		,l.PalletCapacity
		,L.LiterIndex as 'LITER INDEX'
	  	,(select count (*) from ContractTemp_SMT_SN AS SMT_SN where SMT_SN.LOTID = L.ID) as SMTSN_InLOT
		,[SMTStartRange]
        ,[SMTEndRange]
		,(select count (*) from ContractTemp_SMT_SN AS SMT_SN where SMT_SN.IsUsed = 0 and SMT_SN.LOTID = L.ID) as Ready
		,(select count (*) from ContractTemp_SMT_SN AS SMT_SN where SMT_SN.IsUsed = 1 and SMT_SN.LOTID = L.ID) as Used		

		,(select count (*) from ContractTemp_FAS_SN AS FAS_SN where FAS_SN.LOTID = L.ID) as FASSN_InLOT
		,[FASStartRange]
        ,[FASEndRange]
		,(select count (*) from ContractTemp_FAS_SN AS FAS_SN where FAS_SN.IsUsed = 0 and FAS_SN.LOTID = L.ID) as Ready
		,(select count (*) from ContractTemp_FAS_SN AS FAS_SN where FAS_SN.IsUsed = 1 and FAS_SN.LOTID = L.ID) as Used		
		,[CheckFormatSN_SMT]
        ,[CheckFormatSN_FAS]
		,[SMTRangeChecked]
		,[FASRangeChecked]
		,L.ID
	    FROM [FAS].[dbo].[Contract_LOT] as L
        left join M_Models as M ON L.ModelID = M.ModelID
        where L.isactive = 1
		order by L.ID desc"

End Module
