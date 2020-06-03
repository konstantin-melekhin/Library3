Public Module PackingModule
    Public Function GetLastPack(LOTID As Integer, LineID As Integer) As ArrayList
        Dim LastPackCounter As ArrayList = New ArrayList(SelectListString("use fas
        SELECT PalletCounter,BoxCounter,UnitCounter 
        FROM [FAS].[dbo].[FAS_PackingCounter] where lotid = " & LOTID & " and LineID = " & LineID))

        If LastPackCounter.Count = 0 Then
            RunCommand("use fas
            insert into [FAS].[dbo].[FAS_PackingCounter] (PalletCounter,BoxCounter,UnitCounter,LineID,LOTID) values (1,1,1," & LineID & "," & LOTID & ")")
            LastPackCounter = New ArrayList() From {1, 1, 1}
        End If
        Return LastPackCounter
    End Function
End Module
