Public Module GetShiftCounter
    Private SQL As String

    Public Function ShiftInfo(StationID As Integer, idApp As Integer, ShiftID As Integer, LOTID As Integer) As ArrayList
        'поиск выпуска в текущую смену
        Dim ShiftCounterInfo As New ArrayList(SelectListString("USE FAS
            SELECT [ID],[ShiftCounter],[LOT_Counter],[PassLOTRes],[FAilLOTRes] FROM [FAS].[dbo].[FAS_ShiftsCounter]
            where StationID = " & StationID & " and ID_App = " & idApp & " and LOTID = " & LOTID & " and 
            ShiftID = " & ShiftID & "  and FORMAT(CreateDate,'d', 'de-de')='" & DateTime.Today & "'"))
        Return ShiftCounterInfo
    End Function

    Public Function ShiftCounterStart(StationID As Integer, idApp As Integer, LOTID As Integer) As ArrayList
        Dim ShiftID As Integer
        'определение номера смены (дневная или ночная)
        Select Case DateAndTime.Timer
            Case 0 To 86399 ' запуск в период С 00:00:00 ДО 23:59:59
                ShiftID = 3 'при отсутствии ночных смен
        End Select
        'поиск выпуска в текущую смену
        Dim ShiftCounterInfo As New ArrayList(ShiftInfo(StationID, idApp, ShiftID, LOTID))
        If ShiftCounterInfo.Count = 0 Then
            'если записи нет, то делаем новую запись в таблицу счетчика за смену
            SQL = "USE FAS
            insert into [FAS].[dbo].[FAS_ShiftsCounter] (StationID,ID_App,ShiftID,ShiftCounter,CreateDate, LOTID, LOT_Counter,PassLOTRes,FAilLOTRes) 
		    values (" & StationID & "," & idApp & "," & ShiftID & ",0,CURRENT_TIMESTAMP," & LOTID & ",0,0,0)"
            RunCommand(SQL)
            'повторно запрашиваем ShiftcounterID 
            ShiftCounterInfo = ShiftInfo(StationID, idApp, ShiftID, LOTID)
        End If
        Return ShiftCounterInfo
    End Function

    Public Sub ShiftCounterUpdateCT(ShiftCounterID As Integer, ShiftCounter As Integer, LotCounter As Integer, PassLOTRes As Integer, FAilLOTRes As Integer)
        SQL = " Use FAS Update [FAS].[dbo].[FAS_ShiftsCounter] set ShiftCounter = " & ShiftCounter & ",LOT_Counter = " & LotCounter & "
        ,PassLOTRes = " & PassLOTRes & ",FAilLOTRes = " & FAilLOTRes & " where id  = " & ShiftCounterID
        RunCommand(SQL)
    End Sub

End Module