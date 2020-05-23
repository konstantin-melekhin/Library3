Public Module GetShiftCounter
    Private SQL As String

    Public Function ShiftInfo(StationID As Integer, idApp As Integer, ShiftID As Integer) As ArrayList
        'поиск выпуска в текущую смену
        Dim ShiftCounterInfo As New ArrayList(SelectListString("USE FAS
            SELECT [ID],[ShiftCounter],[LOT_Counter] FROM [FAS].[dbo].[FAS_ShiftsCounter]
            where StationID = " & StationID & " and ID_App = " & idApp & " and 
            ShiftID = " & ShiftID & "  and FORMAT(CreateDate,'d', 'de-de')='" & DateTime.Today & "'"))
        Return ShiftCounterInfo
    End Function

    Public Function ShiftCounterStart(StationID As Integer, idApp As Integer, LOTID As Integer) As ArrayList
        Dim ShiftID As Integer
        'определение номера смены (дневная или ночная)
        'Dim i As Integer = CurentTimeSec
        Select Case DateAndTime.Timer
            Case 0 To 86399 ' запуск в период С 00:00:00 ДО 23:59:59
                ShiftID = 3 'при отсутствии ночных смен
        End Select
        'поиск выпуска в текущую смену
        Dim ShiftCounterInfo As New ArrayList(ShiftInfo(StationID, idApp, ShiftID))
        If ShiftCounterInfo.Count = 0 Then
            'если записи нет, то делаем новую запись в таблицу счетчика за смену
            SQL = "USE FAS
            insert into [FAS].[dbo].[FAS_ShiftsCounter] (StationID,ID_App,ShiftID,ShiftCounter,CreateDate, LOTID, LOT_Counter) 
		    values (" & StationID & "," & idApp & "," & ShiftID & ",0,CURRENT_TIMESTAMP," & ShiftID & ",0 )"
            RunCommand(SQL)
            'повторно запрашиваем ShiftcounterID 
            ShiftCounterInfo = ShiftInfo(StationID, idApp, ShiftID)
        End If
        Return ShiftCounterInfo
    End Function

    Public Sub ShiftCounterUpdate(ShiftCounterID As String, ShiftCounter As String, LotCounter As String)
        SQL = " Use FAS Update [FAS].[dbo].[FAS_ShiftsCounter] set ShiftCounter = " & ShiftCounter & ",
                LOT_Counter = " & LotCounter & " where id  = " & ShiftCounterID
        RunCommand(SQL)
    End Sub

End Module