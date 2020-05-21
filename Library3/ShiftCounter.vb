Public Module GetShiftCounter
    Private SQL As String
    Public ShiftCounter As Integer
    Public ShiftCounterID, ShiftID, CurrentDate As String
    Public ShiftPapameters(2) As String
    'универсальный запрос в базу для ShiftCounterID и  ShiftCounter
    Public Function SQLQuery(Value As String, StationID As Integer, idApp As Integer, ShiftID As String, CurrentDate As String)
        SQL = "USE FAS
            SELECT " & Value & " FROM [FAS].[dbo].[FAS_ShiftsCounter]
      where StationID = " & StationID & " and ID_App = " & idApp & " and ShiftID = " & ShiftID & "  and FORMAT(CreateDate,'d', 'de-de')='" & CurrentDate & "'"
        Return SelectString(SQL)

    End Function


    Public Function ShiftCounterStart(CurentTimeSec As Integer, StationID As Integer, idApp As Integer) As Array
        'определение номера смены (дневная или ночная)
        Dim i As Integer = CurentTimeSec
        Select Case i
            Case 0 To 86399 ' запуск в период С 00:00:00 ДО 23:59:59
                ShiftID = 3 'при отсутствии ночных смен
                CurrentDate = DateTime.Today
        End Select
        'поиск выпуска в текущую смену
        ShiftCounterID = SQLQuery("ID", StationID, idApp, ShiftID, CurrentDate)
        If ShiftCounterID <> "" Then
            'если запись для текущей смены уже существует
            ShiftCounter = SQLQuery("ShiftCounter", StationID, idApp, ShiftID, CurrentDate)
        Else
            'если записи нет, то делаем новую запись в таблицу счетчика за смену
            SQL = "USE FAS
            insert into [FAS].[dbo].[FAS_ShiftsCounter] (StationID,ID_App,ShiftID,ShiftCounter,CreateDate) 
		    values (" & StationID & "," & idApp & "," & ShiftID & ",0,CURRENT_TIMESTAMP)"
            RunCommand(SQL)
            'инициализируем счетчик за смену в программе
            'повторно запрашиваем ShiftcounterID 
            ShiftCounterID = SQLQuery("ID", StationID, idApp, ShiftID, CurrentDate)
            ShiftCounter = 0
        End If
        ShiftPapameters = {ShiftCounterID, ShiftCounter}
        Return ShiftPapameters
    End Function

    Public Sub ShiftCounterUpdate(ShiftCounter As String, ShiftCounterID As String)
        SQL = " Update [FAS].[dbo].[FAS_ShiftsCounter] set ShiftCounter = " & ShiftCounter & " where id  = " & ShiftCounterID
        RunCommand(SQL)
    End Sub

End Module