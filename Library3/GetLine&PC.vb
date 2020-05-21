Imports System.Data.SqlClient



Public Module GetLineAndPC
    Private SQL As String
    'функция захвата названия компьютера
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Public Function GetPCName() As String
        Dim strBuffer As String
        Dim strAns As Long
        strBuffer = Space(255)
        strAns = GetComputerName(strBuffer, 255)
        Return strBuffer
    End Function

    'функция поиска ID станции
    Public Function GetStationID(StationName As String) As String
        SQL = "use FAS
        use FAS SELECT StationID  FROM  [FAS].[dbo].[FAS_Stations] where StationName = '" & StationName & "'"
        Return SelectString(SQL)
    End Function

    'функция регистрации новой станции
    Public Function StationRegister(StationName As String) As String
        SQL = "use FAS
        insert into [FAS].[dbo].[FAS_Stations] (StationName, CreateDate) values ('" & StationName & "', CURRENT_TIMESTAMP) "
        RunCommand(SQL)
    End Function

    'функция отображения номера линии
    Public Function GetLine(StationID As String, IDApp As String) As String
        SQL = "use FAS
        select l.LineName  FROM [FAS].[dbo].[FAS_App_ListForPC] as  App
        left join FAS_Lines as L ON l.LineID = App.lineID
        where StationID = " & StationID & " and app_ID = " & IDApp
        Return SelectString(SQL)
    End Function

    'функция отображения / определения LineID
    Public Function GetLineID(StationID As String, IDApp As String) As String
        SQL = "use FAS
        select l.LineID  FROM [FAS].[dbo].[FAS_App_ListForPC] as  App
        left join FAS_Lines as L ON l.LineID = App.lineID
        where StationID = " & StationID & " and app_ID = " & IDApp
        Return SelectString(SQL)
    End Function


    'функция отображения / определения LineForPrint
    Public Function GetLineForPrint(LineID As String) As String
        SQL = "use FAS   SELECT [Print_Line] FROM [FAS].[dbo].[FAS_Lines] where LineID = " & LineID
        Return SelectString(SQL)
    End Function

    'функция отображения / определения IDApp 
    Public Function GetAppName(AppID As String) As String
        SQL = "use FAS SELECT App_Caption FROM [FAS].[dbo].[FAS_Applications] where  App_ID  = " & AppID
        Return SelectString(SQL)
    End Function
    'запрос списка линий 
    Public LineList As String = "Use FAS
        SELECT[LineID],[LineName],[Print_Line]  FROM [FAS].[dbo].[FAS_Lines]where [TipeID] = 3 and LineID != 14"




End Module
