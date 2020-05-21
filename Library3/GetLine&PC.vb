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
    'Функция сбора информации о рабочей станции
    Public Function GetPCInfo(AppID As Integer) As ArrayList
        Dim StationId As Integer = GetStationID(GetPCName())
        'Dim StationId As Integer = GetStationID("WSG121211")
        Dim objlist As New ArrayList(SelectListString("USE FAS
          SELECT List.App_ID, Ap.App_Caption,List.lineID, L.LineName,St.StationName,[CT_ScanStep]
          FROM [FAS].[dbo].[FAS_App_ListForPC] as List
          left join [FAS].[dbo].[FAS_Applications] as Ap On Ap.App_ID = List.App_ID
          left join [FAS].[dbo].[FAS_Stations] as St On St.StationID = List.StationID
          left join FAS_Lines as L ON l.LineID = List.lineID
          where st.StationID = '" & StationId & "' and List.App_ID = " & AppID))
        Return objlist
    End Function

    'функция поиска ID станции
    Public Function GetStationID(StationName As String) As Integer
        Dim StationId As Integer = SelectInt("use FAS SELECT StationID  FROM  [FAS].[dbo].[FAS_Stations] where StationName = '" & StationName & "'")
        If StationId = Nothing Then
            StationId = SelectInt("use FAS
        insert into [FAS].[dbo].[FAS_Stations] (StationName, CreateDate) values ('" & StationName & "', CURRENT_TIMESTAMP) 
        SELECT StationID  FROM  [FAS].[dbo].[FAS_Stations] where StationName = '" & StationName & "'")
        End If
        Return StationId
    End Function

    'функция отображения / определения LineForPrint
    Public Function GetLineForPrint(LineID As String) As String
        SQL = "use FAS   SELECT [Print_Line] FROM [FAS].[dbo].[FAS_Lines] where LineID = " & LineID
        Return SelectString(SQL)
    End Function

    'функция отображения / определения IDApp 
    'запрос списка линий 
    Public LineList As String = "Use FAS
        SELECT[LineID],[LineName],[Print_Line]  FROM [FAS].[dbo].[FAS_Lines]where [TipeID] = 3 and LineID != 14"






End Module
