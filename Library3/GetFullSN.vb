Imports System.Data.SqlClient

Public Module GetFullSN
    Private SQL As String
    'функция конвертирования строки в массив для расчета чексуммы FULLSN 
    Public D As Integer()
    Public Function StringToIntArray(raw As String) As Byte
        D = New Integer((raw.Length) - 1) {}
        Dim i As Integer
        For i = 0 To D.Length - 1
            D(i) = raw.Substring(i, 1)
        Next i
    End Function

    Public ProdDate, ProdTime, PrintCodeSN, PrintTextSN, FullSTBSN_Arr, STBSN As String
    Public Function GenerateFullSTBSN(SerNumber As String, DateText As String, DG_SNForPrint As DataGridView, LineForPrint As String, LOTcode As String, UserID As String) As String
        'выгружаем из таблицы серийных номеров грид содержащий дату на этикетке для выбранного SN
        SQL = "use FAS
        SELECT FORMAT(ManufDate,'d', 'de-de') as ManufDate, FORMAT(cast(ManufDate as time),N'hh\:mm\:ss') as ManufTime
        FROM [FAS].[dbo].[FAS_Start]  where SerialNumber = " & SerNumber
        LoadGridFromDB(DG_SNForPrint, SQL)
        'для формирования полногосерийного номера определяем необходимые данные:
        'ProdDate берем из FAS_SerialNumbers для выбранного серийника
        ProdDate = DG_SNForPrint.Item(0, 0).Value
        'формируем строку, которую будем преобразовываться в массив 
        FullSTBSN_Arr = "0" & Replace(ProdDate, ".", "") & LineForPrint & LOTcode & SerNumber
        StringToIntArray(FullSTBSN_Arr) ' преобразование строки в массив
        Dim result1, result2, r1, r2 As Integer
        result1 = (D(0) * 1 + D(1) * 2 + D(2) * 3 + D(3) * 4 + D(4) * 5 + D(5) * 6 + D(6) * 7 + D(7) * 8 + D(8) * 9 + D(9) * 10 +
               D(10) * 1 + D(11) * 2 + D(12) * 3 + D(13) * 4 + D(14) * 5 + D(15) * 6 + D(16) * 7 + D(17) * 8 + D(18) * 9 + D(19) * 10 +
               D(20) * 1 + D(21) * 2)
        result2 = (D(0) * 3 + D(1) * 4 + D(2) * 5 + D(3) * 6 + D(4) * 7 + D(5) * 8 + D(6) * 9 + D(7) * 10 + D(8) * 1 + D(9) * 2 +
               D(10) * 3 + D(11) * 4 + D(12) * 5 + D(13) * 6 + D(14) * 7 + D(15) * 8 + D(16) * 9 + D(17) * 10 + D(18) * 1 + D(19) * 2 +
               D(20) * 3 + D(21) * 4)
        r1 = result1 Mod 11
        r2 = result2 Mod 11

        Dim FullSTBSN As String
        If r1 = 10 Then
            If r2 = 10 Then
                FullSTBSN = "0" & FullSTBSN_Arr
            Else
                FullSTBSN = r2 & FullSTBSN_Arr
            End If
        Else
            FullSTBSN = r1 & FullSTBSN_Arr
        End If
        SQL = "use FAS        
            Update [FAS].[dbo].[FAS_Start] set [FullSTBSN] = '" & FullSTBSN & "', [AssemblyByID] = " & UserID & " Where [SerialNumber] = " & SerNumber
        RunCommand(SQL)
        Return FullSTBSN
    End Function
    '([SerialNumber],[PCBID],[LineID],[FullSTBSN], [ManufDate],[AssemblyDate],[AssemblyByID]) 
    '    values (" & SerNumber & "," & PCBID & "," & LineID & ",'" & FullSTBSN & "', '" & DateText & "',CURRENT_TIMESTAMP," & UserID & ")"
    'SQL = "use FAS        
    '    Update [FAS].[dbo].[FAS_Start] Set FullSTBSN = '" & FullSTBSN & "',AssemblyByID = " & UserID & " where SerialNumber = " & SerNumber &
    '    "Update [FAS].[dbo].[FAS_SerialNumbers] Set IsUsed = 1, InRepair = 0 where SerialNumber = " & SerNumber &
    '    "Delete [FAS].[dbo].[FAS_TempSerialNumbers] where SerialNumber = " & SerNumber


    'SQL = "use FAS  
    '    declare @SerialNumber int  
    '    select @SerialNumber = (SELECT TOP (1) [SerialNumber] FROM [FAS].[dbo].[FAS_TempSerialNumbers] where lotid = " & LOTID & " and isused = 0)  
    '    Update [FAS].[dbo].[FAS_TempSerialNumbers] Set isused = 1, PCB_ID = " & PCBID & ", LineID = " & LineID & ", ManufDate = " & DateText & ", PrintStationID = " & StationID & "  where SerialNumber = @SerialNumber  
    '    WAITFOR delay '00:00:00:100'
    '    select SerialNumber from [FAS].[dbo].[FAS_Start] where SerialNumber = @SerialNumber and PrintStationID = " & StationID

    Public Function GenerateFullSTBSN_SDTV(SN As Integer, LOTcode As String, LOTID As Integer) As String
        'для формирования полногосерийного номера определяем необходимые данные:
        'ProdDate берем из FAS_SerialNumbers для выбранного серийника
        'формируем строку, которую будем преобразовываться в массив 
        '00101202001 - постоянная часть китайского номера
        FullSTBSN_Arr = "00101202001" & LOTcode & SN
        StringToIntArray(FullSTBSN_Arr) ' преобразование строки в массив
        Dim result1, result2, r1, r2 As Integer
        result1 = (D(0) * 1 + D(1) * 2 + D(2) * 3 + D(3) * 4 + D(4) * 5 + D(5) * 6 + D(6) * 7 + D(7) * 8 + D(8) * 9 + D(9) * 10 +
                   D(10) * 1 + D(11) * 2 + D(12) * 3 + D(13) * 4 + D(14) * 5 + D(15) * 6 + D(16) * 7 + D(17) * 8 + D(18) * 9 + D(19) * 10 +
                   D(20) * 1 + D(21) * 2)
        result2 = (D(0) * 3 + D(1) * 4 + D(2) * 5 + D(3) * 6 + D(4) * 7 + D(5) * 8 + D(6) * 9 + D(7) * 10 + D(8) * 1 + D(9) * 2 +
                   D(10) * 3 + D(11) * 4 + D(12) * 5 + D(13) * 6 + D(14) * 7 + D(15) * 8 + D(16) * 9 + D(17) * 10 + D(18) * 1 + D(19) * 2 +
                   D(20) * 3 + D(21) * 4)
        r1 = result1 Mod 11
        r2 = result2 Mod 11

        Dim FullSTBSN As String
        If r1 = 10 Then
            If r2 = 10 Then
                FullSTBSN = "0" & FullSTBSN_Arr
            Else
                FullSTBSN = r2 & FullSTBSN_Arr
            End If
        Else
            FullSTBSN = r1 & FullSTBSN_Arr
        End If
        RunCommand("Update [FAS].[dbo].[SDTV_Upload] set FullSN = '" & FullSTBSN & "', MAC = '" & GenMAC(SN) & "', IsLocked = 0 where SN = " & SN)
        Return FullSTBSN
    End Function


End Module
