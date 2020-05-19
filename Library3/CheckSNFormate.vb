Public Module CheckSNFormate
    'запрос списка линий 
    Public SMTLineList As String = "Use FAS
       SELECT ID,Name,Print_Line  FROM [FAS].[dbo].[M_LineID] where TypeID = 1 and id !=1"
    'запрос обновления списка лотов для SMT Formats
    Public SMTLotList As String = "Use FAS
        SELECT [LOTCode]
        ,[FullLOTCode]
        ,M.ModelName
        ,[Spec]
        ,[OneSided]
        ,[TOPNumberFormate]
        ,[BOTNumberFormate]
        ,[CheckTopRange]
        ,[StartTOPRange]
        ,[EndTOPRange]
        ,[CheckBOTRange]
        ,[StartBOTRange]
        ,[EndBOTRange]
        ,[CreateDate]
        ,[CreateBy]
        ,[isActiv]
        FROM [FAS].[dbo].[SMT_NumbersFormat] as f
        left join M_Models as M On M.ModelID = F.ModelID
        where isActiv = 1
        order by createdate desc"

    'функция определения маски серийного номера
    Public Masc As String
    Public LineLen, LineStart, ModelLen, ModelStart, Masc1Len, Masc1Start, Masc2Len, Masc2Start, SNLen, SNStart, FullSNLen As Integer
    Dim Literal As String
    Public Function GetSNFormat(Format As String) As String
        Masc = ""
        Dim tempMasc As String
        FullSNLen = Len(Format)
        For i = 1 To FullSNLen
            tempMasc = ""
            Literal = Mid(Format, i, 2)
            Select Case Literal
                Case "L$" To "L$"
                    LineLen = 1
                    LineStart = i + 3
                    'tempMasc = " " + Mid(Format, LineStart, LineLen)
                Case "A$" To "A$"
                    Masc1Len = Mid(Format, (i + 2), 2)
                    Masc1Start = i + 4
                    tempMasc = Mid(Format, Masc1Start, Masc1Len)
                Case "B$" To "B$"
                    SNLen = Mid(Format, (i + 2), 2)
                    SNStart = i + 4
                Case "C$" To "C$"
                    Masc2Len = Mid(Format, (i + 2), 2)
                    Masc2Start = i + 4
                    tempMasc = Mid(Format, Masc2Start, Masc2Len)
                Case "M$" To "M$"
                    ModelLen = Mid(Format, (i + 2), 2)
                    ModelStart = i + 4
                    tempMasc = Mid(Format, ModelStart, ModelLen)
            End Select
            Masc = Masc + tempMasc
        Next
        Return Masc
    End Function

    'Проверка соответствия серийного номера SMT
    'A - Kоординаты Masc1
    'A1 - Длина Masc1
    'C - Kоординаты Masc2
    'C1 - Длина Masc2
    Public Function CheckSMTSNFormatAndRange(ScanSN As String, Format As String, SNMin As String, SNMax As String, A As Integer, A1 As Integer,
                            B As Integer, B1 As Integer, C As Integer, C1 As Integer, CheckRange As Boolean) As Integer
        Dim Result As Integer = 0
        If Mid(ScanSN, A, A1) + Mid(ScanSN, C, C1) = Format Then
            If CheckRange = True Then
                Dim SN As Integer = Mid(ScanSN, B, B1)
                If SN >= SNMin And SN <= SNMax Then
                    Result = 1 'Проверка пройдена
                Else
                    Result = 2 'Серийный номер вне указанного в ЛОТе диапазона!
                End If
            Else
                Result = 1
            End If
        ElseIf Len(ScanSN) = A1 + B1 + C1 Then
            Result = 3 'Маска не соответствует указанной в лоте
        End If
        Return Result
    End Function


    'Проверка соответствия серийного номера FAS
    Public Function CheckFASSNFormatAndRange(ScanSN As String, Format As String, SNMin As String, SNMax As String, A As Integer, A1 As Integer,
                            B As Integer, B1 As Integer, C As Integer, C1 As Integer, CheckRange As Boolean) As Integer
        Dim Result As Integer = 0
        If Mid(ScanSN, A, A1) + Mid(ScanSN, C, C1) = Format Then
            If CheckRange = True Then
                Dim SN As Integer = Mid(ScanSN, B, B1)
                If SN >= SNMin And SN <= SNMax Then
                    Result = 1 'Проверка пройдена
                Else
                    Result = 2 'Серийный номер вне указанного в ЛОТе диапазона!
                End If
            Else
                Result = 1
            End If
        ElseIf Len(ScanSN) = A1 + B1 + C1 Then
            Result = 3 'Маска не соответствует указанной в лоте
        End If
        Return Result
    End Function


End Module
