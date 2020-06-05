Public Module CheckSNFormate
    'функция определения длины серийного номера
    Public Function GetLenSN(Format As String) As Integer
        Dim Coordinats As Integer() = New Integer(2) {}
        For i = 0 To 5 Step 2
            Dim J As Integer
            Coordinats(J) = Convert.ToInt32(Mid(Mid(Format, Len(Format) - 5), i + 1, 2), 16)
            J += 1
        Next
        Return (Coordinats(0) + Coordinats(1) + Coordinats(2))
    End Function
    'функция определения координат серийного номера
    Public Function GetCoordinats(Format As String) As Array
        Dim Coordinats As Integer() = New Integer(2) {}
        For i = 0 To 5 Step 2
            Dim J As Integer
            Coordinats(J) = Convert.ToInt32(Mid(Mid(Format, Len(Format) - 5), i + 1, 2), 16)
            J += 1
        Next
        Return Coordinats
    End Function
    'функция определения формата серийного номера
    Public Function GetSNFormat(FormatSMT As String, FormatFAS As String, SN As String)
        Dim Coordinats As Integer() = New Integer(2) {}
        Dim Res As Integer, Bool As Boolean
        Coordinats = GetCoordinats(FormatSMT)
        ' i = 1 --Номер SMT, i = 2 --Номер FAS, i = 3 --Номер не определен
        For i = 1 To 3
            If i <> 3 Then
                Dim SNBase As String
                SNBase = If(i = 1, FormatSMT, FormatFAS)
                Dim MascBase As String = Mid(SNBase, 1, Coordinats(0)) + Mid(SNBase, Coordinats(0) + Coordinats(1) + 1, Coordinats(2))
                Dim MascSN As String = Mid(SN, 1, Coordinats(0)) + Mid(SN, Coordinats(0) + Coordinats(1) + 1, Coordinats(2))
                Bool = If(MascBase = MascSN, True, False)
                If Bool = True Then
                    Res = i
                    Exit For
                End If
            Else
                Res = i
            End If
        Next
        Return Res
    End Function



End Module
