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

    Public Function GetSNFormat(FormatSMT As String, FormatFAS As String, SN As String, HexSN As Boolean) As ArrayList
        Dim Coordinats() As Integer
        Dim Res As ArrayList = New ArrayList()
        Dim VarSN As Integer
        ' i = 1 --Номер SMT, i = 2 --Номер FAS, i = 3 --Номер не определен
        For i = 1 To 3
            If i <> 3 Then
                Dim SNBase As String
                Coordinats = GetCoordinats(If(i = 1, FormatSMT, FormatFAS))
                SNBase = If(i = 1, FormatSMT, FormatFAS)
                Dim MascBase As String = Mid(SNBase, 1, Coordinats(0)) + Mid(SNBase, Coordinats(0) + Coordinats(1) + 1, Coordinats(2))
                Dim MascSN As String = Mid(SN, 1, Coordinats(0)) + Mid(SN, Coordinats(0) + Coordinats(1) + 1, Coordinats(2))
                If (MascBase = MascSN) = True Then
                    Res.Add(True) 'Res(0)
                    Res.Add(i) 'Res(1)
                    If i = 1 Or (i = 2 And HexSN = False) Then
                        VarSN = Convert.ToInt32(Mid(SN, Coordinats(0) + 1, Coordinats(1)))
                    ElseIf i = 2 And HexSN = True Then
                        VarSN = CInt("&H" & Mid(SN, Coordinats(0) + 1, Coordinats(1)))
                    End If
                    Res.Add(VarSN) 'Res(2)
                    Exit For
                End If
            Else
                Res.Add(False) 'Res(0)
                Res.Add(i) 'Res(1)
                Res.Add(0) 'Res(2)
            End If
        Next

        Select Case Res(1)
            Case 1
                Res.Add("Формат номера " & SN & vbCrLf & "соответствует SMT!") 'Res(3) ' Текст сообщения
            Case 2
                Res.Add("Формат номера " & SN & vbCrLf & "соответствует FAS!")
            Case 3
                Res.Add("Формат номера " & SN & vbCrLf & "не соответствует выбранному лоту!")
        End Select
        Return Res
    End Function



End Module
