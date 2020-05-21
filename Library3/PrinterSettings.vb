Imports System.IO
Imports System.Text

Public Module PrinterSettings
    Public PrSet As Boolean
    Public PrintSetpath As String
    Public fs As FileStream
    Public info As Byte()

    Public Function GetPrinterSettings(LAbelScenarioID As String, TB_IDCode_X As TextBox, TB_IDCode_Y As TextBox, TB_IDText_X As TextBox, TB_IDText_Y As TextBox, TB_IDNum_X As TextBox)
        Dim IDCode_X As String, IDCode_Y As String, IDText_X As String, IDText_Y As String, IDNum_X As String
        If LAbelScenarioID = 1 Then
            PrintSetpath = "C:\PrinterSettings\PrinterSettingsA.txt"
        ElseIf LAbelScenarioID = 2 Then
            PrintSetpath = "C:\PrinterSettings\PrinterSettingsB.txt"
        End If
        Try
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(PrintSetpath)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                While Not MyReader.EndOfData
                    currentRow = MyReader.ReadFields()
                    Dim currentField As String
                    Dim i As Integer = 1
                    For Each currentField In currentRow
                        Select Case i
                            Case 1
                                TB_IDCode_X.Text = currentField
                            Case 2
                                TB_IDCode_Y.Text = currentField
                            Case 3
                                TB_IDText_X.Text = currentField
                            Case 4
                                TB_IDText_Y.Text = currentField
                            Case 5
                                TB_IDNum_X.Text = currentField
                        End Select
                        i += 1
                    Next
                End While
                PrSet = True
            End Using
        Catch ex As Exception
            PrSet = False
        End Try
        Return PrSet
    End Function


    Public Sub CreatePrinterSettings(LAbelScenarioID As String)
        My.Computer.FileSystem.CreateDirectory("C:\PrinterSettings")
        ' Create or overwrite the file.
        fs = File.Create(PrintSetpath)
        ' Add text to the file.
        If LAbelScenarioID = 1 Then
            info = New UTF8Encoding(True).GetBytes("60,220,40,320,144")
        ElseIf LAbelScenarioID = 2 Then
            info = New UTF8Encoding(True).GetBytes("36,133,82,185,166")
        End If
        fs.Write(info, 0, info.Length)
        fs.Close()
    End Sub

End Module
