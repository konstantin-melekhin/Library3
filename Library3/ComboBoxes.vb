Public Module ComboBoxes
    'функция добавления значений в комбобокс
    Function addToList(CB As ComboBox, DG As DataGridView, column As Integer)
        If DG.Rows.Count <> 0 Then
            For J = 0 To DG.Rows.Count - 1
                CB.Items.Add(DG.Rows(J).Cells(column).Value)
            Next
        Else
            MsgBox(DG.Name & " не содержит значений!")
        End If
    End Function

    'функция поиска значений в гриде по выбранному значению в комбобоксе
    Function SearchInList(CB As ComboBox, DG As DataGridView, column1 As Integer, column2 As Integer)
        Dim ID As Integer = 0
#Disable Warning CA1062 ' Проверить аргументы или открытые методы
        For J = 0 To DG.Rows.Count - 1
#Enable Warning CA1062 ' Проверить аргументы или открытые методы
            If DG.Rows(J).Cells(column1).Value = CB.Text Then
                ID = DG.Rows(J).Cells(column2).Value
                Exit For
            End If
        Next
        Return ID
    End Function

    'функция поиска значений в гриде при работе цикла. i - итерация цикла
    Function GetColumnValue(DG As DataGridView, i As Integer, column As Integer)
        Dim Value As String
        Value = DG.Rows(i - 1).Cells(column).Value
        Return Value
    End Function

End Module
