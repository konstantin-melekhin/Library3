Public Module LabelPrint
    'Функция вывода информации в Лэйбл

    Public Sub PrintLabel(Lab As Label, Message As String, Col As Color)
        Lab.Text = Message
        Lab.Location = New Point(13, 349)
        Lab.Font = New Font("Microsoft Sans Serif", 20, FontStyle.Bold)
        Lab.ForeColor = Col
        Lab.Visible = True
    End Sub
    Public Sub PrintLabel(Lab As Label, Message As String, x As Integer, y As Integer, Col As Color)
        Lab.Text = Message
        Lab.Location = New Point(x, y)
        Lab.Font = New Font("Microsoft Sans Serif", 20, FontStyle.Bold)
        Lab.ForeColor = Col
        Lab.Visible = True
    End Sub
    Public Sub PrintLabel(Lab As Label, Message As String, x As Integer, y As Integer, Col As Color, Font As Font)
        Lab.Text = Message
        Lab.Location = New Point(x, y)
        Lab.Font = Font 'Font("Microsoft Sans Serif", 20, FontStyle.Bold)
        Lab.ForeColor = Col
        Lab.Visible = True
    End Sub
End Module
