

Public Module LineSelector
    Public Liter, LiterID As String
    Public Function LiterSelect(LineNumber As String)
        'присваиваем литеру выбранной линии
        'FAS
        If LineNumber = "FAS Line 1" Then
            Liter = "A"
        ElseIf LineNumber = "FAS Line 2" Then
            Liter = "B"
        ElseIf LineNumber = "FAS Line 3" Then
            Liter = "C"
        ElseIf LineNumber = "FAS Line 4" Then
            Liter = "D"
        ElseIf LineNumber = "FAS Line 5" Then
            Liter = "E"
        ElseIf LineNumber = "FAS Line 6" Then
            Liter = "F"
            'THT
        ElseIf LineNumber = "THT Line 1" Then
            Liter = "G"
        ElseIf LineNumber = "THT Line 2" Then
            Liter = "H"
        ElseIf LineNumber = "THT Line 3" Then
            Liter = "K"
        ElseIf LineNumber = "THT Line 4" Then
            Liter = "M"
        ElseIf LineNumber = "THT Line 5" Then
            Liter = "N"
        ElseIf LineNumber = "THT Line 6" Then
            Liter = "P"
        End If
        Return Liter
    End Function

    Public Function LiterIDSelect(LineNumber As String)
        'присваиваем литеру выбранной линии
        'FAS
        If LineNumber = "FAS Line 1" Then
            Liter = 1
        ElseIf LineNumber = "FAS Line 2" Then
            Liter = 2
        ElseIf LineNumber = "FAS Line 3" Then
            Liter = 3
        ElseIf LineNumber = "FAS Line 4" Then
            Liter = 4
        ElseIf LineNumber = "FAS Line 5" Then
            Liter = 5
        ElseIf LineNumber = "FAS Line 6" Then
            Liter = 6
            'THT
        ElseIf LineNumber = "THT Line 1" Then
            Liter = 7
        ElseIf LineNumber = "THT Line 2" Then
            Liter = 8
        ElseIf LineNumber = "THT Line 3" Then
            Liter = 9
        ElseIf LineNumber = "THT Line 4" Then
            Liter = 10
        ElseIf LineNumber = "THT Line 5" Then
            Liter = 11
        ElseIf LineNumber = "THT Line 6" Then
            Liter = 12
        End If
        Return LiterID
    End Function
End Module
