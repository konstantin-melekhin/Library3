Public Module LogInUser
    Public UserID, UserName, UserGroup As String
    Public UserData(2) As String
    Public Function GetUserData(RFID As String, GB_UserData As GroupBox, GB_WorkAria As GroupBox, L_UserName As Label, TB_RFIDIn As TextBox) As Array
        UserID = SelectString("USE FAS SELECT [UserID] FROM [FAS].[dbo].[FAS_Users] where [RFID] = '" & RFID & "' and IsActiv = 1")
        UserName = SelectString("USE FAS SELECT [UserName] FROM [FAS].[dbo].[FAS_Users] where [RFID] = '" & RFID & "' and IsActiv = 1")
        UserGroup = SelectString("USE FAS SELECT [UsersGroupID] FROM [FAS].[dbo].[FAS_Users] where [RFID] = '" & RFID & "' and IsActiv = 1")
        If UserID = "" Then
            MsgBox("Пользователь не зарегистрирован")
            TB_RFIDIn.Text = ""
            TB_RFIDIn.Focus()
        Else
            L_UserName.Text = UserName
            GB_UserData.Visible = False
            GB_WorkAria.Visible = True
            UserData = {UserID, UserGroup}
        End If
        Return UserData
    End Function
End Module
