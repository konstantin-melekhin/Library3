Public Module LogInUser
    Public Function GetUserData(RFID As String, GBUserData As GroupBox, GBWorkAria As GroupBox, LUserName As Label, TBRFIDIn As TextBox) As ArrayList
        Dim UserInfo As New ArrayList(
        SelectListString("USE FAS SELECT [UserID],[UserName],[UsersGroupID] 
                                    FROM [FAS].[dbo].[FAS_Users] where [RFID] = '" & RFID & "' and IsActiv = 1"))
        If UserInfo.Count = 0 Then
            MsgBox("Пользователь не зарегистрирован")
            TBRFIDIn.Text = ""
            TBRFIDIn.Focus()
        Else
            LUserName.Text = UserInfo(1)
            GBUserData.Visible = False
            GBWorkAria.Visible = True
        End If
        Return UserInfo
    End Function
End Module
