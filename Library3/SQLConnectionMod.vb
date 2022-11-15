Imports System.Data.SqlClient



Public Module SQLConnectionMOD
    Public conn As SqlConnection
    Public Function GetConnect() As Boolean
        Try
            conn = New SqlConnection("Data Source= 192.168.180.9\flat; Initial Catalog= ; User Id = melekhin; Password = me1ekhin;")
            'conn = New SqlConnection("Data Source= wsg180123\sqlmel; Initial Catalog= ;Integrated Security=True;")
            '          <!--<connectionStrings>
            '  <add name = "HomeBase" connectionString="Data Source=WIN-3R3T8Q67MUE\SQLEXPRESS; Initial Catalog=CYKP; Integrated Security=True;" providerName ="System.Data.SqlClient"/>
            '</connectionStrings>-->

            conn.Open()
            Return 1
        Catch ex As Exception
            MsgBox("Ошибка подключения к SQL сервер. " & ex.Message)
            Return 0
        End Try
    End Function

    Public Sub LoadGridFromDB(ByVal Grid1 As DataGridView, cmd As String)
        GetConnect()
        Try
            Dim c As New SqlCommand
            Dim da As New SqlDataAdapter
            Dim ds As New DataSet

            c = conn.CreateCommand
            c.CommandText = cmd

            da.SelectCommand = c
            da.Fill(ds, "Table1")

            Grid1.DataSource = ds
            Grid1.DataMember = "Table1"
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Sub

    Public Sub RunCommand(cmd As String)
        GetConnect()
        Try
            Dim c As New SqlCommand
        c = conn.CreateCommand
        c.Connection = conn
        c.CommandType = CommandType.Text
        c.CommandText = cmd
        c.ExecuteNonQuery()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Sub

    Public Function SelectString(cmd As String) As String
        GetConnect()
        Try
            Dim c As New SqlCommand
            Dim r As SqlDataReader
            Dim k As String
            k = ""
            c = conn.CreateCommand
            c.CommandType = CommandType.Text
            c.CommandText = cmd

            r = c.ExecuteReader
            If r.Read Then
                k = r.Item(0)
                r.Close()
            End If

            Return k
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Function

    Public Function SelectFloat(cmd As String) As Double
        GetConnect()
        Try
            Dim c As New SqlCommand
            Dim r As SqlDataReader
            Dim k As Double

            c = conn.CreateCommand
            c.CommandType = CommandType.Text
#Disable Warning CA2100 ' Проверка запросов SQL на уязвимости безопасности
            c.CommandText = cmd
#Enable Warning CA2100 ' Проверка запросов SQL на уязвимости безопасности

            r = c.ExecuteReader
            If r.Read Then
                k = r.Item(0)
                r.Close()
            End If

            Return k
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Function

    Public Function SelectInt(cmd As String) As Integer
        GetConnect()
        Try
            Dim c As New SqlCommand
            Dim r As SqlDataReader
            Dim k As Integer
            c = conn.CreateCommand
            c.CommandType = CommandType.Text
#Disable Warning CA2100 ' Проверка запросов SQL на уязвимости безопасности
            c.CommandText = cmd
#Enable Warning CA2100 ' Проверка запросов SQL на уязвимости безопасности

            r = c.ExecuteReader
            If r.Read Then
                k = r.Item(0)
                r.Close()
            End If

            Return k
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Function

    Public Function SelectBoolean(cmd As String) As Boolean
        GetConnect()
        Try
            Dim c As New SqlCommand
            Dim r As SqlDataReader
            Dim k As Boolean

            c = conn.CreateCommand
            c.CommandType = CommandType.Text
            c.CommandText = cmd

            r = c.ExecuteReader
            If r.Read Then
                k = r.Item(0)
                r.Close()
            End If

            Return k
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Function

    Public Function SelectByte(cmd As String) As Byte()
        GetConnect()
        Try
            Dim c As New SqlCommand
            Dim r As SqlDataReader
            Dim k() As Byte

            c = conn.CreateCommand
            c.CommandType = CommandType.Text
            c.CommandText = cmd

            r = c.ExecuteReader
            If r.Read Then
                k = r.Item(0)
                r.Close()
            End If

            Return k
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Function

    Public Function SelectListString(cmd As String) As ArrayList
        GetConnect()
        Try
            Dim c As New SqlCommand
            Dim r As SqlDataReader
            Dim k As New ArrayList()
            'k = ""
            c = conn.CreateCommand
            c.CommandType = CommandType.Text
            c.CommandText = cmd

            r = c.ExecuteReader
            If r.Read Then
                For i = 0 To r.VisibleFieldCount - 1
                    k.Add(r.Item(i))
                Next
            End If
            r.Close()

            Return k
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Function

    Public Function LoadGridFromDB2(ByVal Grid1 As DataGridView, cmd As String, ds As DataSet) As DataSet
        GetConnect()
        Try
            Dim c As New SqlCommand
            Dim da As New SqlDataAdapter
            'Dim ds As New DataSet

            c = conn.CreateCommand
            c.CommandText = cmd

            da.SelectCommand = c
            da.Fill(ds, "Table1")

            Grid1.DataSource = ds
            Grid1.DataMember = "Table1"

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        conn.Close()
        Return ds
    End Function

    Public Function LoadDS(cmd As String) As DataSet
        GetConnect()
        Dim ds As New DataSet
        Try
            Dim c As New SqlCommand
            Dim da As New SqlDataAdapter
            c = conn.CreateCommand
            c.CommandText = cmd

            da.SelectCommand = c
            da.Fill(ds, "Table1")

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        conn.Close()
        Return ds
    End Function

    Public Sub LoadCombo(ByVal CB As ComboBox, cmd As String)
        GetConnect()
        Dim c As New SqlCommand
        Dim da As New SqlDataAdapter
        Dim ds As New DataTable

        c = conn.CreateCommand
        c.CommandText = cmd

        da.SelectCommand = c
        da.Fill(ds)

        CB.DataSource = ds
        CB.DisplayMember = ds.Columns(0).ToString
        conn.Close()
    End Sub

End Module
