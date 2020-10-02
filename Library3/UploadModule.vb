Imports System
Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.IO.Ports

Public Module UploadModule
    Private SQL As String

    Public crcModel As New RocksoftCrcModel(32, &H4C11DB7, 4294967295, False, False, 0)
    'функция конвертирования текста в HEX
    Public Function StrToHex(Data As String) As String
        Dim sVal As String
        Dim sHex As String = ""
        While Data.Length > 0
            sVal = Conversion.Hex(Strings.Asc(Data.Substring(0, 1).ToString()))
            Data = Data.Substring(1, Data.Length - 1)
            sHex = sHex & sVal
        End While
        Return sHex
    End Function

    '-------------------------------------------------------------------------------------------------------------------------------
    'генерируем посылку в приемник, если команда из одного байта (SCID, CASID, LDS, SWver, SWGS1ver)
    Public Function DataGenerationOneByte(ByteRequest As String) As String
        Dim GeneratedRequest As String = ""
        Dim Header = "0DF20101000C" '\ заголовок с чексуммой, 
        '\ преобразовываем полученную Data в массив данных и считаем DataCS  
        Dim ByteRequestHex As Byte() = New Byte((ByteRequest.Length / 2) - 1) {}
        For i = 0 To ByteRequestHex.Length - 1
            ByteRequestHex(i) = Convert.ToByte(ByteRequest.Substring((i * 2), 2), &H10)
        Next i
        Dim ByteRequestCS = Hex(crcModel.ComputeCrc(ByteRequestHex)) 'считаем КОНТРОЛЬНУЮ СУММУ ДЛЯ ByteRequest
        Dim L As Integer = ByteRequestCS.Length - 1
        ByteRequestCS = ByteRequestCS.Chars(L - 1) & ByteRequestCS.Chars(L) & ByteRequestCS.Chars(L - 3) & ByteRequestCS.Chars(L - 2) '\ обрезаем лишние байты и записываем в обратном порядке байтов
        GeneratedRequest = Header + ByteRequest + ByteRequestCS ' формируем запрос для посылки в приемник
        Return GeneratedRequest
    End Function
    ''------------------------------------------------------------------------------------------------------------
    'генерируем посылку в приемник для SN и MAC
    Public Function DataGenerationSNorMAC(ByteRequest As String) As String
        Dim GeneratedRequest As String
        Dim Header As String = "0DF201180091" '\ заголовок с чексуммой, 
        '\ преобразовываем полученную Data в массив данных и считаем DataCS  
        Dim ByteRequestHex As Byte() = New Byte((ByteRequest.Length / 2) - 1) {}
        For i = 0 To ByteRequestHex.Length - 1
            ByteRequestHex(i) = Convert.ToByte(ByteRequest.Substring((i * 2), 2), &H10)
        Next i
        Dim ByteRequestCS = Hex(crcModel.ComputeCrc(ByteRequestHex)) 'считаем КОНТРОЛЬНУЮ СУММУ ДЛЯ ByteRequest
        Dim L As Integer = ByteRequestCS.Length - 1
        ByteRequestCS = ByteRequestCS.Chars(L - 1) & ByteRequestCS.Chars(L) & ByteRequestCS.Chars(L - 3) & ByteRequestCS.Chars(L - 2) '\ обрезаем лишние байты и записываем в обратном порядке байтов
        GeneratedRequest = Header + ByteRequest + ByteRequestCS ' формируем запрос для посылки в приемник
        Return GeneratedRequest
    End Function

    ''------------------------------------------------------------------------------------------------------------
    'генерируем посылку в приемник HDCP и Cert 
    Public Function DataGenerationOther(Data As String, DataLenght As String) As String
        Dim Request = ""
        Dim Header = "0DF201" & DataLenght '\ собрали заголовок
        '\ преобразовываем заголовок в массив данных и считаем HeaderCS       
        Dim HeaderHex As Byte() = New Byte((Header.Length / 2) - 1) {}
        For i = 0 To HeaderHex.Length - 1
            HeaderHex(i) = Convert.ToByte(Header.Substring((i * 2), 2), &H10)
        Next i
        Dim HeaderCS = Hex(crcModel.ComputeCrc(HeaderHex)) 'считаем HeaderCS   
        HeaderCS = HeaderCS.Chars(6) & HeaderCS.Chars(7) '\ обрезаем лишние байты

        '\ преобразовываем полученную Data в массив данных и считаем DataCS  
        Dim DataHex As Byte() = New Byte((Data.Length / 2) - 1) {}
        For i = 0 To DataHex.Length - 1
            DataHex(i) = Convert.ToByte(Data.Substring((i * 2), 2), &H10)
        Next i
        Dim ByteRequestCS = Hex(crcModel.ComputeCrc(DataHex)) 'считаем КОНТРОЛЬНУЮ СУММУ ДЛЯ ByteRequest
        Dim L As Integer = ByteRequestCS.Length - 1
        ByteRequestCS = ByteRequestCS.Chars(L - 1) & ByteRequestCS.Chars(L) & ByteRequestCS.Chars(L - 3) & ByteRequestCS.Chars(L - 2)  '\ обрезаем лишние байты и записываем в обратном порядке байтов
        Request = Header + HeaderCS + Data + ByteRequestCS ' формируем запрос для посылки в приемник
        Return Request
    End Function
    '-------------------------------------------------------------------------------------------------------------------------------
    'функция чтения файла в массив и преобразование в строку формата HEX 
    Public Function FileToArray(filePath As String) As String
        Dim FileByteArray() As Byte
        FileByteArray = File.ReadAllBytes(filePath)
        Dim FileContent As String = ""
        FileContent += System.BitConverter.ToString(FileByteArray, 0, FileByteArray.Length)
        FileContent = Replace(FileContent, "-", "")
        Return FileContent
    End Function
    ''------------------------------------------------------------------------------------------------------------
    'Функция заполнения массива данных в датагрид
    Public Function addRowToGrid(MyArray As String(), Grid As DataGridView)
        'если длинна массива больше чем количество элементов массива то добавляем столбцы
        While (MyArray.Length > Grid.ColumnCount)
            'если колонок нехватает добавляем их пока их будет хватать
            Grid.Columns.Add("", "")
            'пока столбцов не станет нужное количество
        End While
        'создаем новую запись, вносим целиком массив
        Grid.Rows.Add(MyArray)
    End Function
    ''------------------------------------------------------------------------------------------------------------
    'Функция чтения файлов из указанной дериктории в datagrid
    Public Function ListFiles(folderPath As String, MyGrid As DataGridView)
        MyGrid.Rows.Clear()
        Dim fileNames As String() = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
        ' записывает названия файлов в листбокс
        For Each fileName As String In fileNames
            Dim MyArr As String() = {Path.GetFileName(fileName), fileName}
            addRowToGrid(MyArr, MyGrid)
        Next
    End Function

    ''------------------------------------------------------------------------------------------------------------
    'Функция копирования файлов из указанной дериктории в нужную
    Public Function Copy_File(sFileName As String, sNewFileName As String)
        'sFileName = "C:\1\2.txt"    'имя файла для копирования
        'sNewFileName = "D:\1.txt"    'имя копируемого файла. Директория(в данном случае диск D) должна существовать
        If Dir(sFileName, 16) = "" Then
            MsgBox("Нет такого файла", vbCritical, "Ошибка")
            Exit Function
        End If
        FileCopy(sFileName, sNewFileName) 'копируем файл
        'MsgBox("Файл скопирован", vbInformation, "www.excel-vba.ru")
    End Function




    'не позволяет писать буквы в текстбокс
    'Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged 'Устанавливается правило, при котором в TextBox нельзя писать буквы, только цифры от 0 до 9
    '    If TextBox3.Text Like "*[!0-9]*" Then _
    ' TextBox3.Text = TextBox3.Tag _
    ' Else : TextBox3.Tag = TextBox3.Text
    'End Sub

    '________________________________________________________________________________________________
    'Новые функции для Upload Station
    ''------------------------------------------------------------------------------------------------------------
    'сформировать мак адрес из серийного номера
    Public Function GenMAC(SNshort As String)
        Dim MACHex As String
        MACHex = Hex(Mid(SNshort, 2, 7)) '- используется 7 знаков справа
        If Len(MACHex) < 6 Then ' длина для МАС адреса должна быть строго 6 знаков
            For j = 1 To 5
                MACHex = "0" & MACHex
                If Len(MACHex) = 6 Then
                    Exit For
                End If
            Next
        End If
        Return "00:21:52:" & Mid(MACHex, 1, 2) & ":" & Mid(MACHex, 3, 2) & ":" & Mid(MACHex, 5, 2)
    End Function


    ''------------------------------------------------------------------------------------------------------------
    'подготовка данных для печати серийника
    'Загрузка данных из БД для прошивки приемника
    Public UploadData As Array
    Public Function LoadSNData(SNshort As String, DG_SNTable As DataGridView) As Array
        Dim PrintCodeSN, PrintTextSN, MACTable, HDCPName, CertName, LabelDate, LabelTime, FullSTBSN As String 'переменные для вывода значений из таблицы FAS_SerialNumbers
        'выгружаем HDCP, CERT, дату производства и время (печать этикетки)  для введенного номера из грида  проверки СН
        SQL = "use fas
                SELECT h.HDCPName, c.CERTName, [FullSTBSN],Format([ManufDate],'dd.MM.yyyy', 'de-de'),
                Format([ManufDate],'HH:mm:ss', 'de-de' )
                FROM [FAS].[dbo].[FAS_Start] as FSt
                left join FAS_HDCP as H on H.SerialNumber = FSt.SerialNumber
                left join FAS_CERT as C on C.SerialNumber = FSt.SerialNumber
                where FSt.SerialNumber = " & SNshort
        LoadGridFromDB(DG_SNTable, SQL)

        If DG_SNTable.Rows.Count <> 0 Then
            MACTable = GenMAC(SNshort) 'UPD(1)
            HDCPName = DG_SNTable.Rows(0).Cells(0).Value 'UPD(2)
            CertName = DG_SNTable.Rows(0).Cells(1).Value 'UPD(3)
            FullSTBSN = DG_SNTable.Rows(0).Cells(2).Value 'UPD(4)
            LabelDate = DG_SNTable.Rows(0).Cells(3).Value 'UPD(5)
            LabelTime = DG_SNTable.Rows(0).Cells(4).Value 'UPD(6)
            PrintCodeSN = FullSTBSN.Substring(0, 22) & ">6" & FullSTBSN.Substring(22, 1) 'UPD(7)
            PrintTextSN = FullSTBSN.Substring(0, 2) & " " & FullSTBSN.Substring(2, 4) & " " & FullSTBSN.Substring(6, 2) & " " & FullSTBSN.Substring(8, 2) &
                         " " & FullSTBSN.Substring(10, 2) & " " & FullSTBSN.Substring(12, 3) & " " & FullSTBSN.Substring(15, 8) 'UPD(8)
            UploadData = {1, MACTable, HDCPName, CertName, FullSTBSN, LabelDate, LabelTime, PrintCodeSN, PrintTextSN}
        Else
            UploadData = {0}
        End If
        Return UploadData
    End Function

    '______________________________________________________________________________________________
    'функция конвертирования строки в массив
    Public SendData As Byte()
    Public Function StringToByteArray(raw As String) As Byte
        SendData = New Byte((raw.Length / 2) - 1) {}
        Dim i As Integer
        For i = 0 To SendData.Length - 1
            SendData(i) = Convert.ToByte(raw.Substring((i * 2), 2), &H10)
        Next i
    End Function
    ''------------------------------------------------------------------------------------------------------------
    'функция отправки в ком порт команд для прошивки.
    Dim arrBuffer() As Byte
    Public intSize As Integer
    Public Sub SendToCOM(ComPort As SerialPort, GeneratedRequest As String, TimeOut As Integer)
        StringToByteArray(GeneratedRequest)
        arrBuffer = New Byte(1024) {}
        Try
            ComPort.Open()
            ComPort.Write(SendData, 0, SendData.Length)
            System.Threading.Thread.Sleep(TimeOut)
            While ComPort.BytesToRead() > 0
                intSize = ComPort.Read(arrBuffer, 0, 1024)
            End While
            ComPort.Close()
        Catch ex As Exception
            ComPort.Close()
            MsgBox("Проверь настройку COM порта. Для прошивки должен быть установлен COM9")
        End Try
    End Sub
    ''------------------------------------------------------------------------------------------------------------
    'Функция получения из приемника: модель, SCID, DUID, SW, SWGS1
    Public Function GetDatafromSTB(ComPort As SerialPort, ByteRequest As String, ByteAnswer As String, ErrorLabal As Label, LabelText As String) As String
        Dim ResHex, ResHexOut, ResText
        Dim Dalay = 0, TimeOut As Integer = 0
        For i = 1 To 3
            ResHex = ""
            Dalay += 200
            'ByteRequest = "AB" '\ аргумент, байт запроса
            '\ генерация блока данных для передачи в приемник
            TimeOut = 300 + Dalay '\ таймаут, константа
            SendToCOM(ComPort, DataGenerationOneByte(ByteRequest), TimeOut) '\отправка блока данных в приемник
            ResHex += System.BitConverter.ToString(arrBuffer, 0, intSize) '/ Читаем ответ приемника из СОМ порта. Переводим массив в текст
            ResHex = Replace(ResHex, "-", "") '/ Удаляем лишние символы (бит конвертер добавляет в текст символ "-")
            If InStr(ResHex, ByteAnswer) = 13 And i < 4 Then '/ Проверка правильного ответа.
                ResHexOut = Mid(ResHex, 19, 2 * ("&H" & Mid(ResHex, 17, 2))) '/ выбираем из полученного ответа символы соответствующие SCID в HEX 
                For x = 0 To ResHexOut.Length - 1 Step 2 '/ Цикл чтения по два символа из выделенной строки
                    ResText &= ChrW(CInt("&H" & ResHexOut.Substring(x, 2))) '/ Перевод значения TextHex в Text
                Next
                Exit For
            Else
                ErrorLabal.Text = LabelText
                ErrorLabal.ForeColor = Color.Red
            End If
        Next
        Return ResText
    End Function
    ''------------------------------------------------------------------------------------------------------------
    ' Функция прошивки HDCP
    Dim DataLenght As String
    Public Function WriteHDCP(ComPort As SerialPort, SN As String, ErrorLabal As Label, LabelText As String) As Boolean
        Dim HDCpKeyByte As Byte()
        Dim HDCpKey As String
        Dim Dalay = 0, TimeOut As Integer = 0
        Dim Res As Boolean
        HDCpKeyByte = SelectByte("use fas  SELECT HDCPContent
            FROM [FAS].[dbo].[FAS_HDCP] where SerialNumber = " & SN) 'чтение файла HDCP из DB FAS_HDCP
        HDCpKey = Replace(System.BitConverter.ToString(HDCpKeyByte, 0, HDCpKeyByte.Length), "-", "")
        For i = 1 To 3
            Dalay += 200
            Dim HDCPAnswer As String ' рабочие переменные
            Dim HDCPData = "8B" & "00000000" & HDCpKey ' формируем HDCP Key длиной 316 байт (дописываем 4 байта из Init файла)+
            DataLenght = "0" & Hex(HDCPData.Length / 2) '\ если блок данных состоит из боле чем 255 байт (HDCP и Cert)
            DataLenght = DataLenght.Chars(2) & DataLenght.Chars(3) & DataLenght.Chars(0) & DataLenght.Chars(1) '\ записываем в обратном порядке байты длины
            TimeOut = 800 + Dalay '\ таймаут, константа
            SendToCOM(ComPort, DataGenerationOther(HDCPData, DataLenght), TimeOut) '\отправка блока данных в приемник
            HDCPAnswer += System.BitConverter.ToString(arrBuffer, 0, intSize)
            HDCPAnswer = Replace(HDCPAnswer, "-", "")
            If InStr(HDCPAnswer, "0B00") = 13 Then
                Res = True
                Exit For
            Else
                Res = False
                ErrorLabal.Text = LabelText
                ErrorLabal.ForeColor = Color.Red
            End If
        Next
        Return Res
    End Function

    '------------------------------------------------------------------------------------------------------------
    'Функция прошивки сертификатов
    Public Function WriteCERT(ComPort As SerialPort, SN As String, ErrorLabal As Label, LabelText As String) As Boolean
        Dim CERTKeyByte As Byte()
        Dim CertKey As String
        Dim Dalay = 0, TimeOut As Integer = 0
        Dim Res As Boolean
        CERTKeyByte = SelectByte("use fas    SELECT CertContent
        FROM [FAS].[dbo].[FAS_CERT] where SerialNumber = " & SN) 'чтение файла HDCP из DB FAS_HDCP
        CertKey = Replace(System.BitConverter.ToString(CERTKeyByte, 0, CERTKeyByte.Length), "-", "")
        For i = 1 To 3
            Dalay += 200
            Dim CertAnswer, KeyLenght As String ' рабочие переменные
            ' определяем значение длины ключа в HEX 
            KeyLenght = "0" & Hex(CertKey.Length / 2) '\ определяем длину ключа
            KeyLenght = KeyLenght.Chars(2) & KeyLenght.Chars(3) & KeyLenght.Chars(0) & KeyLenght.Chars(1) '\ записываем в обратном порядке байты длины
            Dim CertData = "A5" & "0B" & "6B6579735061636B616765" & KeyLenght & CertKey ' формируем Cert Key длиной xxx байт путем чтения файла Cert по адресу из DG_SNTable
            DataLenght = "0" & Hex(CertData.Length / 2) '\ если блок данных состоит из боле чем 255 байт (HDCP и Cert)
            DataLenght = DataLenght.Chars(2) & DataLenght.Chars(3) & DataLenght.Chars(0) & DataLenght.Chars(1) '\ записываем в обратном порядке байты длины
            TimeOut = 800 + Dalay '\ таймаут, константа
            SendToCOM(ComPort, DataGenerationOther(CertData, DataLenght), TimeOut) '\отправка блока данных в приемник
            CertAnswer += System.BitConverter.ToString(arrBuffer, 0, intSize)
            CertAnswer = Replace(CertAnswer, "-", "")
            If InStr(CertAnswer, "2500") = 13 Then
                Res = True
                Exit For
            Else
                Res = False
                ErrorLabal.Text = LabelText
                ErrorLabal.ForeColor = Color.Red
            End If
        Next
        Return Res
    End Function

    ''------------------------------------------------------------------------------------------------------------
    'Функция прошивки MAC
    Public Function WriteMAC(ComPort As SerialPort, MAC As String, ErrorLabal As Label, LabelText As String) As Boolean
        Dim Dalay = 0, TimeOut As Integer = 0
        Dim Res As Boolean
        For i = 1 To 6
            Dalay += 300
            Dim MACAnswer As String
            'добавить константу для мас текст и....
            Dim MACData = "A4" & "04" & "65746830" & "11" & StrToHex(MAC) '\ константа
            TimeOut = 200 + Dalay
            SendToCOM(ComPort, DataGenerationSNorMAC(MACData), TimeOut)
            MACAnswer += System.BitConverter.ToString(arrBuffer, 0, intSize)
            MACAnswer = Replace(MACAnswer, "-", "")
            If InStr(MACAnswer, "2400") = 13 Then
                Res = True
                Exit For
            Else
                Res = False
                ErrorLabal.Text = LabelText
                ErrorLabal.ForeColor = Color.Red
            End If
        Next
        Return Res
    End Function
    '-----------------------------------------------------------------------------------------------------------
    'Функция стерки: номера
    Public Function EraseSN(ComPort As SerialPort) As String
        SendToCOM(ComPort, DataGenerationOneByte("8A"), 100) '\отправка блока данных в приемник
        Return True
    End Function
    '------------------------------------------------------------------------------------------------------------
    'Прошивка и проверка серийного номера
    Public Function SetSN(ComPort As SerialPort, SN As String, ErrorLabal As Label, LabelText As String) As String
        Dim Dalay = 0, TimeOut As Integer
        Dim Res, SNData As String
        Dim R_SN, R_SN_OUT As String
        System.Threading.Thread.Sleep(200)
        For i = 1 To 5
            Res = ""
            Dalay += 100
            SNData = "8A" & StrToHex(SN) '\ константа
            TimeOut = Dalay
            SendToCOM(ComPort, DataGenerationSNorMAC(SNData), TimeOut)
            System.Threading.Thread.Sleep(300 + Dalay)
            R_SN = ""
            SendToCOM(ComPort, DataGenerationOneByte("82"), TimeOut)
            R_SN += System.BitConverter.ToString(arrBuffer, 0, intSize) '/Переводим массив в текст
            R_SN = Replace(R_SN, "-", "") '/ Удаляем лишние символы (бит конвертер добавляет в текст символ "-")
            If InStr(R_SN, "0200") = 13 Then '/ Проверка правильного ответа.
                'выбираем из полученного ответа символы соответствующие SN в HEX 
                R_SN_OUT = ""
                R_SN_OUT = Mid(R_SN, 17, 46) '/ длина номера 23*2 символа
                For x = 0 To R_SN_OUT.Length - 1 Step 2 '/ Цикл чтения по два символа из выделенной строки
                    Res &= ChrW(CInt("&H" & R_SN_OUT.Substring(x, 2))) '/ Перевод значения TextHex в Text
                Next
                Exit For
            Else
                ErrorLabal.Text = LabelText
                ErrorLabal.ForeColor = Color.Red
            End If
        Next
        Return Res
    End Function

End Module
