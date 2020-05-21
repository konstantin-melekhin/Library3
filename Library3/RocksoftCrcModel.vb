''' <summary>
''' Класс реализует табличный алгоритм расчёта CRC.
''' </summary>
'''<author>aave, http://soltau.ru/</author>
Public Class RocksoftCrcModel

#Region "PROPS AND FIELDS"

    ''' <summary>
    ''' Таблица предвычисленных значений для расчёта контрольной суммы.
    ''' </summary>
    Public ReadOnly CrcLookupTable(255) As UInteger

    ''' <summary>
    ''' Порядок CRC, в битах (8/16/32).
    ''' Изменение этого свойства ведёт к пересчёту таблицы.
    ''' </summary>
    Public Property CrcWidth As Integer
        Get
            Return _CrcWidth
        End Get
        Set(ByVal value As Integer)
            If _CrcWidth <> value Then
                _CrcWidth = value
                _TopBit = getBitMask(_CrcWidth - 1)
                _WidMask = (((1UI << (_CrcWidth - 1)) - 1UI) << 1) Or 1UI
                generateLookupTable()
            End If
        End Set
    End Property
    Private _CrcWidth As Integer = 32

    ''' <summary>
    ''' Образующий многочлен.
    ''' Изменение этого свойства ведёт к пересчёту таблицы.
    ''' </summary>
    Public Property Polynom As UInteger
        Get
            Return _Polynom
        End Get
        Set(ByVal value As UInteger)
            If _Polynom <> value Then
                _Polynom = value
                generateLookupTable()
            End If
        End Set
    End Property
    Private _Polynom As UInteger = &H4C11DB7

    ''' <summary>
    ''' Обращать байты сообщения?
    ''' Изменение этого свойства ведёт к пересчёту таблицы.
    ''' </summary>
    Public Property ReflectIn As Boolean
        Get
            Return _ReflectIn
        End Get
        Set(ByVal value As Boolean)
            If _ReflectIn <> value Then
                _ReflectIn = value
                generateLookupTable()
            End If
        End Set
    End Property
    Private _ReflectIn As Boolean = True

    ''' <summary>
    ''' Начальное одержимое регситра.
    ''' </summary>
    Public Property InitRegister As UInteger
        Get
            Return _InitRegister
        End Get
        Set(ByVal value As UInteger)
            If _InitRegister <> value Then
                _InitRegister = value
            End If
        End Set
    End Property
    Private _InitRegister As UInteger = &HFFFFFFFFUI

    ''' <summary>
    ''' Обращать выходное значение CRC?
    ''' </summary>
    Public Property ReflectOut As Boolean
        Get
            Return _ReflectOut
        End Get
        Set(ByVal value As Boolean)
            If _ReflectOut <> value Then
                _ReflectOut = value
            End If
        End Set
    End Property
    Private _ReflectOut As Boolean = True

    ''' <summary>
    ''' Значение, с которым XOR-ится выходное значение CRC.
    ''' </summary>
    Public Property XorOut As UInteger
        Get
            Return _XorOut
        End Get
        Set(ByVal value As UInteger)
            If _XorOut <> value Then
                _XorOut = value
            End If
        End Set
    End Property
    Private _XorOut As UInteger = &HFFFFFFFFUI

#End Region '/PROPS AND FIELDS

#Region "READ-ONLY PROPS"

    ''' <summary>
    ''' Возвращает старший разряд полинома.
    ''' </summary>
    ReadOnly Property TopBit As UInteger
        Get
            Return _TopBit
            'Return getBitMask(CrcWidth - 1) 'рассчитываем один раз при изменении порядка полинома.
        End Get
    End Property
    Private _TopBit As UInteger = getBitMask(CrcWidth - 1)

    ''' <summary>
    ''' Возвращает длинное слово со значением (2^width)-1.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property WidMask As UInteger
        Get
            Return _WidMask
            'Return (((1UI << (CrcWidth - 1)) - 1UI) << 1) Or 1UI
        End Get
    End Property
    Private _WidMask As UInteger = (((1UI << (CrcWidth - 1)) - 1UI) << 1) Or 1UI

#End Region '/READ-ONLY PROPS

#Region "CTOR"

    ''' <summary>
    ''' Конструктор, инициализированный параметрами по умолчанию для алгоритма CRC32.
    ''' </summary>
    Public Sub New()
        generateLookupTable()
    End Sub

    ''' <summary>
    ''' Инициализирует новый экземпляр параметрической модели CRC с настраиваемыми параметрами.
    ''' </summary>
    ''' <param name="width">Разрядность контрольной суммы в битах.</param>
    ''' <param name="poly">Полином.</param>
    ''' <param name="initReg">начальн7ое содержимое регистра.</param>
    ''' <param name="isReflectIn">Обращать ли входящие байты сообщения?</param>
    ''' <param name="isReflectOut">Обратить ли CRC перед финальным XOR.</param>
    ''' <param name="xorOut">Конечное значение XOR.</param>
    Public Sub New(ByVal width As Integer, ByVal poly As UInteger,
                   Optional ByVal initReg As UInteger = &HFFFFFFFFUI,
                   Optional ByVal isReflectIn As Boolean = True,
                   Optional ByVal isReflectOut As Boolean = True,
                   Optional ByVal xorOut As UInteger = &HFFFFFFFFUI)
        Me.CrcWidth = width
        Me.Polynom = poly
        Me.InitRegister = initReg
        Me.ReflectIn = isReflectIn
        Me.ReflectOut = isReflectOut
        Me.XorOut = xorOut
        generateLookupTable()
    End Sub

#End Region '/CTOR

#Region "ВЫЧИСЛЕНИЕ CRC"

    ''' <summary>
    ''' Вычисляет значение контрольной суммы переданного сообщения.
    ''' </summary>
    ''' <param name="message">Исходное сообщение, для которого нужно посчитать контрольную сумму.</param>
    ''' <returns></returns>
    Public Function ComputeCrc(ByRef message As Byte()) As UInteger
        Dim registerContent As UInteger = InitRegister 'Содержимое регистра в процессе пересчёта CRC.
        For Each b As Byte In message
            registerContent = getNextRegisterContent(registerContent, b)
        Next
        Dim finalCrc As UInteger = getFinalCrc(registerContent)
        Return finalCrc
    End Function

    ''' <summary>
    ''' Вычисляет значение контрольной суммы переданного сообщения и возвращает его в виде массива байтов.
    ''' </summary>
    ''' <param name="message">Исходное сообщение, для которого нужно посчитать контрольную сумму.</param>
    ''' <returns></returns>
    Public Function ComputeCrcAsBytes(ByVal message As Byte()) As Byte()
        Dim crc As UInteger = ComputeCrc(message)
        Dim crcBytes As Byte() = BitConverter.GetBytes(crc)
        Dim crcBytesOrdered(crcBytes.Length - 1) As Byte
        For i As Integer = 0 To crcBytes.Length - 1
            crcBytesOrdered(i) = crcBytes(crcBytes.Length - 1 - i)
        Next
        Return crcBytesOrdered
    End Function

    ''' <summary>
    ''' Обрабатывает один байт сообщения (0..255).
    ''' </summary>
    ''' <param name="prevRegContent">Содержимое регистра на предыдущем шаге.</param>
    ''' <param name="value">Значение очередного байта из сообщения.</param>
    Private Function getNextRegisterContent(ByVal prevRegContent As UInteger, ByVal value As Byte) As UInteger
        Dim uValue As UInteger = value
        If ReflectIn Then
            uValue = reflect(uValue, 8)
        End If
        Dim reg As UInteger = prevRegContent
        reg = reg Xor (uValue << (CrcWidth - 8))
        For i As Integer = 0 To 7
            If (reg And TopBit) = TopBit Then
                reg = (reg << 1) Xor Polynom
            Else
                reg <<= 1
            End If
            reg = reg And WidMask()
        Next
        Return reg
    End Function

    ''' <summary>
    ''' Возвращает значение CRC для обработанного сообщения.
    ''' </summary>
    ''' <param name="regContent">Значение регистра до финального обращения и XORа.</param>
    ''' <returns></returns>
    Private Function getFinalCrc(ByVal regContent As UInteger) As UInteger
        If ReflectOut Then
            Dim res As UInteger = XorOut Xor reflect(regContent, CrcWidth)
            Return res
        Else
            Dim res As UInteger = XorOut Xor regContent
            Return res
        End If
    End Function

#End Region '/ВЫЧИСЛЕНИЕ CRC

#Region "РАСЧЁТ ТАБЛИЦЫ"

    ''' <summary>
    ''' Вычисляет таблицу предвычисленных значений для расчёта контрольной суммы.
    ''' </summary>
    Private Sub generateLookupTable()
        For i As Integer = 0 To 255
            CrcLookupTable(i) = generateTableItem(i)
        Next
    End Sub

    ''' <summary>
    ''' Рассчитывает один байт таблицы значений для расчёта контрольной суммы
    ''' по алгоритму Rocksoft^tm Model CRC Algorithm.
    ''' </summary>
    ''' <param name="index">Индекс записи в таблице, 0..255.</param>
    Private Function generateTableItem(ByVal index As Integer) As UInteger

        Dim inbyte As UInteger = CUInt(index)

        If ReflectIn Then
            inbyte = reflect(inbyte, 8)
        End If

        Dim reg As UInteger = inbyte << (CrcWidth - 8)

        For i As Integer = 0 To 7
            If (reg And TopBit) = TopBit Then
                reg = (reg << 1) Xor Polynom
            Else
                reg <<= 1
            End If
        Next

        If ReflectIn Then
            reg = reflect(reg, CrcWidth)
        End If

        Dim res As UInteger = reg And WidMask
        Return res

    End Function

#End Region '/РАСЧЁТ ТАБЛИЦЫ

#Region "ВСПОМОГАТЕЛЬНЫЕ"

    ''' <summary>
    ''' Возвращает наибольший разряд числа.
    ''' </summary>
    ''' <param name="number">Число, разрядность которого следует определить, степень двойки.</param>
    ''' <returns></returns>
    Private Function getBitMask(ByVal number As Integer) As UInteger
        Dim res As UInteger = (1UI << number)
        Return res
    End Function

    ''' <summary>
    ''' Обращает заданное число младших битов переданного числа.
    ''' </summary>
    ''' <param name="value">Число, которое требуется обратить ("отзеркалить").</param>
    ''' <param name="bitsToReflect">Сколько младших битов числа обратить, 0..32.</param>
    ''' <returns></returns>
    ''' <remarks>Например: reflect(0x3E23, 3) == 0x3E26.</remarks>
    Private Function reflect(ByVal value As UInteger, Optional ByVal bitsToReflect As Integer = 32) As UInteger
        Dim t As UInteger = value
        Dim reflected As UInteger = value
        For i As Integer = 0 To bitsToReflect - 1
            Dim bm As UInteger = getBitMask(bitsToReflect - 1 - i)
            If (t And 1) = 1 Then
                reflected = reflected Or bm
            Else
                reflected = reflected And Not bm
            End If
            t >>= 1
        Next
        Return reflected
    End Function

#End Region '/ВСПОМОГАТЕЛЬНЫЕ

End Class '/RocksoftCrc