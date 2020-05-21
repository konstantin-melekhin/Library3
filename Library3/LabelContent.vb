' Описание операторов ZPL 
'^XA - начало этикетки
'^XZ - конец этикетки
'^MMT - определяет как будут печататься этикетки, групоой или одиночке
'^LS10 сдвиг влево всех элементов до начала отсчета
'^FT40,293^A0N,108,108^FH\^FDID^FS где
'^FT40,293 позиция FTx,y
'^A0N - поворот (Т = Normal),108,108 - высота шрифта, ширина текста

Public Module LabelContent

    Public LabelSNContent, LabelIDContent, LabelScenario As String
    'Сценарий А(этикетка 58/30 для серийника и ID номера)
    'формирование контента серийного номера
    Public Function LabelSN(Model As String, LabelDate As String, LabelTime As String, PrintTextSN As String, PrintCodeSN As String, labelCount As Integer, LabelScenario As String)
        If LabelScenario = "A" Then
            'формируется контент для сценария А
            LabelSNContent = "
^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH10,35^JMA^PR2,2^MD10^JUS^LRN^CI0^XZ
^XA
^MMT
^PW685
^LL0354
^LS0
положение и контент сервисная служба
^FO44,185^GFA,04352,04352,00068,:Z64:
eJzt1D9v1DAUAPBnMpgpb2UINt8AxlS4DR+Fj5CKAZ8uOlOdBBv9AsDnYEDgKBIdb2VA4KhCbJUjFg/hzHOuFCRaKIIBpHtjfPrd8/tjgG1sYxvb+C+i+XOCjxef4fpyBv4FQ4afnF3SED85w8sRoP6CUV54kpGRRc661hkwFkbuZl56nWfO6cpWDioWFVdKlRYdOPBl4eod50uo/ALyfeuAi8nIuraPB9GytfTzUAUtuPfaOOPAZLFIxpzokgytvCZDgxki5EMyim9GF222lsGM1aj3MDgdXXQQl1FgodT9sQqqBt0oH8ioIZKBycgVjnzFORld11n+WinTKKUX8prXbvCODctVMrQZb+ni1NADGX3XnholGYccqR6sW1oOZdHsqUIvFDgNg3OZzVaYC6WbPSjzkoyyJiPU4DoLeEBGoWUy5JmhCy0oj7Wim58ahygkFUGQoaAcy9lYD8GzPhnMelAegYzrXSSjswhmJCPMyIgB/FcDo0sGL0Ct9extvR8c9G2AnFEek4F8lwxDNd0YlZ+9S4b0kQz6i4Ing24tQJ3UyRgdDO2Y8qC71JOx08U+2jNjnE+G8WZjvJiMKqQ8Tuqdt27+hBpmF5s8hJaAyPfJ6CdDQ6qHeZYMf2bgLtVDpXpcX9W779z8qYX4QG7qQX2BHPnQtcdte/zVEEakvtAkJQOz59QXLXTqi1z56rG7/dQyC3yTR5oxgbxPRpqPKxpoPJJxlQDvyTjELKTeFlrRfMgjJw9dJe1ARo6sn2b9hrjJWzKmGXuori4IWQhRpjycY/0rZIHmlBquaU7xERm2QjvEjGPOPpGRy7WJR2T0H5Oxrmggd5tGyOD1QPPBIhmehGpsQukn49HG6N5LZNEDo3eMjINpX5bTvqSda8S0L2RAfLkxZKjewB0yrORW4gMfj9cyh3g3vR/JSLv/+Whp2Yhh2v3N3t6jihlFhlYCnXySfr60OBnDh2FEhObu9IbUVOLvQv/4yrhzXp7U1W/xS4P9K4b94RN9jN/LvzSy8wzal/O+/p5hLia2sY1tbOP/ii8gSQz/:A196
положение и контент general satellite
^FO224,217^GFA,02688,02688,00028,:Z64:
eJzt1M1KxDAQB/CEHPYYwatsHmFfoDT4RF4rVLOwh30LH0UKHjzuIxjpxYNgZBEjhIwzST+268dREDpdevkxZPJPs4zNNdc/qLP0PsGHfl1xwPI8MBWMd6UOYBe2zCbInAhMo3k04waTyWRgBi2QedmbJrMKjdUsliW7CtJ1ZsgaHVjorKpVb5DMBObRINnSl+OYsEZzzDy93OpntGpigFYY/35Ls5SdpS0A7y2SFXW2xYHVbB9LZlxdHFgUnb2SWRMPLCzQfGeXrjc5GOXicL0PB/gezMvernWonFfHRhNazLP6sHKVTKWoFZry2LdEu3JSj2Y1Gl8fn/doX0vlqH+29W8GPxv73mh/kZPpJhgw+6YQO3gY9h5PyUxn9cSCIINv+zBqtLfzaKAWuW9YzydzPAqLVoimUsP5YdRkIls9MYwaLYidsKbNfWk9kb5qsphsn/tGW5FdbHebaR/PUdP53W830z6Wo0aT97CZzkl3hTywxQ7upnPiHYsTA70fDD8VngwexB2uR7H160nwItuNaMn0aByazpa8PerDu5zNl/yxMS3odrTx72Suueb60/oEc4oKnw==:460E
модель
^FT20,32^A0N,33,33^FH\^FD" & Model & "^FS
дата
^FT320,32^A0N,33,33^FH\^FD" & LabelDate & "^FS
время
^FT490,32^A0N,33,33^FH\^FD" & LabelTime & "^FS
сн_текст 
^FT86,68^A0N,33,36^FH\^FD" & PrintTextSN & "^FS
штрихкод
^BY3,3,125^FT45,197^BCN,,N,N
^FD>;" & PrintCodeSN & "^FS
made_QC
^FT20,282^A0N,33,31^FH\^FDMade in Russia^FS
^FT448,282^A0N,33,33^FH\^FDQC Passed^FS
количество этикеток
^PQ" & labelCount & ",0,1,Y^XZ"
            '____________________________________________________________________________________________________________________________________________
        ElseIf LabelScenario = "B" Then
            LabelSNContent = "
^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH10,35^JMA^PR2,2^MD10^JUS^LRN^CI0^XZ
^XA
^MMT
^PW685
^LL0354
^LS0
положение и контент сервисная служба
^FO44,185^GFA,04352,04352,00068,:Z64:
eJzt1D9v1DAUAPBnMpgpb2UINt8AxlS4DR+Fj5CKAZ8uOlOdBBv9AsDnYEDgKBIdb2VA4KhCbJUjFg/hzHOuFCRaKIIBpHtjfPrd8/tjgG1sYxvb+C+i+XOCjxef4fpyBv4FQ4afnF3SED85w8sRoP6CUV54kpGRRc661hkwFkbuZl56nWfO6cpWDioWFVdKlRYdOPBl4eod50uo/ALyfeuAi8nIuraPB9GytfTzUAUtuPfaOOPAZLFIxpzokgytvCZDgxki5EMyim9GF222lsGM1aj3MDgdXXQQl1FgodT9sQqqBt0oH8ioIZKBycgVjnzFORld11n+WinTKKUX8prXbvCODctVMrQZb+ni1NADGX3XnholGYccqR6sW1oOZdHsqUIvFDgNg3OZzVaYC6WbPSjzkoyyJiPU4DoLeEBGoWUy5JmhCy0oj7Wim58ahygkFUGQoaAcy9lYD8GzPhnMelAegYzrXSSjswhmJCPMyIgB/FcDo0sGL0Ct9extvR8c9G2AnFEek4F8lwxDNd0YlZ+9S4b0kQz6i4Ing24tQJ3UyRgdDO2Y8qC71JOx08U+2jNjnE+G8WZjvJiMKqQ8Tuqdt27+hBpmF5s8hJaAyPfJ6CdDQ6qHeZYMf2bgLtVDpXpcX9W779z8qYX4QG7qQX2BHPnQtcdte/zVEEakvtAkJQOz59QXLXTqi1z56rG7/dQyC3yTR5oxgbxPRpqPKxpoPJJxlQDvyTjELKTeFlrRfMgjJw9dJe1ARo6sn2b9hrjJWzKmGXuori4IWQhRpjycY/0rZIHmlBquaU7xERm2QjvEjGPOPpGRy7WJR2T0H5Oxrmggd5tGyOD1QPPBIhmehGpsQukn49HG6N5LZNEDo3eMjINpX5bTvqSda8S0L2RAfLkxZKjewB0yrORW4gMfj9cyh3g3vR/JSLv/+Whp2Yhh2v3N3t6jihlFhlYCnXySfr60OBnDh2FEhObu9IbUVOLvQv/4yrhzXp7U1W/xS4P9K4b94RN9jN/LvzSy8wzal/O+/p5hLia2sY1tbOP/ii8gSQz/:A196
положение и контент general satellite
^FO224,217^GFA,02688,02688,00028,:Z64:
eJzt1M1KxDAQB/CEHPYYwatsHmFfoDT4RF4rVLOwh30LH0UKHjzuIxjpxYNgZBEjhIwzST+268dREDpdevkxZPJPs4zNNdc/qLP0PsGHfl1xwPI8MBWMd6UOYBe2zCbInAhMo3k04waTyWRgBi2QedmbJrMKjdUsliW7CtJ1ZsgaHVjorKpVb5DMBObRINnSl+OYsEZzzDy93OpntGpigFYY/35Ls5SdpS0A7y2SFXW2xYHVbB9LZlxdHFgUnb2SWRMPLCzQfGeXrjc5GOXicL0PB/gezMvernWonFfHRhNazLP6sHKVTKWoFZry2LdEu3JSj2Y1Gl8fn/doX0vlqH+29W8GPxv73mh/kZPpJhgw+6YQO3gY9h5PyUxn9cSCIINv+zBqtLfzaKAWuW9YzydzPAqLVoimUsP5YdRkIls9MYwaLYidsKbNfWk9kb5qsphsn/tGW5FdbHebaR/PUdP53W830z6Wo0aT97CZzkl3hTywxQ7upnPiHYsTA70fDD8VngwexB2uR7H160nwItuNaMn0aByazpa8PerDu5zNl/yxMS3odrTx72Suueb60/oEc4oKnw==:460E
модель
^FT20,32^A0N,33,33^FH\^FD" & Model & "^FS
дата
^FT320,32^A0N,33,33^FH\^FD" & LabelDate & "^FS
время
^FT490,32^A0N,33,33^FH\^FD" & LabelTime & "^FS
сн_текст 
^FT86,68^A0N,33,36^FH\^FD" & PrintTextSN & "^FS
штрихкод
^BY3,3,125^FT45,197^BCN,,N,N
^FD>;" & PrintCodeSN & "^FS
made_QC
^FT20,282^A0N,33,31^FH\^FDMade in Russia^FS
^FT448,282^A0N,33,33^FH\^FDQC Passed^FS
количество этикеток
^PQ" & labelCount & ",0,1,Y^XZ"
        End If
        Return (LabelSNContent)
    End Function
    'формирование контента ID номера
    Public Function LabelID(SC_Text As String, labelCount As Integer, LabelScenario As String)
        If LabelScenario = "A" Then
            LabelIDContent = "       
^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2^MD10^JUS^LRN^CI0^XZ
^XA
^MMT
^PW685
^LL0354
^LS0
_положение_ID_
^FT40,320
_ориентация и размер надписи
^A0N,108,108
^FH\^FDID^FS
_положение_номера ID_
^FT144,320
_ориентация и размер надписи
^A0N,108,72
^FH\^FD" & SC_Text & "^FS
_масштаб штрихкода_5 и 3 не трогать, 130 высота штрихкода
^BY5,3,150
_положение_штрихкода_
^FT60,220
_ориентация штрихкода
^BCN,,N,N
^FD>;" & SC_Text & "^FS
^PQ" & labelCount & ",0,1,Y^XZ"

        ElseIf LabelScenario = "B" Then
            LabelIDContent = "       
^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^MD10^JUS^LRN^CI0^XZ
^XA
^MMT
^PW650
^LL0201
^LS0
_положение_номера ID_
^FT166,185
_ориентация и размер надписи
^A0N,58,57
^FH\^FD" & SC_Text & "^FS
_положение_ID_
^FT82,185
_ориентация и размер надписи
^A0N,58,57
^FH\^FDID:^FS
_масштаб штрихкода_5 и 3 не трогать, 130 высота штрихкода
^BY5,3,94
_положение_штрихкода_
^FT36,133
_ориентация штрихкода
^BCN,,N,N
^FD>;" & SC_Text & "^FS
^PQ" & labelCount & ",0,1,Y^XZ"
        End If
        Return (LabelIDContent)
    End Function



    Public Function LabelSN_Micron(DevEUI As String, ShortSN As String, QR_Code As String)
        LabelIDContent = " 
^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA~MD15^JUS^LRN^CI0^XZ
^XA
^MMT
^PW685
^LL0354
^LS0
^FO448,32^GFA,02048,02048,00016,:Z64:
eJxjYKA5YL0DZcRCKPYfUH59A0Q+BMoPdcAufwOmv4E8/TEU2n8XBqD2/4cBqHwoDDhgl78bC4UMZOrHb/9QA9T2P0H9sTBApP3A+GdECltQ+mO8i+R+sDyCD0o/jKF45O9QqJ9U+yn1P7XDHy39DzYADz+U+GdAKX/A4Y9U/qDIE9IPi3+k8ock/YTsp3b6p3L8j4JBDAD6fWCc:BD3B
^FO608,32^GFA,02048,02048,00008,:Z64:
eJzdlE9Kw0AUxhNiTaBgWxDsojjXiFLbLHoAF4obtYgXaLBIFy3JSlz1BOJJXMwJxAOIeASXXWjq4vteYcZOmkVXvs2PJG/me//yPO+/W1ZYvAd7/K7I4BP0v/jiG5jZ54tyvfMzhy55Rb9GDtbe+IK6ybT8/GBgCTaBqA12OmDcBSd8Hu05yO89+ivm3VhY8Vlm63V5flxR7/jFoZeYOj4p/VGsk6L/ySU4cekdgjE5Y337C/M+ud8y/5RkmEFhMjwgw4q8wHyEy2Itt277TKDHBDLhM3gnnDv4QX/HPLaa2423pkGZB5kPmc9R1XmuOF+22fkN38Egr0ZXnVwm9VOWf59xp8zndhe8ET7wO/sT81yb56KYAlZ/Nu2nkVfOcUJ/1nM1T7KfrunIfng7pn4tB9Uj+QQecQ7TeTln7IfUR/bAn/83glBLacS3JIvEiDflXnDNU0pmr55xri59HeI+r6GZYGKEkfBRo4C+njJcxB1q9DdkvTYyh3+g8T/4P6zDknssI+tIf7U/19svWvSc4Q==:44BB
^FO416,64^GFA,01792,01792,00008,:Z64:
eJzlkztOw0AQQGdNQYEUl1BEjgRdToBSpeMIHIAL5AasSIsCZegQ6VKFlgLZHR1H4NdBQxcJpDgS+8bSrrxxkKjINE/rmXm7tmdF/n1M7Q9Mmfs8d2kzCnjh18s3HPR/5/8kfwjbMHX1sp3X9qmvin14At/Yr4A2r2eh58d3Sn83OH8S8b+oP7LPun6NDJbUz+HZ6veQifX7FlC/ZxBbpNNnxxZsD9Zj9uUEHf6H8u/i2kEHxDysLu+684ieg+9gbuEMXpIfO5on1o/0H0X8B+qHC/rHDbyCTfPLsvqvQzgnfw/34J2t98ZC7wdzK6/W3y+cpyF1ZcDIPEkP6nx/4C8i/up+4G24Hwn3L5w3Zcb7ddAnzGf1IM299c47S/zZTb1/97iPUDY4lkKBkN8=:68E8
^FO384,32^GFA,02048,02048,00008,:Z64:
eJzllLFKA0EQhi+IRhE0hUGLYPIIV8ZCvUewOWwMbCHWG9CkUEnAF/IR7hHsbPMI9wRGcP4/sHMzeqCdf/MVOzvz78zuZtk/1BX4BN5/H35SCccrYYzCkNucYv0S8UPs318KOyp/CT+LOuUY6wPwkPuRL1vZ+/o8n9Ie4p9V/Ax+73Kb82jX8dQHeX7Wi06/YhCyv4t1yrK06xwshaeV8AL7fzxPnfo7Luz8u1XqazN/5xw56p4jP+e+rfJz/pwn486Y/5f3iwotqeda9jJTzM96rO/5HY2QD/PTc/V0BG7mCXrznDn9YH+1vPe2k7Vjo1/Oe9Pzpb+5up+3+j1gnf3TKnpp/BQ+eH+Gyu+1c15P+v0+BPjx3i/85ohr3H+VPzh8fLV96rhC5fvr/1mL/wPjeD7Op9EP/gfBrqP7Qb1I/c772xe3PmRf90b8dgctOQHXdUJTn2BYkZc=:6544
^FO384,64^GFA,00896,00896,00004,:Z64:
eJy1kstKw0AUhv9xIsmi2IgEIqW2+AASVHBWbR/BTcGN4CNkI24K5tHqG7hxP48QECEgMv5zaZN6Q8Qe+BbnMuf85zDA/5mYX0KaxgHNQNkhmKx+4KnzPtgOEkhckQbioQae2fhuBhwx2WP+ZLx+I/bTIOSbQX+az/iScLRY0JmQmIE+dYy42EJB1ApSk1CPiEzZ8J51rxpiWbK/+oSNC11DvLC5YW32iyMfj9m7QUTiYQq5R23xjLM1oqx0OUfXdnmv/o3Xe6ZI7jngvUa1x+60Uc/9hqxXrC9yT8r6HvtMTeXI5p0hCfuE2Rd8d5q3cwb0D+0BHzvlSauV+t3xVlj/qz1Wd/+oq6CuXPk9bN7qX7/RPj5pNvfICwjzVkljKnF963wXt/nzsr1L+Dfbsnduqm9w:8987
^FO352,0^GFA,02816,02816,00008,:Z64:
eJzllj1Ow0AQhf0TsONETiiiGBPFKSk4AEKRcgkOQkEZCZcUOUKKlAgh6hQULjmGS0roLIgUinljaUa7YBdQwDRfrJ15O7N53sRx/koMO8RkTpwuiYOKGJREH8yukHdOnEyIYd+sP7pE3U7Sm6nELVhJuquvebAy60en5n6Ohub87i10czN7d+Y6W0S2/ALM1TM4BjOsZy31Od9Hvcv6fK5Lcx3r6bB9f+mMmDhmnmyV/gMW2Ceh3Cd1ZP+9F9mvs8iJN0T/vaD8gjjFgQ7NYzSO0TX6UPMmz2ABluArccDzrokRz9mR+jb/xfBfgDx/gw9436Ktuc4WNn/Y/K3Z1B+cz/Mzu8g/XimuZV4AX3qlRb9AfiWZ4v1OcM58/4QwYHyB/veFoC34fgvwfcboj+833qcm/NvHelzJem8j9W3vCTNteb+0De4nUH61zRdivg4a1Ofo3efGfdyN1K/9AH9/6wf0V/tfxdj2fkLfdr5ZJfOjJyxY7iPev6kf2H+TxDyP9gOH9jf3mZ6Z9dv6u6nf+Fnn6fv0t34/f/r86/agl9P/H/cNfX2UxEWB/VFwiPpHmsfd7wT/SXwC7YujDg==:D877
^FO511,81^GB0,12,8^FS
^FO476,112^GB0,12,9^FS
^FO511,112^GB0,12,9^FS
^FO545,112^GB0,12,9^FS
^FO476,81^GB0,12,9^FS
^FO545,38^GB0,12,9^FS
^FO476,37^GB0,13,9^FS
^FO476,124^GB78,0,10^FS
^FO476,93^GB78,0,10^FS
^FO476,71^GB78,0,11^FS
^FO476,49^GB78,0,11^FS
^FT41,325^BQN,2,4
^FDMA," & QR_Code & "^FS
^FT352,240^A0B,33,33^FH\^FD" & ShortSN & "^FS
^FT605,298^A0B,33,33^FH\^FD" & DevEUI & "^FS
^FT466,330^BQN,2,5
^FDMA," & DevEUI & "^FS
^PQ3,0,1,Y^XZ
"
        Return (LabelIDContent)
    End Function

End Module
