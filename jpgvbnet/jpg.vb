Option Explicit On

Public Class jpg

    Private Structure JpegType           'some type definitions (for coherence)
        Public Property Rows As Integer             'image height
        Public Property Cols As Integer             'image width
        Public Property SamplesY As Integer         'sampling ratios
        Public Property SamplesCbCr As Integer
        Public Property QuantTableY As Integer      'quantization table numbers
        Public Property QuantTableCbCr As Integer
        Public Property HuffDCTableY As Integer     'huffman table numbers
        Public Property HuffDCTableCbCr As Integer
        Public Property HuffACTableY As Integer
        Public Property HuffACTableCbCr As Integer
        Public Property NumComp As Integer          'number of components
    End Structure

    Private Structure HuffmanEntry               'a type for huffman tables
        Public Property Index As Long
        Public Property Code As Integer
        Public Property Length As Integer
    End Structure

    Public bmp As Bitmap

    'a few global variables
    Private curByte As Integer, curBits As Integer, jfile As Integer
    Private EOI As Integer, DCTables As Integer, ACTables As Integer, QTables As Integer
    Private Image As JpegType

    'All large arrays are now dynamic
    Private HuffmanDC(1, 255) As HuffmanEntry
    Private HuffmanAC(1, 255) As HuffmanEntry
    Private Dct(7, 7, 7, 7) As Integer
    Private Zig1(63) As Integer
    Private Zig2(63) As Integer
    Private ZigIndex As Byte
    Private ScanVector() As Byte
    Public findex As Long, flen As Long

    Private Sub Read(ByRef n1 As Byte, ByRef n2 As Byte)

        n1 = Zig1(ZigIndex)
        n2 = Zig2(ZigIndex)
        ZigIndex = ZigIndex + 1

    End Sub

    Private Function Decode(inArray(,) As HuffmanEntry, hnum As Integer) As Integer

        Dim n1 As Byte, n2 As Byte, i As Integer, l As Integer
        Dim CurVal As Long
        Dim MatchFound As Integer

        If GetByte() = 255 Then
            n1 = GetByte()

            If n1 >= &HD0 And n1 <= &HD7 Then
                n2 = 2 ^ curBits - 1
                findex = findex - 2
                If curByte And n2 = n2 Then    'if the remaining bits are 1
                    EOI = 1
                    Decode = 0
                    Exit Function
                End If
            Else
                findex = findex - 2
            End If
        Else
            findex = findex - 1
        End If

        CurVal = 0
        MatchFound = -1
        For l = 1 To 16    'cycle through 16 possible Huffman lengths
            CurVal = CurVal * 2 + NextBit()

            If EOI Then
                Return 0
            End If

            For i = 0 To 255              'look for a match in the Huffman table
                If inArray(hnum, i).Length > l Then Exit For
                If inArray(hnum, i).Length = l Then
                    If inArray(hnum, i).Index = CurVal Then
                        MatchFound = i
                        Exit For
                    End If
                End If
            Next i

            If MatchFound > -1 Then Exit For

        Next l

        If MatchFound = -1 Then
            Beep()
            Return -1
        End If

        Return inArray(hnum, MatchFound).Code  'return the appropriate code

    End Function

    Private Sub GetBlock(ByRef vector(,) As Integer, HuffDC(,) As HuffmanEntry, HuffDCNum As Integer, HuffAC(,) As HuffmanEntry, HuffACNum As Integer, Quant(,,) As Integer, QuantNum As Integer, ByRef dcCoef As Integer)

        Dim d As Integer, XPos As Byte, YPos As Byte, Sum As Integer
        Dim bits As Byte, zeros As Byte, bitVal As Integer, ACCount As Byte
        Dim j As Integer, X As Integer, Y As Integer, v As Integer, u As Integer
        Dim temp As Integer, temp0 As Byte, Add1 As Integer
        Dim array2(7, 7) As Integer

        EOI = 0
        temp0 = Decode(HuffDC, HuffDCNum)   'Get the DC coefficient

        If EOI Then d = 0

        dcCoef = dcCoef + ReceiveBits(temp0)
        array2(0, 0) = dcCoef * Quant(QuantNum, 0, 0)
        XPos = 0 : YPos = 0
        ZigIndex = 1
        ACCount = 1

        Do
            d = Decode(HuffAC, HuffACNum)
            If EOI Then d = 0
            zeros = d \ 16
            bits = d And 15
            bitVal = ReceiveBits(bits)
            If zeros = 0 And bits = 0 Then   'EOB Encountered
                Exit Do
            ElseIf zeros = 15 And bits = 0 Then  'ZRL encountered
                ZigIndex = ZigIndex + 15
                ACCount = ACCount + 16
            Else
                ZigIndex = ZigIndex + zeros
                ACCount = ACCount + zeros
                If ACCount >= 64 Then Exit Do

                Read(XPos, YPos)
                array2(XPos, YPos) = bitVal * Quant(QuantNum, XPos, YPos)
                ACCount = ACCount + 1
            End If
            If ACCount >= 64 Then Exit Do
        Loop

        If HuffDCNum = Image.HuffDCTableY Then Add1 = 128

        For X = 0 To 7            'the IDCT routine (pretty fast)
            For Y = 0 To 7
                Sum = 0
                For v = 0 To 7
                    For u = 0 To 7
                        temp = array2(u, v)
                        If temp Then temp = temp * Dct(X, Y, u, v) / 64
                        Sum = Sum + temp
                    Next u
                Next v
                vector(X, Y) = Sum + Add1
            Next Y
        Next X

        Erase array2

    End Sub

    Private Function GetByte() As Integer

        GetByte = ScanVector(findex)
        findex = findex + 1

    End Function

    Private Function GetHuffTables() As Integer

        Dim HuffAmount(16) As Integer
        Dim l0 As Long, c0 As Integer, temp0 As Integer, temp1 As Integer, total As Integer
        Dim i As Integer, t0 As Integer, CurNum As Long, CurIndex As Integer, j As Integer

        l0 = GetWord()
        c0 = 2
        Do
            temp0 = GetByte()
            c0 = c0 + 1
            t0 = (temp0 And 16) \ 16
            temp0 = temp0 And 15
            Select Case t0
                Case 0        'DC Table
                    total = 0
                    For i = 1 To 16
                        temp1 = GetByte()
                        c0 = c0 + 1
                        total = total + temp1
                        HuffAmount(i) = temp1
                    Next i
                    For i = 0 To total - 1
                        HuffmanDC(temp0, i).Code = GetByte()
                        c0 = c0 + 1
                    Next i
                    CurNum = 0
                    CurIndex = -1
                    For i = 1 To 16
                        For j = 1 To HuffAmount(i)
                            CurIndex = CurIndex + 1
                            HuffmanDC(temp0, CurIndex).Index = CurNum
                            HuffmanDC(temp0, CurIndex).Length = i
                            CurNum = CurNum + 1
                        Next j
                        CurNum = CurNum * 2
                    Next i
                    DCTables = DCTables + 1
                Case 1
                    total = 0
                    For i = 1 To 16
                        temp1 = GetByte()
                        c0 = c0 + 1
                        total = total + temp1
                        HuffAmount(i) = temp1
                    Next i
                    For i = 0 To total - 1
                        HuffmanAC(temp0, i).Code = GetByte()
                        c0 = c0 + 1
                    Next i
                    CurNum = 0
                    CurIndex = -1
                    For i = 1 To 16
                        For j = 1 To HuffAmount(i)
                            CurIndex = CurIndex + 1
                            HuffmanAC(temp0, CurIndex).Index = CurNum
                            HuffmanAC(temp0, CurIndex).Length = i
                            CurNum = CurNum + 1
                        Next j
                        CurNum = CurNum * 2
                    Next i
                    ACTables = ACTables + 1
            End Select
        Loop Until c0 >= l0

        Return 1

    End Function

    Private Function GetImageAttr() As Integer

        Dim temp4 As Long, temp0 As Integer, temp1 As Integer, i As Integer, id As Integer

        temp4 = GetWord()               'Length of segment
        temp0 = GetByte()               'Data precision
        If temp0 <> 8 Then
            Return 0                    'we do not support 12 or 16-bit samples
        End If

        Image.Rows = GetWord()          'Image Height
        Image.Cols = GetWord()          'Image Width
        temp0 = GetByte()               'Number of components
        For i = 1 To temp0
            id = GetByte()

            Select Case id
                Case 1
                    temp1 = GetByte()
                    Image.SamplesY = (temp1 And 15) * (temp1 \ 16)
                    Image.QuantTableY = GetByte()
                Case 2, 3
                    temp1 = GetByte()
                    Image.SamplesCbCr = (temp1 And 15) * (temp1 \ 16)
                    Image.QuantTableCbCr = GetByte()
            End Select
        Next i

        Return 1

    End Function

    Private Function GetQuantTables(ByRef inArray(,,) As Integer) As Integer

        Dim l0 As Long, c0 As Integer, temp0 As Byte
        Dim xp As Byte, yp As Byte, i As Byte

        l0 = GetWord()
        c0 = 2
        Do
            temp0 = GetByte()
            c0 = c0 + 1
            If temp0 And &HF0 Then
                Return 0                'we don't support 16-bit tables
            End If
            temp0 = temp0 And 15
            ZigIndex = 0
            xp = 0 : yp = 0
            For i = 0 To 63
                Read(xp, yp)
                inArray(temp0, xp, yp) = GetByte()
                c0 = c0 + 1
            Next i
            QTables = QTables + 1
        Loop Until c0 >= l0

        Return 1

    End Function

    Private Function GetSOI() As Integer

        Dim d As Integer

        findex = 0
        d = 0
        If GetByte() = 255 Then
            If GetByte() = &HD8 Then d = 1
        End If

        Return d

    End Function

    Private Function GetSOS() As Integer

        Dim temp4 As Long, temp0 As Byte, temp1 As Byte, temp2 As Byte
        Dim i As Integer

        temp4 = GetWord()
        temp0 = GetByte()
        If temp0 <> 1 And temp0 <> 3 Then
            Return 0
        End If

        Image.NumComp = temp0
        For i = 1 To temp0
            temp1 = GetByte()

            Select Case temp1
                Case 1
                    temp2 = GetByte()
                    Image.HuffACTableY = temp2 And 15
                    Image.HuffDCTableY = temp2 \ 16
                Case 2
                    temp2 = GetByte()
                    Image.HuffACTableCbCr = temp2 And 15
                    Image.HuffDCTableCbCr = temp2 \ 16
                Case 3
                    temp2 = GetByte()
                    Image.HuffACTableCbCr = temp2 And 15
                    Image.HuffDCTableCbCr = temp2 \ 16
                Case Else
                    Return 0
            End Select 'temp1
        Next i
        findex = findex + 3

        Return 1

    End Function

    Private Function GetWord() As Long

        Dim l0 As Long

        l0 = CLng(GetByte()) * 256
        l0 = l0 + GetByte()

        Return l0

    End Function

    Private Function NextBit() As Integer

        Dim t0 As Byte, v0 As Byte

        t0 = 2 ^ curBits
        v0 = -((curByte And t0) <> 0)
        curBits = curBits - 1
        If curBits < 0 Then
            curBits = 7 : curByte = GetByte()

            If curByte = 255 Then
                If GetByte() = &HD9 Then
                    EOI = 1
                    Return 0
                End If
            End If
        End If

        Return v0

    End Function

    Public Function DoJPG(FName As String) As Bitmap

        Try

            jfile = FreeFile()

            'The only place where the file is actually read
            ScanVector = System.IO.File.ReadAllBytes(FName)
            ReDim Preserve ScanVector(ScanVector.Length + 15)
            findex = 0
            flen = FileSystem.FileLen(FName)

            If GetSOI() = 0 Then
                MsgBox("This is not a valid JPEG/JFIF image.", vbCritical, "Error")
                Exit Try
            End If

            Do While findex < flen
                ProcesScan()    'more than one scan in file?
            Loop

        Catch ex As Exception

            MsgBox("There was an error reading the file:" + vbCrLf + ex.Message + vbCrLf + ex.StackTrace, vbCritical, "Error")

        End Try

        Return bmp

    End Function

    Private Sub ProcesScan()

        Dim r As Integer, g As Integer, b As Integer, Y As Integer
        Dim i As Integer, j As Integer, Restart As Long
        Dim d As Integer, temp0 As Integer, temp1 As Long
        Dim XPos As Integer, YPos As Integer, dcY As Integer, dcCb As Integer, dcCr As Integer
        Dim xindex As Integer, yindex As Integer, mcu As Integer
        Dim i2 As Integer, j2 As Integer, cb As Integer, cr As Integer
        Dim xj As Integer, yi As Integer

        Dim YVector1(7, 7) As Integer               '4 vectors for Y attribute
        Dim YVector2(7, 7) As Integer               '(not all may be needed)
        Dim YVector3(7, 7) As Integer
        Dim YVector4(7, 7) As Integer
        Dim CbVector(7, 7) As Integer               '1 vector for Cb attribute
        Dim CrVector(7, 7) As Integer               '1 vector for Cr attribute
        Dim QuantTable(1, 7, 7) As Integer          '2 quantization tables (Y, CbCr)

        QTables = 0     'Initialize some checkpoint variables
        ACTables = 0
        DCTables = 0
        Restart = 0

        Do      'Search for markers
            If GetByte() = 255 Then         'Marker Found
                d = GetByte()

                Select Case d               'which one is it?
                    Case &HC0               'SOF0
                        If GetImageAttr() = 0 Then
                            Throw New Exception("Error getting Start Of Frame 0 Marker.")
                        End If
                    Case &HC1               'SOF1
                        If GetImageAttr() = 0 Then
                            Throw New Exception("Error getting Start Of Frame 1 Marker.")
                        End If
                    Case &HC9               'SOF9
                        Throw New Exception("Arithmetic Coding Not Supported.")
                    Case &HC4               'DHT
                        If ACTables < 2 Or DCTables < 2 Then
                            If GetHuffTables() = 0 Then
                                Throw New Exception("Error getting Huffman tables.")
                            End If
                        End If
                    Case &HCC               'DAC
                        Throw New Exception("Arithmetic Coding Not Supported.")
                    Case &HDA               'SOS
                        If GetSOS() = 0 Then
                            Throw New Exception("Error getting SOS marker.")
                        End If
                        If (DCTables = 2 And ACTables = 2 And QTables = 2) Or Image.NumComp = 1 Then
                            EOI = 0
                            Exit Do         'Go on to secondary control loop
                        Else
                            Throw New Exception("Unexpected file format.")
                        End If
                    Case &HDB               'DQT
                        If QTables < 2 Then
                            If GetQuantTables(QuantTable) = 0 Then
                                Throw New Exception("Error getting quantization tables.")
                            End If
                        End If
                    Case &HDD               'DRI
                        Restart = GetWord()
                    Case &HE0               'APP0
                        temp1 = GetWord()    'Length of segment
                        findex = findex + 5
                        temp0 = GetByte()    'Major revision
                        temp0 = GetByte()    'Minor revision
                        temp0 = GetByte()    'Density definition
                        temp0 = GetByte()    'X-Density
                        temp0 = GetByte()    'Y-Density
                        temp0 = GetByte()    'Thumbnail Width
                        temp1 = GetByte()    'Thumbnail Height
                    Case &HFE              'COM
                        findex = findex + GetWord() - 2
                End Select 'd
            End If 'Marker found
        Loop Until findex >= flen

        XPos = 0 : YPos = 0               'Initialize active variables
        dcY = 0 : dcCb = 0 : dcCr = 0
        xindex = 0 : yindex = 0 : mcu = 0
        r = 0 : g = 0 : b = 0

        curBits = 7          'Start with the seventh bit
        curByte = GetByte()    'Of the first byte

        If findex < flen Then

            'resize the given control to image's dimensions
            bmp = New Bitmap(Image.Cols, Image.Rows)

            Select Case Image.NumComp        'How many components does the image have?
                Case 3                           '3 components (Y-Cb-Cr)
                    Select Case Image.SamplesY   'What's the sampling ratio of Y to CbCr?
                        Case 4                       '4 pixels to 1
                            Do                       'Process 16x16 blocks of pixels
                                GetBlock(YVector1, HuffmanDC, Image.HuffDCTableY, HuffmanAC, Image.HuffACTableY, QuantTable, Image.QuantTableY, dcY)
                                GetBlock(YVector2, HuffmanDC, Image.HuffDCTableY, HuffmanAC, Image.HuffACTableY, QuantTable, Image.QuantTableY, dcY)
                                GetBlock(YVector3, HuffmanDC, Image.HuffDCTableY, HuffmanAC, Image.HuffACTableY, QuantTable, Image.QuantTableY, dcY)
                                GetBlock(YVector4, HuffmanDC, Image.HuffDCTableY, HuffmanAC, Image.HuffACTableY, QuantTable, Image.QuantTableY, dcY)
                                GetBlock(CbVector, HuffmanDC, Image.HuffDCTableCbCr, HuffmanAC, Image.HuffACTableCbCr, QuantTable, Image.QuantTableCbCr, dcCb)
                                GetBlock(CrVector, HuffmanDC, Image.HuffDCTableCbCr, HuffmanAC, Image.HuffACTableCbCr, QuantTable, Image.QuantTableCbCr, dcCr)
                                'YCbCr vectors have been obtained
                                For i = 0 To 7           'Draw top left 8x8 pixels
                                    For j = 0 To 7
                                        Y = YVector1(i, j)
                                        i2 = i \ 2
                                        j2 = j \ 2
                                        cb = CbVector(i2, j2)
                                        cr = CrVector(i2, j2)
                                        ToRGB(Y, cb, cr, r, g, b)
                                        xj = xindex + j
                                        yi = yindex + i
                                        If xj < Image.Cols And yi < Image.Rows Then bmp.SetPixel(xj, yi, Color.FromArgb(r, g, b))
                                    Next j
                                Next i
                                For i = 0 To 7           'Draw top right 8x8 pixels
                                    For j = 8 To 15
                                        Y = YVector2(i, j - 8)
                                        i2 = i \ 2
                                        j2 = j \ 2
                                        cb = CbVector(i2, j2)
                                        cr = CrVector(i2, j2)
                                        ToRGB(Y, cb, cr, r, g, b)
                                        xj = xindex + j
                                        yi = yindex + i
                                        If xj < Image.Cols And yi < Image.Rows Then bmp.SetPixel(xj, yi, Color.FromArgb(r, g, b))
                                    Next j
                                Next i
                                For i = 8 To 15          'Draw bottom left 8x8 pixels
                                    For j = 0 To 7
                                        Y = YVector3(i - 8, j)
                                        i2 = i \ 2
                                        j2 = j \ 2
                                        cb = CbVector(i2, j2)
                                        cr = CrVector(i2, j2)
                                        ToRGB(Y, cb, cr, r, g, b)
                                        xj = xindex + j
                                        yi = yindex + i
                                        If xj < Image.Cols And yi < Image.Rows Then bmp.SetPixel(xj, yi, Color.FromArgb(r, g, b))
                                    Next j
                                Next i
                                For i = 8 To 15          'Draw bottom right 8x8 pixels
                                    For j = 8 To 15
                                        Y = YVector4(i - 8, j - 8)
                                        i2 = i \ 2
                                        j2 = j \ 2
                                        cb = CbVector(i2, j2)
                                        cr = CrVector(i2, j2)
                                        ToRGB(Y, cb, cr, r, g, b)
                                        xj = xindex + j
                                        yi = yindex + i
                                        If xj < Image.Cols And yi < Image.Rows Then bmp.SetPixel(xj, yi, Color.FromArgb(r, g, b))
                                    Next j
                                Next i
                                xindex = xindex + 16
                                If xindex >= Image.Cols Then xindex = 0 : yindex = yindex + 16 : mcu = 1
                                If mcu = 1 And Restart <> 0 Then 'Execute the restart interval
                                    curByte = GetByte()
                                    curByte = GetByte()
                                    curByte = GetByte()
                                    curBits = 7
                                    dcY = 0 : dcCb = 0 : dcCr = 0 : mcu = 0 'Reset the DC value
                                End If
                            Loop Until findex >= flen Or yindex >= Image.Rows
                        Case 2           '2 pixels to 1
                            Do
                                GetBlock(YVector1, HuffmanDC, Image.HuffDCTableY, HuffmanAC, Image.HuffACTableY, QuantTable, Image.QuantTableY, dcY)
                                GetBlock(YVector2, HuffmanDC, Image.HuffDCTableY, HuffmanAC, Image.HuffACTableY, QuantTable, Image.QuantTableY, dcY)
                                GetBlock(CbVector, HuffmanDC, Image.HuffDCTableCbCr, HuffmanAC, Image.HuffACTableCbCr, QuantTable, Image.QuantTableCbCr, dcCb)
                                GetBlock(CrVector, HuffmanDC, Image.HuffDCTableCbCr, HuffmanAC, Image.HuffACTableCbCr, QuantTable, Image.QuantTableCbCr, dcCr)
                                'YCbCr vectors have been obtained
                                For i = 0 To 7             'Draw left 8x8 pixels
                                    For j = 0 To 7
                                        Y = YVector1(i, j)
                                        j2 = j \ 2
                                        cb = CbVector(i, j2)
                                        cr = CrVector(i, j2)
                                        ToRGB(Y, cb, cr, r, g, b)
                                        xj = xindex + j
                                        yi = yindex + i
                                        If xj < Image.Cols And yi < Image.Rows Then bmp.SetPixel(xj, yi, Color.FromArgb(r, g, b))
                                    Next j
                                Next i
                                For i = 0 To 7             'Draw right 8x8 pixels
                                    For j = 8 To 15
                                        Y = YVector2(i, j - 8)
                                        j2 = j \ 2
                                        cb = CbVector(i, j2)
                                        cr = CrVector(i, j2)
                                        ToRGB(Y, cb, cr, r, g, b)
                                        xj = xindex + j
                                        yi = yindex + i
                                        If xj < Image.Cols And yi < Image.Rows Then bmp.SetPixel(xj, yi, Color.FromArgb(r, g, b))
                                    Next j
                                Next i
                                xindex = xindex + 16
                                If xindex >= Image.Cols Then xindex = 0 : yindex = yindex + 8 : mcu = 1
                                If mcu = 1 And Restart <> 0 Then 'execute the restart interval
                                    curByte = GetByte()
                                    curByte = GetByte()
                                    curByte = GetByte()
                                    curBits = 7
                                    dcY = 0 : dcCb = 0 : dcCr = 0 : mcu = 0
                                End If
                            Loop Until findex >= flen Or yindex >= Image.Rows
                        Case 1        '1 pixel to 1
                            Do
                                GetBlock(YVector1, HuffmanDC, Image.HuffDCTableY, HuffmanAC, Image.HuffACTableY, QuantTable, Image.QuantTableY, dcY)
                                GetBlock(CbVector, HuffmanDC, Image.HuffDCTableCbCr, HuffmanAC, Image.HuffACTableCbCr, QuantTable, Image.QuantTableCbCr, dcCb)
                                GetBlock(CrVector, HuffmanDC, Image.HuffDCTableCbCr, HuffmanAC, Image.HuffACTableCbCr, QuantTable, Image.QuantTableCbCr, dcCr)
                                'YCbCr vectors have been obtained
                                For i = 0 To 7            'Draw 8x8 pixels
                                    For j = 0 To 7
                                        Y = YVector1(i, j)
                                        i2 = i \ 2
                                        j2 = j \ 2
                                        cb = CbVector(i2, j2)
                                        cr = CrVector(i2, j2)
                                        ToRGB(Y, cb, cr, r, g, b)
                                        xj = xindex + j
                                        yi = yindex + i
                                        If xj < Image.Cols And yi < Image.Rows Then bmp.SetPixel(xj, yi, Color.FromArgb(r, g, b))
                                    Next j
                                Next i
                                xindex = xindex + 8
                                If xindex >= Image.Cols Then xindex = 0 : yindex = yindex + 8 : mcu = 1
                                If mcu = 1 And Restart <> 0 Then 'execute the restart interval
                                    curByte = GetByte()
                                    curByte = GetByte()
                                    curByte = GetByte()
                                    curBits = 7
                                    dcY = 0 : dcCb = 0 : dcCr = 0 : mcu = 0
                                End If
                            Loop Until findex >= flen Or yindex >= Image.Rows
                    End Select 'Ratio
                Case 1
                    Do
                        GetBlock(YVector1, HuffmanDC, Image.HuffDCTableY, HuffmanAC, Image.HuffACTableY, QuantTable, Image.QuantTableY, dcY)
                        'Y vector has been obtained
                        For i = 0 To 7           'Draw 8x8 pixels
                            For j = 0 To 7
                                Y = YVector1(i, j)
                                If Y < 0 Then Y = 0
                                If Y > 255 Then Y = 255
                                xj = xindex + j : yi = yindex + i
                                If xj < Image.Cols And yi < Image.Rows Then bmp.SetPixel(xj, yi, Color.FromArgb(Y, Y, Y))
                            Next j
                        Next i
                        xindex = xindex + 8
                        If xindex >= Image.Cols Then xindex = 0 : yindex = yindex + 8 : mcu = 1
                        If mcu = 1 And Restart <> 0 Then 'execute the restart interval
                            curByte = GetByte()
                            curByte = GetByte()
                            curByte = GetByte()
                            curBits = 7
                            dcY = 0 : mcu = 0
                        End If
                    Loop Until findex >= flen Or yindex >= Image.Rows
            End Select 'Components

        End If

    End Sub

    Public Sub New()

        Dim a As String
        Dim t As Single

        For X = 0 To 7           'Initialize our cosine table (used for DCT)
            For Y = 0 To 7
                For u = 0 To 7
                    For v = 0 To 7
                        t = Math.Cos((2 * X + 1) * u * 0.1963495) * Math.Cos((2 * Y + 1) * v * 0.1963495)
                        If u = 0 Then t = t * 0.707107
                        If v = 0 Then t = t * 0.707107
                        'Multiply by 16 to retain precision while staying integer
                        Dct(X, Y, u, v) = t * 16
                    Next v
                Next u
            Next Y
        Next X

        'store values for Zig-zag reordering
        a = "00011020110203122130403122130405142332415060514233241506071625344352617071625344352617273645546372736455463747566574756657677677"
        For i = 0 To 63
            Zig1(i) = Val(Mid(a, i * 2 + 1, 1))
            Zig2(i) = Val(Mid(a, i * 2 + 2, 1))
        Next i
        a = ""

    End Sub

    Private Function ReceiveBits(cat As Byte) As Integer

        Dim temp0 As Long, i As Byte

        temp0 = 0
        For i = 1 To cat
            temp0 = temp0 * 2 + NextBit()
        Next i
        If temp0 >= 2 ^ (cat - 1) Then
            ReceiveBits = temp0
        Else
            ReceiveBits = -(2 ^ cat - 1) + temp0
        End If

    End Function

    Private Sub ToRGB(y0 As Integer, cb0 As Integer, cr0 As Integer, ByRef r0 As Integer, ByRef g0 As Integer, ByRef b0 As Integer)

        'Do color space conversion from YCbCr to RGB
        r0 = y0 + cr0 * 7 / 5
        g0 = y0 - cb0 * 7 / 20 - cr0 * 18 / 25
        b0 = y0 + cb0 * 40 / 23
        If r0 > 255 Then r0 = 255
        If r0 < 0 Then r0 = 0
        If g0 > 255 Then g0 = 255
        If g0 < 0 Then g0 = 0
        If b0 > 255 Then b0 = 255
        If b0 < 0 Then b0 = 0

    End Sub

End Class
