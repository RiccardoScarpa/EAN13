Imports Microsoft.VisualBasic

'Exstend the original version of
' ///////////////          EAN13 Barcode Component            \\\\\\\\\\\\\\\ '
' ///////////////             By: Tammam Koujan               \\\\\\\\\\\\\\\ '
' /////////////// http://www.enashir.com/blogs/TammamKoujan   \\\\\\\\\\\\\\\ '
' ///////////////          email:tammamkoujan@gmail.com       \\\\\\\\\\\\\\\ '
' ///////////////     UAE - Dubai , Mobile : 050 6772279      \\\\\\\\\\\\\\\ '

Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Text
Imports System.IO

Public Class ClsBarCodeEAN13

#Region "Consts"
    Const StartMark As String = "101"
    Const SplittingMark As String = "01010"
    Const EndMark As String = "101"

#End Region

#Region "Declarations"
    Private Structure Ean13Tables
        Public TableA As String
        Public TableB As String
        Public TableC As String
    End Structure
    Private Tables(0 To 9) As Ean13Tables
    Private BarcodeValue As StringBuilder
#End Region
    Public Property BarWidth As Double
    Public Property BarHeight As Double
    Public Property BarcodeText As String
    Public Property ShowBarcodeText As Boolean 'Overrides
    Public Property ShowCheckSum() As Boolean 'Overrides

    Private m_CheckSum As Byte
    Public ReadOnly Property CheckSum() As Byte
        Get
            CalculateCheckSum()
            Return m_CheckSum
        End Get
    End Property

    Public Property Font As Font

    Public Sub New()

        InitBarcode()
        Me.BarWidth = 0.33 ' mm
        Me.Font = New Font("Arial", 18)
        InitBarcode()
        InitEAN13Tables()
    End Sub

#Region "Init Procedures"

    Private Sub InitBarcode()
        BarcodeText = "000000000000"
    End Sub

    Public Sub InitEAN13Tables()
        '          Zero
        Tables(0).TableA = "0001101"
        Tables(0).TableB = "0100111"
        Tables(0).TableC = "1110010"
        '          One
        Tables(1).TableA = "0011001"
        Tables(1).TableB = "0110011"
        Tables(1).TableC = "1100110"
        '          Two
        Tables(2).TableA = "0010011"
        Tables(2).TableB = "0011011"
        Tables(2).TableC = "1101100"
        '          Three
        Tables(3).TableA = "0111101"
        Tables(3).TableB = "0100001"
        Tables(3).TableC = "1000010"
        '          Four
        Tables(4).TableA = "0100011"
        Tables(4).TableB = "0011101"
        Tables(4).TableC = "1011100"
        '          Five
        Tables(5).TableA = "0110001"
        Tables(5).TableB = "0111001"
        Tables(5).TableC = "1001110"
        '          Six
        Tables(6).TableA = "0101111"
        Tables(6).TableB = "0000101"
        Tables(6).TableC = "1010000"
        '          Seven
        Tables(7).TableA = "0111011"
        Tables(7).TableB = "0010001"
        Tables(7).TableC = "1000100"
        '          Eight
        Tables(8).TableA = "0110111"
        Tables(8).TableB = "0001001"
        Tables(8).TableC = "1001000"
        '          Nine
        Tables(9).TableA = "0001011"
        Tables(9).TableB = "0010111"
        Tables(9).TableC = "1110100"

    End Sub

#End Region

    Private Function CalculateCheckSum() As Boolean
        Dim X As Integer = 0
        Dim Y As Integer = 0
        Dim j As Integer = 11
        Try
            For i As Integer = 1 To 12
                If i Mod 2 = 0 Then
                    X += Val(BarcodeText(j))
                Else
                    Y += Val(BarcodeText(j))
                End If
                j -= 1
            Next

            Dim Z As Integer = X + (3 * Y)
            m_CheckSum = ((10 - (Z Mod 10)) Mod 10)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function CalculateValue() As Boolean
        ' Clear any previous Value
        BarcodeValue = New StringBuilder(95)
        Try
            ' Add The Start Mark
            BarcodeValue.Append(StartMark)
            Select Case BarcodeText(0)
                Case "0"
                    For i As Integer = 1 To 6
                        BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                    Next
                Case "1"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 2) Or (i = 4) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
                Case "2"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 2) Or (i = 5) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
                Case "3"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 2) Or (i = 6) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
                Case "4"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 3) Or (i = 4) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
                Case "5"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 4) Or (i = 5) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
                Case "6"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 5) Or (i = 6) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
                Case "7"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 3) Or (i = 5) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
                Case "8"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 3) Or (i = 6) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
                Case "9"
                    For i As Integer = 1 To 6
                        If (i = 1) Or (i = 4) Or (i = 6) Then
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableA)
                        Else
                            BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableB)
                        End If
                    Next
            End Select
            ' Add The Splitting Mark
            BarcodeValue.Append(SplittingMark)

            For i As Integer = 7 To (BarcodeText.Length - 1)
                BarcodeValue.Append(Tables(Val(BarcodeText(i))).TableC)
            Next
            ' Add Checksum
            BarcodeValue.Append(Tables(CheckSum).TableC)
            ' Add The End Mark
            BarcodeValue.Append(EndMark)

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function


    Private Sub DrawBarcodeText(ByVal e As Graphics)


        ' Create font and brush. 
        Dim drawFont = Me.Font
        ' Create rectangle for drawing. 
        Dim x As Single = 3.61
        Dim y As Single = (30 + (5 * 0.33F)) '31.4F

        ' Create string to draw. 
        Dim drawString As [String] = BarcodeText
        If ShowCheckSum = True Then
            drawString += CheckSum.ToString()
            x -= 1.2F
        End If

        ' Measure string. 
        Dim stringSize As New SizeF
        stringSize = e.MeasureString(drawString, drawFont)

        Dim drawRect As New RectangleF(x, y, stringSize.Width, stringSize.Height)

        ' Set format of string. 
        Dim drawFormat As New StringFormat
        drawFormat.Alignment = StringAlignment.Center


        ' Draw string to screen.
        e.DrawString(drawString, drawFont, Brushes.Black, drawRect, drawFormat)
    End Sub

    Private Sub EAN13BarcodeDraw(e As Graphics)
        CalculateValue()
        ' Change the page scale.  
        e.PageUnit = GraphicsUnit.Millimeter
        Dim s As Single = 3
        For i As Integer = 0 To 94
            If BarcodeValue(i) = "1" Then
                Select Case i
                    Case 0, 1, 2, 45, 46, 47, 48, 49, 92, 93, 94
                        e.FillRectangle(Brushes.Black, s + 0.11F, 10, 0.5F, (20 + (5 * 0.33F)))
                    Case Else
                        e.FillRectangle(Brushes.Black, s + 0.11F, 10, 0.5F, 20)
                End Select
            ElseIf BarcodeValue(i) = "0" Then
                Select Case i
                    Case 0, 1, 2, 45, 46, 47, 48, 49, 92, 93, 94
                        e.FillRectangle(Brushes.White, s + 0.11F, 10, 0.5F, (20 + (5 * 0.33F)))
                    Case Else
                        e.FillRectangle(Brushes.White, s + 0.11F, 10, 0.5F, 20)
                End Select
            End If

            s += 0.5F
        Next

        If ShowBarcodeText = True Then
            DrawBarcodeText(e)
        End If

    End Sub

    ''' <summary>
    ''' Riccardo Scarpa
    ''' Print to a memorystream
    ''' </summary>

    Public Function GetBMPStream() As MemoryStream
        Dim lx As Integer = 200
        Dim ly As Integer = 190

        Dim gr As Graphics
        Dim bmp1 As New Bitmap(lx, ly, Imaging.PixelFormat.Format24bppRgb)
        gr = Graphics.FromImage(bmp1)
        Dim hdc As IntPtr = gr.GetHdc
        ' Make the Metafile, using the reference hDC.
        Dim mf As System.Drawing.Imaging.Metafile
        Dim bounds As New RectangleF(0, 0, lx, ly)

        Dim fstr As New IO.MemoryStream()
        mf = New System.Drawing.Imaging.Metafile(fstr, hdc, bounds, System.Drawing.Imaging.MetafileFrameUnit.Pixel)
        gr.ReleaseHdc(hdc)
        gr.Dispose()

        gr = Graphics.FromImage(mf)
        gr.PageUnit = GraphicsUnit.Millimeter
        gr.Clear(Color.White)
        gr.FillRectangle(Brushes.White, 0, 0, lx, ly)

        Call EAN13BarcodeDraw(gr)

        gr.Dispose()
        mf.Dispose()
        bmp1.Dispose()

        fstr.Seek(0, IO.SeekOrigin.Begin)

        Return fstr
    End Function

End Class







