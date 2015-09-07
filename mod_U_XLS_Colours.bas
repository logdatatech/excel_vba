''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Function            Color
'   Purpose             Determine the Background Color Of a Cell
'   @Param rng          Range to Determine Background Color of
'   @Param formatType   Default Value = 0
'                       0   Integer
'                       1   Hex
'                       2   RGB
'                       3   Excel Color Index
'   Usage               Color(A1)      -->   9507341
'                       Color(A1, 0)   -->   9507341
'                       Color(A1, 1)   -->   91120D
'                       Color(A1, 2)   -->   13, 18, 145
'                       Color(A1, 3)   -->   6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Color(rng As Range, Optional formatType As Integer = 0)     As Variant
    Dim colorVal As Variant
    colorVal = Cells(rng.Row, rng.Column).Interior.Color
    Select Case formatType
        Case 1
            Color = Hex(colorVal)
        Case 2
            Color = (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & (colorVal \ 65536)
        Case 3
            Color = Cells(rng.Row, rng.Column).Interior.ColorIndex
        Case Else
            Color = colorVal
    End Select
End Function


If you have XL2007 or higher:
MsgBox "&H" & Application.WorksheetFunction.Dec2Hex(ActiveCell.Interior.Color, 6) & "&"




Sub colors56()
'57 colors, 0 to 56

Dim i As Long
Dim str0 As String, str As String
Cells(1, 1) = "Interior"
Cells(1, 2) = "Font"
Cells(1, 3) = "HTML"
Cells(1, 4) = "RED"
Cells(1, 5) = "GREEN"
Cells(1, 6) = "BLUE"
Cells(1, 7) = "COLOR"


For i = 0 To 56
    Cells(i + 2, 1).Interior.ColorIndex = i  
    Cells(i + 2, 2).Font.ColorIndex = i
    Cells(i + 2, 2).Value = "[Color " & i & "]"
    str0 = Right("000000" & Hex(Cells(i + 2, 1).Interior.Color), 6)
    'Excel shows nibbles in reverse order so make it as RGB
   str = Right(str0, 2) & Mid(str0, 3, 2) & Left(str0, 2)
    'generating 2 columns in the HTML table
   Cells(i + 2, 3) = "#" & str
    Cells(i + 2, 4).Formula = "=Hex2dec(""" & Right(str0, 2) & """)"
    Cells(i + 2, 5).Formula = "=Hex2dec(""" & Mid(str0, 3, 2) & """)"
    Cells(i + 2, 6).Formula = "=Hex2dec(""" & Left(str0, 2) & """)"
    Cells(i + 2, 7) = "[Color " & i & "]"
Next i
End Sub


Set  background:
Worksheets(“Calendar”).Range(“A1:M1”).Interior.Color = RGB(218,225,130)
Set font  color:
Worksheets(“Calendar”).Range(“A1:M1”).Font.Color = RGB(218,225,130)
- See more at: http://www.exceldigest.com/myblog/2012/09/20/how-to-use-colors-in-excel-2010-vba-code/#sthash.psArF8td.dpuf


Public Function XLSCellColour2Number (lRange As Long) As Long

	



End Function




Sub ShowColorIndex()
Dim i As Integer, j As Integer
For i = 1 To 4
For j = 1 To 14
Worksheets(“ColorIndex”).Cells(j, (i – 1) * 2 + 1).Value = (i – 1) * 14 + j
Worksheets(“ColorIndex”).Cells(j, i * 2).Interior.ColorIndex = (i – 1) * 14 + j
Next j
Next i
End Sub
'- See more at: http://www.exceldigest.com/myblog/2012/09/20/how-to-use-colors-in-excel-2010-vba-code/#sthash.psArF8td.dpuf


Function CellColor(rCell As Range, Optional ColorName As Boolean)

	Dim strColor As String, iIndexNum As Integer
	'Written by Dave Hawley of OzGrid.com

Select Case rCell.Interior.ColorIndex

   Case 1

    strColor = "Black"

    iIndexNum = 1

   Case 53

    strColor = "Brown"

    iIndexNum = 53

   Case 52

    strColor = "Olive Green"

    iIndexNum = 52

   Case 51

    strColor = "Dark Green"

    iIndexNum = 51

   Case 49

    strColor = "Dark Teal"

    iIndexNum = 49

   Case 11

    strColor = "Dark Blue"

    iIndexNum = 11

   Case 55

    strColor = "Indigo"

    iIndexNum = 55

   Case 56

    strColor = "Gray-80%"

    iIndexNum = 56

   Case 9

    strColor = "Dark Red"

    iIndexNum = 9

   Case 46

    strColor = "Orange"

    iIndexNum = 46

   Case 12

    strColor = "Dark Yellow"

    iIndexNum = 12

   Case 10

    strColor = "Green"

    iIndexNum = 10

   Case 14

    strColor = "Teal"

    iIndexNum = 14

   Case 5

    strColor = "Blue"

    iIndexNum = 5

   Case 47

    strColor = "Blue-Gray"

    iIndexNum = 47

   Case 16

    strColor = "Gray-50%"

    iIndexNum = 16

   Case 3

    strColor = "Red"

    iIndexNum = 3

   Case 45

    strColor = "Light Orange"

    iIndexNum = 45

   Case 43

    strColor = "Lime"

    iIndexNum = 43

   Case 50

    strColor = "Sea Green"

    iIndexNum = 50

   Case 42

    strColor = "Aqua"

    iIndexNum = 42

   Case 41

    strColor = "Light Blue"

    iIndexNum = 41

   Case 13

    strColor = "Violet"

    iIndexNum = 13

   Case 48

    strColor = "Gray-40%"

    iIndexNum = 48

   Case 7

    strColor = "Pink"

    iIndexNum = 7

   Case 44

    strColor = "Gold"

    iIndexNum = 44

   Case 6

    strColor = "Yellow"

    iIndexNum = 6

   Case 4

    strColor = "Bright Green"

    iIndexNum = 4

   Case 8

    strColor = "Turqoise"

    iIndexNum = 8

   Case 33

    strColor = "Sky Blue"

    iIndexNum = 33

   Case 54

    strColor = "Plum"

    iIndexNum = 54

   Case 15

    strColor = "Gray-25%"

    iIndexNum = 15

   Case 38

    strColor = "Rose"

    iIndexNum = 38

   Case 40

    strColor = "Tan"

    iIndexNum = 40

   Case 36

    strColor = "Light Yellow"

    iIndexNum = 36

   Case 35

    strColor = "Light Green"

    iIndexNum = 35

   Case 34

    strColor = "Light Turqoise"

    iIndexNum = 34

   Case 37

    strColor = "Pale Blue"

    iIndexNum = 37

   Case 39

    strColor = "Lavendar"

    iIndexNum = 39

   Case 2

    strColor = "White"

    iIndexNum = 2

  Case Else

    strColor = "Custom color or no fill"

End Select



	If ColorName = True Or _

		strColor = "Custom color or no fill" Then

		CellColor = strColor

	Else

		CellColor = iIndexNum

	End If



End Function
