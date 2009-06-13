<%
' ***************************************************************
' *************** Standard font pack for ASPCanvas **************
' ***************************************************************
'
'This is the standard font pack for rendering text using ASPCanvas
'
'               Chris Read (aka Centurix/askdaquack)
'
'                            6/1/2002
'
' ***************************************************************
'
' Standard font pack for ASPCanvas
'
' ***************************************************************

' Definitions of chars 32-126

' Font and Letter must be defined to work correctly
Dim Font
Dim VFont ' Used for the vector fonts lower down
Dim Letter(5)
Dim VLetter() ' User for the vector fonts lower down

Set Font = Server.CreateObject("Scripting.Dictionary")

' Letter layouts in a 5x5 grid, this can be expanded
' As long as ALL the letters have the same width and height
' Colour fonts are supported, but only 9 colours 1-9
' the font colour for this set is colour index 1
' the background colour (0) is ignored
Letter(0) = "00000"
Letter(1) = "00000"
Letter(2) = "00000"
Letter(3) = "00000"
Letter(4) = "00000"

Font.Add " ",Letter

Letter(0) = "01000"
Letter(1) = "01000"
Letter(2) = "01000"
Letter(3) = "00000"
Letter(4) = "01000"

Font.Add "!",Letter

Letter(0) = "10100"
Letter(1) = "10100"
Letter(2) = "00000"
Letter(3) = "00000"
Letter(4) = "00000"

Font.Add Chr(34),Letter

Letter(0) = "01010"
Letter(1) = "11111"
Letter(2) = "01010"
Letter(3) = "11111"
Letter(4) = "01010"

Font.Add "#",Letter

Letter(0) = "01111"
Letter(1) = "10100"
Letter(2) = "01110"
Letter(3) = "00101"
Letter(4) = "11110"

Font.Add "$",Letter

Letter(0) = "10001"
Letter(1) = "10010"
Letter(2) = "00100"
Letter(3) = "01001"
Letter(4) = "10001"

Font.Add "%",Letter

Letter(0) = "01100"
Letter(1) = "10010"
Letter(2) = "01101"
Letter(3) = "10010"
Letter(4) = "01101"

Font.Add "&",Letter

Letter(0) = "10000"
Letter(1) = "10000"
Letter(2) = "00000"
Letter(3) = "00000"
Letter(4) = "00000"

Font.Add "'",Letter

Letter(0) = "01000"
Letter(1) = "10000"
Letter(2) = "10000"
Letter(3) = "10000"
Letter(4) = "01000"

Font.Add "(",Letter

Letter(0) = "10000"
Letter(1) = "01000"
Letter(2) = "01000"
Letter(3) = "01000"
Letter(4) = "10000"

Font.Add ")",Letter

Letter(0) = "00100"
Letter(1) = "11111"
Letter(2) = "01110"
Letter(3) = "10001"
Letter(4) = "00000"

Font.Add "*",Letter

Letter(0) = "00000"
Letter(1) = "01000"
Letter(2) = "11100"
Letter(3) = "01000"
Letter(4) = "00000"

Font.Add "+",Letter

Letter(0) = "00000"
Letter(1) = "00000"
Letter(2) = "00000"
Letter(3) = "01000"
Letter(4) = "10000"

Font.Add ",",Letter

Letter(0) = "00000"
Letter(1) = "00000"
Letter(2) = "11100"
Letter(3) = "00000"
Letter(4) = "00000"

Font.Add "-",Letter

Letter(0) = "00000"
Letter(1) = "00000"
Letter(2) = "00000"
Letter(3) = "00000"
Letter(4) = "10000"

Font.Add ".",Letter

Letter(0) = "00001"
Letter(1) = "00010"
Letter(2) = "00100"
Letter(3) = "01000"
Letter(4) = "10000"

Font.Add "/",Letter

Letter(0) = "01110"
Letter(1) = "10011"
Letter(2) = "10101"
Letter(3) = "11001"
Letter(4) = "01110"

Font.Add "0",Letter

Letter(0) = "01000"
Letter(1) = "11000"
Letter(2) = "01000"
Letter(3) = "01000"
Letter(4) = "11100"

Font.Add "1",Letter

Letter(0) = "01100"
Letter(1) = "10010"
Letter(2) = "00110"
Letter(3) = "01000"
Letter(4) = "11110"

Font.Add "2",Letter

Letter(0) = "01100"
Letter(1) = "10010"
Letter(2) = "00100"
Letter(3) = "10010"
Letter(4) = "01100"

Font.Add "3",Letter

Letter(0) = "00100"
Letter(1) = "01100"
Letter(2) = "10100"
Letter(3) = "11110"
Letter(4) = "00100"

Font.Add "4",Letter

Letter(0) = "11110"
Letter(1) = "10000"
Letter(2) = "11100"
Letter(3) = "00010"
Letter(4) = "11100"

Font.Add "5",Letter

Letter(0) = "01110"
Letter(1) = "10000"
Letter(2) = "11100"
Letter(3) = "10010"
Letter(4) = "01100"

Font.Add "6",Letter

Letter(0) = "11110"
Letter(1) = "00010"
Letter(2) = "00100"
Letter(3) = "01000"
Letter(4) = "01000"

Font.Add "7",Letter

Letter(0) = "01110"
Letter(1) = "10001"
Letter(2) = "01110"
Letter(3) = "10001"
Letter(4) = "01110"

Font.Add "8",Letter

Letter(0) = "01110"
Letter(1) = "10010"
Letter(2) = "01110"
Letter(3) = "00010"
Letter(4) = "01100"

Font.Add "9",Letter

Letter(0) = "00000"
Letter(1) = "10000"
Letter(2) = "00000"
Letter(3) = "10000"
Letter(4) = "00000"

Font.Add ":",Letter

Letter(0) = "00000"
Letter(1) = "01000"
Letter(2) = "00000"
Letter(3) = "01000"
Letter(4) = "10000"

Font.Add ";",Letter

Letter(0) = "00011"
Letter(1) = "01100"
Letter(2) = "10000"
Letter(3) = "01100"
Letter(4) = "00011"

Font.Add "<",Letter

Letter(0) = "00000"
Letter(1) = "01110"
Letter(2) = "00000"
Letter(3) = "01110"
Letter(4) = "00000"

Font.Add "=",Letter

Letter(0) = "11000"
Letter(1) = "00110"
Letter(2) = "00001"
Letter(3) = "00110"
Letter(4) = "11000"

Font.Add ">",Letter

Letter(0) = "01110"
Letter(1) = "10001"
Letter(2) = "00110"
Letter(3) = "00000"
Letter(4) = "00100"

Font.Add "?",Letter

Letter(0) = "01110"
Letter(1) = "10101"
Letter(2) = "10111"
Letter(3) = "10000"
Letter(4) = "01111"

Font.Add "@",Letter

Letter(0) = "01100"
Letter(1) = "10010"
Letter(2) = "10010"
Letter(3) = "11110"
Letter(4) = "10010"

Font.Add "A",Letter

Letter(0) = "11100"
Letter(1) = "10010"
Letter(2) = "11100"
Letter(3) = "10010"
Letter(4) = "11100"

Font.Add "B",Letter

Letter(0) = "01100"
Letter(1) = "10010"
Letter(2) = "10000"
Letter(3) = "10010"
Letter(4) = "01100"

Font.Add "C",Letter

Letter(0) = "11100"
Letter(1) = "10010"
Letter(2) = "10010"
Letter(3) = "10010"
Letter(4) = "11100"

Font.Add "D",Letter

Letter(0) = "11110"
Letter(1) = "10000"
Letter(2) = "11100"
Letter(3) = "10000"
Letter(4) = "11110"

Font.Add "E",Letter

Letter(0) = "11110"
Letter(1) = "10000"
Letter(2) = "11100"
Letter(3) = "10000"
Letter(4) = "10000"

Font.Add "F",Letter

Letter(0) = "01110"
Letter(1) = "10000"
Letter(2) = "10110"
Letter(3) = "10010"
Letter(4) = "01110"

Font.Add "G",Letter

Letter(0) = "10010"
Letter(1) = "10010"
Letter(2) = "11110"
Letter(3) = "10010"
Letter(4) = "10010"

Font.Add "H",Letter

Letter(0) = "11100"
Letter(1) = "01000"
Letter(2) = "01000"
Letter(3) = "01000"
Letter(4) = "11100"

Font.Add "I",Letter

Letter(0) = "11110"
Letter(1) = "00010"
Letter(2) = "00010"
Letter(3) = "10010"
Letter(4) = "01100"

Font.Add "J",Letter

Letter(0) = "10010"
Letter(1) = "10100"
Letter(2) = "11000"
Letter(3) = "10100"
Letter(4) = "10010"

Font.Add "K",Letter

Letter(0) = "10000"
Letter(1) = "10000"
Letter(2) = "10000"
Letter(3) = "10000"
Letter(4) = "11110"

Font.Add "L",Letter

Letter(0) = "10001"
Letter(1) = "11011"
Letter(2) = "10101"
Letter(3) = "10001"
Letter(4) = "10001"

Font.Add "M",Letter

Letter(0) = "10001"
Letter(1) = "11001"
Letter(2) = "10101"
Letter(3) = "10011"
Letter(4) = "10001"

Font.Add "N",Letter

Letter(0) = "01100"
Letter(1) = "10010"
Letter(2) = "10010"
Letter(3) = "10010"
Letter(4) = "01100"

Font.Add "O",Letter

Letter(0) = "11100"
Letter(1) = "10010"
Letter(2) = "11100"
Letter(3) = "10000"
Letter(4) = "10000"

Font.Add "P",Letter

Letter(0) = "01100"
Letter(1) = "10010"
Letter(2) = "10010"
Letter(3) = "10110"
Letter(4) = "01110"

Font.Add "Q",Letter

Letter(0) = "11100"
Letter(1) = "10010"
Letter(2) = "11100"
Letter(3) = "10100"
Letter(4) = "10010"

Font.Add "R",Letter

Letter(0) = "01110"
Letter(1) = "10000"
Letter(2) = "01100"
Letter(3) = "00010"
Letter(4) = "11100"

Font.Add "S",Letter

Letter(0) = "11100"
Letter(1) = "01000"
Letter(2) = "01000"
Letter(3) = "01000"
Letter(4) = "01000"

Font.Add "T",Letter

Letter(0) = "10010"
Letter(1) = "10010"
Letter(2) = "10010"
Letter(3) = "10010"
Letter(4) = "01100"

Font.Add "U",Letter

Letter(0) = "10001"
Letter(1) = "10001"
Letter(2) = "01010"
Letter(3) = "01010"
Letter(4) = "00100"

Font.Add "V",Letter

Letter(0) = "10001"
Letter(1) = "10001"
Letter(2) = "10101"
Letter(3) = "10101"
Letter(4) = "01010"

Font.Add "W",Letter

Letter(0) = "10001"
Letter(1) = "01010"
Letter(2) = "00100"
Letter(3) = "01010"
Letter(4) = "10001"

Font.Add "X",Letter

Letter(0) = "10001"
Letter(1) = "01010"
Letter(2) = "00100"
Letter(3) = "00100"
Letter(4) = "00100"

Font.Add "Y",Letter

Letter(0) = "11110"
Letter(1) = "00100"
Letter(2) = "01000"
Letter(3) = "10000"
Letter(4) = "11110"

Font.Add "Z",Letter

Letter(0) = "11000"
Letter(1) = "10000"
Letter(2) = "10000"
Letter(3) = "10000"
Letter(4) = "11000"

Font.Add "[",Letter

Letter(0) = "10000"
Letter(1) = "01000"
Letter(2) = "00100"
Letter(3) = "00010"
Letter(4) = "00001"

Font.Add "\",Letter

Letter(0) = "11000"
Letter(1) = "01000"
Letter(2) = "01000"
Letter(3) = "01000"
Letter(4) = "11000"

Font.Add "]",Letter

Letter(0) = "01000"
Letter(1) = "10100"
Letter(2) = "10100"
Letter(3) = "00000"
Letter(4) = "00000"

Font.Add "^",Letter

Letter(0) = "00000"
Letter(1) = "00000"
Letter(2) = "00000"
Letter(3) = "00000"
Letter(4) = "11111"

Font.Add "_",Letter

Letter(0) = "10000"
Letter(1) = "01000"
Letter(2) = "00000"
Letter(3) = "00000"
Letter(4) = "00000"

Font.Add "`",Letter

' a - z
Letter(0) = "01100"
Letter(1) = "00010"
Letter(2) = "01110"
Letter(3) = "10010"
Letter(4) = "01110"

Font.Add "a",Letter

Letter(0) = "10000"
Letter(1) = "10000"
Letter(2) = "11100"
Letter(3) = "10010"
Letter(4) = "11100"

Font.Add "b",Letter

Letter(0) = "00000"
Letter(1) = "01100"
Letter(2) = "10000"
Letter(3) = "10000"
Letter(4) = "01100"

Font.Add "c",Letter

Letter(0) = "00010"
Letter(1) = "00010"
Letter(2) = "01110"
Letter(3) = "10010"
Letter(4) = "01110"

Font.Add "d",Letter

Letter(0) = "01110"
Letter(1) = "10001"
Letter(2) = "11110"
Letter(3) = "10000"
Letter(4) = "01110"

Font.Add "e",Letter

Letter(0) = "01110"
Letter(1) = "10000"
Letter(2) = "11100"
Letter(3) = "10000"
Letter(4) = "10000"

Font.Add "f",Letter

Letter(0) = "01110"
Letter(1) = "10010"
Letter(2) = "01110"
Letter(3) = "00010"
Letter(4) = "01110"

Font.Add "g",Letter

Letter(0) = "10000"
Letter(1) = "10000"
Letter(2) = "11100"
Letter(3) = "10010"
Letter(4) = "10010"

Font.Add "h",Letter

Letter(0) = "01000"
Letter(1) = "00000"
Letter(2) = "01000"
Letter(3) = "01000"
Letter(4) = "01000"

Font.Add "i",Letter

Letter(0) = "00100"
Letter(1) = "00000"
Letter(2) = "00100"
Letter(3) = "00100"
Letter(4) = "11000"

Font.Add "j",Letter

Letter(0) = "00000"
Letter(1) = "10100"
Letter(2) = "11000"
Letter(3) = "10100"
Letter(4) = "10010"

Font.Add "k",Letter

Letter(0) = "10000"
Letter(1) = "10000"
Letter(2) = "10000"
Letter(3) = "10000"
Letter(4) = "10000"

Font.Add "l",Letter

Letter(0) = "00000"
Letter(1) = "01010"
Letter(2) = "10101"
Letter(3) = "10101"
Letter(4) = "10101"

Font.Add "m",Letter

Letter(0) = "00000"
Letter(1) = "11100"
Letter(2) = "10010"
Letter(3) = "10010"
Letter(4) = "10010"

Font.Add "n",Letter

Letter(0) = "00000"
Letter(1) = "01100"
Letter(2) = "10010"
Letter(3) = "10010"
Letter(4) = "01100"

Font.Add "o",Letter

Letter(0) = "00000"
Letter(1) = "11100"
Letter(2) = "10010"
Letter(3) = "11100"
Letter(4) = "10000"

Font.Add "p",Letter

Letter(0) = "00000"
Letter(1) = "01110"
Letter(2) = "10010"
Letter(3) = "01110"
Letter(4) = "00011"

Font.Add "q",Letter

Letter(0) = "00000"
Letter(1) = "01100"
Letter(2) = "10000"
Letter(3) = "10000"
Letter(4) = "10000"

Font.Add "r",Letter

Letter(0) = "00000"
Letter(1) = "01110"
Letter(2) = "11100"
Letter(3) = "00010"
Letter(4) = "11100"

Font.Add "s",Letter

Letter(0) = "10000"
Letter(1) = "11000"
Letter(2) = "10000"
Letter(3) = "10000"
Letter(4) = "01100"

Font.Add "t",Letter

Letter(0) = "00000"
Letter(1) = "10010"
Letter(2) = "10010"
Letter(3) = "10010"
Letter(4) = "01100"

Font.Add "u",Letter

Letter(0) = "00000"
Letter(1) = "10001"
Letter(2) = "10001"
Letter(3) = "01010"
Letter(4) = "00100"

Font.Add "v",Letter

Letter(0) = "00000"
Letter(1) = "10001"
Letter(2) = "10101"
Letter(3) = "10101"
Letter(4) = "01010"

Font.Add "w",Letter

Letter(0) = "00000"
Letter(1) = "10010"
Letter(2) = "01100"
Letter(3) = "01100"
Letter(4) = "10010"

Font.Add "x",Letter

Letter(0) = "10010"
Letter(1) = "10010"
Letter(2) = "01110"
Letter(3) = "00010"
Letter(4) = "01100"

Font.Add "y",Letter

Letter(0) = "00000"
Letter(1) = "11110"
Letter(2) = "00100"
Letter(3) = "01000"
Letter(4) = "11110"

Font.Add "z",Letter

Letter(0) = "01100"
Letter(1) = "01000"
Letter(2) = "10000"
Letter(3) = "01000"
Letter(4) = "01100"

Font.Add "{",Letter

Letter(0) = "10000"
Letter(1) = "10000"
Letter(2) = "10000"
Letter(3) = "10000"
Letter(4) = "10000"

Font.Add "|",Letter

Letter(0) = "11000"
Letter(1) = "01000"
Letter(2) = "00100"
Letter(3) = "01000"
Letter(4) = "11000"

Font.Add "}",Letter

Letter(0) = "00000"
Letter(1) = "01010"
Letter(2) = "10100"
Letter(3) = "00000"
Letter(4) = "00000"

Font.Add "~",Letter

' Vector font support, intended to supercede the above bitmap font support
' The format is:
' VFont(Line,X/Y/Pen down) = Number
' Line = Line number to draw, lines are drawn in this order!
' X = 0, X position for point
' Y = 1, Y position for point
' Pen down = Line is actually drawn from last point to this point
'
' All fonts are drawn with the current foreground colour, which is different from the
' bitmapped fonts where the colour information is stored in the font arrays.
' 
' All characters must have at LEAST one point, space is a special character and is defined
' here by a single vertical line.
' 
' If you make a better vector font pack than this, I'd be interested in including it with
' ASPCanvas. Mail it to: mrjolly@bigpond.net.au
'
' In particular, if you come up with an OCR-A (or even a nice OCR-B) font, fixed width 
' courier would be nice too...
'
' The font set below is based roughly on the bitmap font from above when scaled to 1
' As of 1.0.3, the vector fonts are a bit innacurate...

Set VFont = Server.CreateObject("Scripting.Dictionary")

ReDim VLetter(4,3)

VLetter(0,0) = 2 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 0
VLetter(3,0) = 2 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add "!",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 1 : VLetter(2,2) = 0
VLetter(3,0) = 3 : VLetter(3,1) = 2 : VLetter(3,2) = 1

VFont.Add Chr(34),VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 2 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 1 : VLetter(2,2) = 0
VLetter(3,0) = 4 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 2 : VLetter(4,2) = 0
VLetter(5,0) = 5 : VLetter(5,1) = 2 : VLetter(5,2) = 1
VLetter(6,0) = 1 : VLetter(6,1) = 4 : VLetter(6,2) = 0
VLetter(7,0) = 5 : VLetter(7,1) = 4 : VLetter(7,2) = 1

VFont.Add "#",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 5 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 3 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 5 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 5 : VLetter(5,2) = 1
VLetter(6,0) = 1 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 3 : VLetter(7,1) = 1 : VLetter(7,2) = 0
VLetter(8,0) = 3 : VLetter(8,1) = 5 : VLetter(8,2) = 1

VFont.Add "$",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 5 : VLetter(2,1) = 1 : VLetter(2,2) = 0
VLetter(3,0) = 1 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 5 : VLetter(4,1) = 5 : VLetter(4,2) = 0
VLetter(5,0) = 5 : VLetter(5,1) = 4 : VLetter(5,2) = 1

VFont.Add "%",VLetter

ReDim VLetter(13,3)

VLetter(0,0) = 5 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 3 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 1 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 1 : VLetter(5,2) = 1
VLetter(6,0) = 4 : VLetter(6,1) = 2 : VLetter(6,2) = 1
VLetter(7,0) = 3 : VLetter(7,1) = 3 : VLetter(7,2) = 1
VLetter(8,0) = 2 : VLetter(8,1) = 3 : VLetter(8,2) = 1
VLetter(9,0) = 1 : VLetter(9,1) = 4 : VLetter(9,2) = 1
VLetter(10,0) = 2 : VLetter(10,1) = 5 : VLetter(10,2) = 1
VLetter(11,0) = 3 : VLetter(11,1) = 5 : VLetter(11,2) = 1
VLetter(12,0) = 5 : VLetter(12,1) = 3 : VLetter(12,2) = 1

VFont.Add "&",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1

VFont.Add "'",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 2 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 2 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add "(",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add ")",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 3 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 4 : VLetter(1,2) = 1
VLetter(2,0) = 5 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 5 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 1 : VLetter(5,2) = 1

VFont.Add "*",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 2 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 2 : VLetter(2,2) = 0
VLetter(3,0) = 3 : VLetter(3,1) = 2 : VLetter(3,2) = 1

VFont.Add "+",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 2 : VLetter(0,1) = 4 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1

VFont.Add ",",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 3 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 3 : VLetter(1,2) = 1

VFont.Add "-",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1

VFont.Add ".",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 5 : VLetter(1,1) = 1 : VLetter(1,2) = 1

VFont.Add "/",VLetter

ReDim VLetter(12,3)

VLetter(0,0) = 1 : VLetter(0,1) = 4 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 5 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 5 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 4 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 2 : VLetter(7,1) = 5 : VLetter(7,2) = 1
VLetter(8,0) = 1 : VLetter(8,1) = 4 : VLetter(8,2) = 1
VLetter(9,0) = 2 : VLetter(9,1) = 4 : VLetter(9,2) = 1
VLetter(10,0) = 4 : VLetter(10,1) = 2 : VLetter(10,2) = 1
VLetter(11,0) = 4 : VLetter(11,1) = 1 : VLetter(11,2) = 1

VFont.Add "0",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 5 : VLetter(4,2) = 1

VFont.Add "1",VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 3 : VLetter(5,2) = 1
VLetter(6,0) = 1 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 5 : VLetter(7,2) = 1

VFont.Add "2",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 2 : VLetter(7,1) = 5 : VLetter(7,2) = 1
VLetter(8,0) = 1 : VLetter(8,1) = 4 : VLetter(8,2) = 1

VFont.Add "3",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 3 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 3 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 4 : VLetter(4,2) = 1

VFont.Add "4",VLetter

ReDim VLetter(7,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 1 : VLetter(5,2) = 1
VLetter(6,0) = 4 : VLetter(6,1) = 1 : VLetter(6,2) = 1

VFont.Add "5",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 1 : VLetter(0,1) = 3 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 5 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 1 : VLetter(6,1) = 2 : VLetter(6,2) = 1
VLetter(7,0) = 2 : VLetter(7,1) = 1 : VLetter(7,2) = 1
VLetter(8,0) = 4 : VLetter(8,1) = 1 : VLetter(8,2) = 1

VFont.Add "6",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 2 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 5 : VLetter(4,2) = 1

VFont.Add "7",VLetter

ReDim VLetter(12,3)

VLetter(0,0) = 2 : VLetter(0,1) = 3 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 5 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 3 : VLetter(5,2) = 1
VLetter(6,0) = 2 : VLetter(6,1) = 3 : VLetter(6,2) = 1
VLetter(7,0) = 1 : VLetter(7,1) = 4 : VLetter(7,2) = 1
VLetter(8,0) = 2 : VLetter(8,1) = 5 : VLetter(8,2) = 1
VLetter(9,0) = 4 : VLetter(9,1) = 5 : VLetter(9,2) = 1
VLetter(10,0) = 5 : VLetter(10,1) = 4 : VLetter(10,2) = 1
VLetter(11,0) = 4 : VLetter(11,1) = 3 : VLetter(11,2) = 1

VFont.Add "8",VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 2 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 1 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 2 : VLetter(5,2) = 1
VLetter(6,0) = 2 : VLetter(6,1) = 3 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 3 : VLetter(7,2) = 1

VFont.Add "9",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 4 : VLetter(2,2) = 0
VLetter(3,0) = 1 : VLetter(3,1) = 4 : VLetter(3,2) = 1

VFont.Add ":",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 2 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 4 : VLetter(2,2) = 0
VLetter(3,0) = 2 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 5 : VLetter(4,2) = 1

VFont.Add ";",VLetter

ReDim VLetter(3,3)

VLetter(0,0) = 5 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 5 : VLetter(2,1) = 5 : VLetter(2,2) = 1

VFont.Add "<",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 2 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 4 : VLetter(2,2) = 0
VLetter(3,0) = 4 : VLetter(3,1) = 4 : VLetter(3,2) = 1

VFont.Add "=",VLetter

ReDim VLetter(3,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 5 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 5 : VLetter(2,2) = 1

VFont.Add ">",VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 5 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 3 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 5 : VLetter(6,2) = 0
VLetter(7,0) = 3 : VLetter(7,1) = 5 : VLetter(7,2) = 1

VFont.Add "?",VLetter

ReDim VLetter(10,3)

VLetter(0,0) = 5 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 1 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 1 : VLetter(5,2) = 1
VLetter(6,0) = 5 : VLetter(6,1) = 2 : VLetter(6,2) = 1
VLetter(7,0) = 5 : VLetter(7,1) = 3 : VLetter(7,2) = 1
VLetter(8,0) = 3 : VLetter(8,1) = 2 : VLetter(8,2) = 1
VLetter(9,0) = 3 : VLetter(9,1) = 1 : VLetter(9,2) = 1

VFont.Add "@",VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 5 : VLetter(5,2) = 1
VLetter(6,0) = 1 : VLetter(6,1) = 4 : VLetter(6,2) = 0
VLetter(7,0) = 4 : VLetter(7,1) = 4 : VLetter(7,2) = 1

VFont.Add "A",VLetter

ReDim VLetter(10,2)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 3 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 3 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 2 : VLetter(7,2) = 1
VLetter(8,0) = 3 : VLetter(8,1) = 1 : VLetter(8,2) = 1
VLetter(9,0) = 1 : VLetter(9,1) = 1 : VLetter(9,2) = 1

VFont.Add "B",VLetter

ReDim VLetter(8,2)

VLetter(0,0) = 4 : VLetter(0,1) = 4 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 2 : VLetter(5,1) = 1 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 1 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 2 : VLetter(7,2) = 1

VFont.Add "C",VLetter

ReDim VLetter(7,2)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 1 : VLetter(5,2) = 1
VLetter(6,0) = 1 : VLetter(6,1) = 1 : VLetter(6,2) = 1

VFont.Add "D",VLetter

ReDim VLetter(8,2)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 3 : VLetter(5,2) = 1
VLetter(6,0) = 1 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 5 : VLetter(7,2) = 1

VFont.Add "E",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 3 : VLetter(3,2) = 0
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1

VFont.Add "F",VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 3 : VLetter(0,1) = 3 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 2 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 2 : VLetter(5,2) = 1
VLetter(6,0) = 2 : VLetter(6,1) = 1 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 1 : VLetter(7,2) = 1

VFont.Add "G",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 1 : VLetter(2,2) = 0
VLetter(3,0) = 4 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 3 : VLetter(4,2) = 0
VLetter(5,0) = 4 : VLetter(5,1) = 3 : VLetter(5,2) = 1

VFont.Add "H",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 5 : VLetter(2,2) = 0
VLetter(3,0) = 3 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 1 : VLetter(4,2) = 0
VLetter(5,0) = 2 : VLetter(5,1) = 5 : VLetter(5,2) = 1

VFont.Add "I",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 5 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 4 : VLetter(5,2) = 1

VFont.Add "J",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 2 : VLetter(4,2) = 0
VLetter(5,0) = 4 : VLetter(5,1) = 5 : VLetter(5,2) = 1

VFont.Add "K",VLetter

ReDim VLetter(3,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 5 : VLetter(2,2) = 1

VFont.Add "L",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 3 : VLetter(2,2) = 1
VLetter(3,0) = 5 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 5 : VLetter(4,1) = 5 : VLetter(4,2) = 1

VFont.Add "M",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 5 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 5 : VLetter(3,1) = 1 : VLetter(3,2) = 1

VFont.Add "N",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 1 : VLetter(0,1) = 4 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 2 : VLetter(7,1) = 5 : VLetter(7,2) = 1
VLetter(8,0) = 1 : VLetter(8,1) = 4 : VLetter(8,2) = 1

VFont.Add "O",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 3 : VLetter(5,2) = 1

VFont.Add "P",VLetter

ReDim VLetter(11,3)

VLetter(0,0) = 1 : VLetter(0,1) = 4 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 2 : VLetter(7,1) = 5 : VLetter(7,2) = 1
VLetter(8,0) = 1 : VLetter(8,1) = 4 : VLetter(8,2) = 1
VLetter(9,0) = 3 : VLetter(9,1) = 4 : VLetter(9,2) = 0
VLetter(10,0) = 4 : VLetter(10,1) = 5 : VLetter(10,2) = 1

VFont.Add "Q",VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 3 : VLetter(5,2) = 1
VLetter(6,0) = 2 : VLetter(6,1) = 3 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 5 : VLetter(7,2) = 1

VFont.Add "R",VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 2 : VLetter(5,2) = 1
VLetter(6,0) = 2 : VLetter(6,1) = 1 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 1 : VLetter(7,2) = 1

VFont.Add "S",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 1 : VLetter(2,2) = 0
VLetter(3,0) = 2 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add "T",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 4 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 1 : VLetter(5,2) = 1

VFont.Add "U",VLetter

ReDim VLetter(3,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 5 : VLetter(2,1) = 1 : VLetter(2,2) = 1

VFont.Add "V",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 4 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 4 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 5 : VLetter(7,1) = 4 : VLetter(7,2) = 1
VLetter(8,0) = 5 : VLetter(8,1) = 1 : VLetter(8,2) = 1

VFont.Add "W",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 5 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 5 : VLetter(2,2) = 0
VLetter(3,0) = 5 : VLetter(3,1) = 1 : VLetter(3,2) = 1

VFont.Add "X",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 5 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 3 : VLetter(3,2) = 0
VLetter(4,0) = 3 : VLetter(4,1) = 5 : VLetter(4,2) = 1

VFont.Add "Y",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 5 : VLetter(4,2) = 1

VFont.Add "Z",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 2 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 2 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add "[",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 5 : VLetter(1,1) = 5 : VLetter(1,2) = 1

VFont.Add "\",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add "]",VLetter

ReDim VLetter(3,3)

VLetter(0,0) = 1 : VLetter(0,1) = 3 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 3 : VLetter(2,2) = 1

VFont.Add "^",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 5 : VLetter(1,1) = 5 : VLetter(1,2) = 1

VFont.Add "_",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 2 : VLetter(1,2) = 1

VFont.Add "`",VLetter

' Lower case fonts

ReDim VLetter(8,3)

VLetter(0,0) = 2 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 5 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 2 : VLetter(6,1) = 3 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 3 : VLetter(7,2) = 1

VFont.Add "a",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 3 : VLetter(5,2) = 1

VFont.Add "b",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 3 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 3 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 5 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 5 : VLetter(5,2) = 1

VFont.Add "c",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 4 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 3 : VLetter(5,2) = 1

VFont.Add "d",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 4 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 1 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 1 : VLetter(5,2) = 1
VLetter(6,0) = 5 : VLetter(6,1) = 2 : VLetter(6,2) = 1
VLetter(7,0) = 4 : VLetter(7,1) = 3 : VLetter(7,2) = 1
VLetter(8,0) = 1 : VLetter(8,1) = 3 : VLetter(8,2) = 1

VFont.Add "e",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 0
VLetter(5,0) = 1 : VLetter(5,1) = 3 : VLetter(5,2) = 1

VFont.Add "f",VLetter

ReDim VLetter(7,3)

VLetter(0,0) = 2 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 1 : VLetter(2,2) = 1
VLetter(3,0) = 2 : VLetter(3,1) = 1 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 2 : VLetter(5,1) = 3 : VLetter(5,2) = 1
VLetter(6,0) = 4 : VLetter(6,1) = 3 : VLetter(6,2) = 1

VFont.Add "g",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 3 : VLetter(2,2) = 0
VLetter(3,0) = 3 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 5 : VLetter(5,2) = 1

VFont.Add "h",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 2 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 3 : VLetter(2,2) = 0
VLetter(3,0) = 2 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add "i",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 3 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 3 : VLetter(2,2) = 0
VLetter(3,0) = 3 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 5 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 5 : VLetter(5,2) = 1

VFont.Add "j",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 5 : VLetter(2,2) = 0
VLetter(3,0) = 1 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 2 : VLetter(4,2) = 0
VLetter(5,0) = 1 : VLetter(5,1) = 4 : VLetter(5,2) = 1

VFont.Add "k",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1

VFont.Add "l",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 5 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 3 : VLetter(5,2) = 0
VLetter(6,0) = 4 : VLetter(6,1) = 2 : VLetter(6,2) = 1
VLetter(7,0) = 5 : VLetter(7,1) = 3 : VLetter(7,2) = 1
VLetter(8,0) = 5 : VLetter(8,1) = 5 : VLetter(8,2) = 1

VFont.Add "m",VLetter

ReDim VLetter(5,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 5 : VLetter(4,2) = 1

VFont.Add "n",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 1 : VLetter(0,1) = 4 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 2 : VLetter(7,1) = 5 : VLetter(7,2) = 1
VLetter(8,0) = 1 : VLetter(8,1) = 4 : VLetter(8,2) = 1

VFont.Add "o",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 1 : VLetter(5,1) = 4 : VLetter(5,2) = 1

VFont.Add "p",VLetter

ReDim VLetter(7,3)

VLetter(0,0) = 4 : VLetter(0,1) = 4 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 4 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 3 : VLetter(2,2) = 1
VLetter(3,0) = 2 : VLetter(3,1) = 2 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 2 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 5 : VLetter(5,2) = 1
VLetter(6,0) = 5 : VLetter(6,1) = 5 : VLetter(6,2) = 1

VFont.Add "q",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 3 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 2 : VLetter(3,2) = 1

VFont.Add "r",VLetter

ReDim VLetter(7,3)

VLetter(0,0) = 1 : VLetter(0,1) = 5 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 4 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 2 : VLetter(5,1) = 2 : VLetter(5,2) = 1
VLetter(6,0) = 4 : VLetter(6,1) = 2 : VLetter(6,2) = 1

VFont.Add "s",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 4 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 1 : VLetter(4,1) = 2 : VLetter(4,2) = 0
VLetter(5,0) = 2 : VLetter(5,1) = 2 : VLetter(5,2) = 1

VFont.Add "t",VLetter

ReDim VLetter(6,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 4 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 5 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 4 : VLetter(5,1) = 2 : VLetter(5,2) = 1

VFont.Add "u",VLetter

ReDim VLetter(3,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 3 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 5 : VLetter(2,1) = 2 : VLetter(2,2) = 1

VFont.Add "v",VLetter

ReDim VLetter(9,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 4 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 4 : VLetter(3,2) = 1
VLetter(4,0) = 3 : VLetter(4,1) = 3 : VLetter(4,2) = 1
VLetter(5,0) = 3 : VLetter(5,1) = 4 : VLetter(5,2) = 0
VLetter(6,0) = 4 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 5 : VLetter(7,1) = 4 : VLetter(7,2) = 1
VLetter(8,0) = 5 : VLetter(8,1) = 2 : VLetter(8,2) = 1

VFont.Add "w",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 5 : VLetter(1,2) = 1
VLetter(2,0) = 4 : VLetter(2,1) = 2 : VLetter(2,2) = 0
VLetter(3,0) = 1 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add "x",VLetter

ReDim VLetter(8,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 3 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 4 : VLetter(4,1) = 1 : VLetter(4,2) = 0
VLetter(5,0) = 4 : VLetter(5,1) = 4 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 5 : VLetter(6,2) = 1
VLetter(7,0) = 2 : VLetter(7,1) = 5 : VLetter(7,2) = 1

VFont.Add "y",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 2 : VLetter(0,2) = 1
VLetter(1,0) = 4 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 1 : VLetter(2,1) = 5 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 5 : VLetter(3,2) = 1

VFont.Add "z",VLetter

ReDim VLetter(7,3)

VLetter(0,0) = 3 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 1 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 2 : VLetter(5,1) = 5 : VLetter(5,2) = 1
VLetter(6,0) = 3 : VLetter(6,1) = 5 : VLetter(6,2) = 1

VFont.Add "{",VLetter

ReDim VLetter(2,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 1 : VLetter(1,1) = 5 : VLetter(1,2) = 1

VFont.Add "|",VLetter

ReDim VLetter(7,3)

VLetter(0,0) = 1 : VLetter(0,1) = 1 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 1 : VLetter(1,2) = 1
VLetter(2,0) = 2 : VLetter(2,1) = 2 : VLetter(2,2) = 1
VLetter(3,0) = 3 : VLetter(3,1) = 3 : VLetter(3,2) = 1
VLetter(4,0) = 2 : VLetter(4,1) = 4 : VLetter(4,2) = 1
VLetter(5,0) = 2 : VLetter(5,1) = 5 : VLetter(5,2) = 1
VLetter(6,0) = 1 : VLetter(6,1) = 5 : VLetter(6,2) = 1

VFont.Add "}",VLetter

ReDim VLetter(4,3)

VLetter(0,0) = 1 : VLetter(0,1) = 3 : VLetter(0,2) = 1
VLetter(1,0) = 2 : VLetter(1,1) = 2 : VLetter(1,2) = 1
VLetter(2,0) = 3 : VLetter(2,1) = 3 : VLetter(2,2) = 1
VLetter(3,0) = 4 : VLetter(3,1) = 2 : VLetter(3,2) = 1

VFont.Add "~",VLetter

%>