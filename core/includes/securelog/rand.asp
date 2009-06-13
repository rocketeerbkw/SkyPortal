<%
'Nanosoft India
'Product: IAO Login 2004 ES
'Date: 1 May 2004
'Author: Nanosoft
'Copyright: Nanosoft 2004
'File: rand.asp
'Description: Genrates Random Numbers
' Modified for MWP.info by Hawk92

            Function RandomPW(myLength)
            'These constant are the minimum and maximum length for random
            'length passwords.  Adjust these values to your needs.
            Const minLength = 4
            Const maxLength = 20
    
            Dim X, Y, strPW
    
            If myLength = 0 Then
                Randomize
                myLength = Int((maxLength * Rnd) + minLength)
            End If
    
    
            For X = 1 To myLength
                'Randomize the type of this character
                Y = Int((3 * Rnd) + 1) '(1) Numeric 1-9, (2) Uppercase A-N,(3) Uppercase P-Z
    
                Select Case Y
                    Case 1
                        'Numeric character (zero omitted for readability
                        Randomize
                        strPW = strPW & CHR(Int((8 * Rnd) + 49))
                    Case 2
                        'Uppercase character A-N
                        Randomize
                        strPW = strPW & Chr(Int((13 * Rnd) + 65))
                    Case 3
                        'UpperCase character P-Z
                        Randomize
                        strPW = strPW & CHR(Int((10 * Rnd) + 80))    
                End Select
            Next
    
            RandomPW = strPW
    
    End Function
        
%>
