'This module is licensed under the MIT Open Source license, (c) 2010 Neoptolemos Melonas <neo@neomelonas.com>
'You should have a copy of the license included as a file called LICENSE
'if this file is not around, the license text can be found at
'http://opensource.org/licenses/mit-license.php
'
Sub RandNumGen()
    'DECLARATIONS
    'WOOOOOOOOOOO
    Dim theFirst As Long, _
        theLast As Long, _
        theCount As Long, _
        theArray() As Variant, _
        aCount As Long, _
        theTrash As Long, _
        theCompany As String, _
        theFYE As String, _
        lastRow As String, _
        firstRow As String, _
        anotherVar As Long, _
        otherJunk As Long, _
        varOfHolding As Long, _
        checker As Boolean, _
        mine As String
        
    Application.Volatile
    
    'Make sure the output doesn't have overlap
    Worksheets("Output").Range("A7:AA400").Clear
    'Set some defaults if you forget to enter inputs.
    'Defaults company name to "Test Corporation"
    If Worksheets("Main").Range("E7").Value = vbNullString Then
        theCompany = "Test Corporation"
    Else: theCompany = Worksheets("Main").Range("E7").Value
    End If
    'Defaults FYE to 12/31 of last calendar year
    If Worksheets("Main").Range("E8").Value = vbNullString Or Not IsDate(Worksheets("Main").Range("E8").Value) Then
        theFYE = "12/31/" & Year(Date) - 1
    Else: theFYE = Worksheets("Main").Range("E8").Value
    End If
    'Defaults the Company Name in the title line to blank.
    If Worksheets("Main").Range("E12").Value = vbNullString Then
        mine = ""
    Else: mine = Worksheets("Main").Range("E12").Value & " "
    End If
    
    'Defaults to 15 picks
    If Worksheets("Main").Range("E11").Value = vbNullString Or Not IsNumeric(Worksheets("Main").Range("E11").Value) Then
        theCount = 15
    Else: theCount = Worksheets("Main").Range("E11").Value
    End If
    
    If Worksheets("Main").Populationing = True Then
        
        Dim popColCount As Long, _
            popRowCount As Long, _
            popColLetter As String
            
        popColCount = Application.CountIf(Worksheets("Population").Rows(1), "*")
        'popRowCount = Application.CountIf(Worksheets("Population").Columns(1), "*")
        popRowCount = Worksheets("Population").UsedRange.Rows.Count
        
        popColLetter = GetColumnLetter(popColCount)
        
        theFirst = 1
        theLast = popRowCount
                
    Else:
        
        'Defaults the range minimum to 1
        If Worksheets("Main").Range("E9").Value = vbNullString Or Not IsNumeric(Worksheets("Main").Range("E9").Value) Then
            theFirst = 1
        Else: theFirst = Worksheets("Main").Range("E9").Value
        End If
        'Defaults the range maximum to 100
        If Worksheets("Main").Range("E10").Value = vbNullString Or Not IsNumeric(Worksheets("Main").Range("E10").Value) Then
            theLast = 100
        Else: theLast = Worksheets("Main").Range("E10").Value
        End If
        'Checks to see if first is bigger than last. If it is, swap values, for sanity and logic's sake.
        If theLast < theFirst Then
            varOfHolding = theLast
            theLast = theFirst
            theFirst = varOfHolding
        End If
        
    End If

    If (theCount > (theLast - theFirst)) Then
        theCount = Round((theLast - theFirst) * Rnd, 0)
        Worksheets("Output").Range("E5").Value = "Don't be dumb, pick a sane value for this, or I'll just keep making one up."
    End If
    
    'Shift stuff down
    otherJunk = theCount + 7
    'It all goes down on column D, yo.
    lastRow = "D" & otherJunk
    'This is a counter... watch as it counts.
    'Majestically.
    aCount = 0
    'Reassign the var type/size to be the correct size + 1, for overflow/lazy math reasons.
    ReDim theArray(0 To theCount)

    'We preset this to true, so the super-inner-doing-stuff while loop works.
    'This would be false, because that's what actually happens, but VBA sucks at
    'checking if things are false. Shut up, it made sense when this was written.
    checker = True
    'This is where the numbers get picked
    For j = 0 To theCount
    
        'Set the first random number generated to the array, because if it's the first run,
        'there is no way that the thing can be a repeat.
        If aCount = 0 Then
            theTrash = Round((Rnd * (theLast - theFirst)) + theFirst, 0)
            theArray(j) = theTrash
        End If
        
        'This loop checks theArray to see if the random number is a duplicate
        'PS: this is the super-inner-doing-stuff while loop.
        Do While checker = True And (j > 0)
            'If it's false, the loop ends, meaning things worked
            'We reset it here to false so that if things work, the loop actually ends,
            'otherwise, we get eternal-loop attacked.
            checker = False
            'GENERATE A RANDOM NUMBER WOO
            theTrash = Round((Rnd * (theLast - theFirst)) + theFirst, 0)
            'The actual purpose of this giant loop structure is seriously ALL within this tiny for
            'loop. I just wanted to make the code look enormous.
            'Search through theArray, if anything matches, sets checker to true, which in turn will
            're-run the entire process of generating a random number.
            For I = 0 To UBound(theArray)
                If theTrash = theArray(I) Then checker = True
            Next
            'This will add the random number to theArray(j).
            If checker = False Then
                theArray(j) = theTrash
            End If
        Loop
        'Reset checker, because we always have to run that inner for-loop.
        checker = True
    Next j
     
    'Just some dumb, but necessary placeholder variables.
    'Never mind their names, they get used to point out cell refs in the oncoming loop.
    Dim junk As Long, _
        trash As Long, _
        countaz As Long
        
    'This var serves as the index for theArray in the coming loop.
    anotherVar = 0
    'this is where they get displayed
    
    'Formatting, because some of the defaults for the datatypes disagree with what we want.
    Worksheets("Output").Range("A2:D5").Font.Bold = True
    Worksheets("Output").Range("A2").Font.Size = 14
    Worksheets("Output").Range("D2").HorizontalAlignment = xlLeft
    'Flows input data onto the "output" page
    Worksheets("Output").Range("A2") = theCompany
    Worksheets("Output").Range("E2") = theFYE
    Worksheets("Output").Range("D3") = theFirst
    Worksheets("Output").Range("D4") = theLast
    Worksheets("Output").Range("D5") = theCount
    Worksheets("Output").Range("A1") = mine & "Random Number Generator"
    
    Do While anotherVar < theCount
        'Junk is the cell location on "Output", trash is the counter, as well as
        'the reference point for where in the sequence the number was generated
        '(To mirror the functionality of the old Lotus 1-2-3 script)
        junk = anotherVar + 7
        trash = anotherVar + 1
        'Display, in order:
        'What the sample number is, from 1 to theCount
        'When, in the sequence, the number was selected (also from 1 to theCount, but less obvious)
        'The randomly generated, non-repeating Long from within the range theFirst to theLast
        Worksheets("Output").Cells(junk, 1) = trash
        Worksheets("Output").Cells(junk, 2) = trash
        Worksheets("Output").Cells(junk, 3) = theArray(anotherVar)
        countaz = 0
        If Worksheets("Main").Populationing.Value = True Then
        'Add column titles & attributes
            Do While countaz < popColCount
                Worksheets("Output").Cells(6, countaz + 4) = Worksheets("Population").Cells(1, countaz + 1)
                Worksheets("Output").Cells(junk, 4 + countaz) = Worksheets("Population").Cells(theArray(anotherVar) + 1, countaz + 1)
                countaz = countaz + 1
            Loop
        End If
        'Increment that beast.
        anotherVar = anotherVar + 1
    Loop
    Dim attributes As Long
    attributes = 4
    Do While attributes > 0
        Worksheets("Output").Cells(6, popColCount + 5 + attributes) = attributes
        attributes = attributes - 1
    Loop
    'This sorts the generated random numbers from smallest to largest, takes into consideration
    'which method of sampling was used.
    If Worksheets("Main").Populationing = True Then
        Dim lastPopSampleRow As String, _
            colCountTrue As Long, _
            colLetterTrue As String
        colCountTrue = popColCount + 3
        colLetterTrue = GetColumnLetter(colCountTrue)
        
        lastPopSampleRow = colLetterTrue & otherJunk
        Worksheets("Output").Range("B7:" & lastPopSampleRow).Sort Key1:=Worksheets("Output").Columns("C")
    Else:
        Worksheets("Output").Range("B7:" & lastRow).Sort Key1:=Worksheets("Output").Columns("C")
    End If
    'Switch active sheets to the output, on clicking the "Generate!" button, after the generator is done.
    Sheets("Output").Select
End Sub

Function GetColumnLetter(ColumnNumber As Long) As String
    If ColumnNumber < 26 Then
        ' Columns A-Z
        GetColumnLetter = Chr(ColumnNumber + 64)
    Else
        GetColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & _
                          Chr(((ColumnNumber - 1) Mod 26) + 65)
    End If
End Function



