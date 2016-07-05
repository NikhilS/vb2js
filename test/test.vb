' Condition 0
        If (Not IsArray(SourceArray)) Or (Not IsArray(DestArray)) Then
        Exit Sub
        End If

        ' Condition 8
        If UBound(SourceArray) = 0 Then
        ReDim DestArray(0 To 0)
        Exit Sub
        End If

        ' Condition 7
        If startingpoint < 1 Then
        startingpoint = 1
        End If
        ' Condition 7
        If numberofvalues < 0 Then
        numberofvalues = 0
        End If

        ' Condition 4
        If startingpoint > UBound(SourceArray) Then
        Exit Sub
        End If

        ' Condition 6
        If numberofvalues > 0 And _
        (startingpoint + numberofvalues) > UBound(SourceArray) Then
        Exit Sub
        End If

        Dim lngDestLength As Long, lngCnt As Long

        ' Condition 5
        If numberofvalues = 0 Then
        lngDestLength = (UBound(SourceArray) - startingpoint) + 1
        Else
        lngDestLength = numberofvalues
        End If

        ' Conditions 1, 2, 3
        ReDim DestArray(lngDestLength)

        For lngCnt = 1 To lngDestLength
        DestArray(lngCnt) = SourceArray((startingpoint + lngCnt) - 1)
        Next


        Exit Sub


        MsgBox "CopyArray(): Error occurred.", vbInformation, "Error occurred."
        MsgBox Err.Number & ": " & Err.Description


        End Sub
