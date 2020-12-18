Attribute VB_Name = "nextBirthday"
Public Function nextBirthday(Arg1 As Range) As Integer

    today = Date
    currYear = Right(today, 4)

    Dim i As Integer
    nearBirthday = -1
    
    For i = 1 To Arg1.Rows.Count
        birthDate = Arg1(i)
        birthDate = Left(birthDate, 5) & "." & currYear
        
        Debug.Print today
        Debug.Print birthDate
        
        daysBetween = DateDiff("d", today, birthDate)
        
        Debug.Print daysBetween
        
        If daysBetween >= 0 Then
            If nearBirthday = -1 Then
                nearBirthday = daysBetween
            End If
            If daysBetween < nearBirthday Then
                nearBirthday = daysBetween
            End If
        End If
        
        If i = Arg1.Rows.Count Then
            If nearBirthday > -1 Then
                nextBirthday = nearBirthday
            ElseIf nearBirthday < 0 Then
                currYear = currYear + 1
                i = 0   ' For some reason, the value of the first index seems to change
            End If
        End If
        
    Next i

End Function






    
        
