Attribute VB_Name = "Akamai_Func_Procedures"
Function GetNumeric(CellRef As String) As Long
    Dim StringLength As Integer
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
    Next i    
    GetNumeric = result
End Function


Function AK_VoteCheck(Age As Integer, Nationality As String) As String
Nationality = LCase(Nationality)
If Nationality = "indian" And Age >= 18 Then
AK_VoteCheck = "Eligible"
Else
AK_VoteCheck = "Not Eligible"

End Function

Public Function AK_TAX11(Salary As Currency) As Currency
If Salary <= 250000 Then
    AK_TAX = 0
ElseIf Salary > 250000 And Salary <= 500000 Then
    AK_TAX = 0.05 * (Salary - 250000)
ElseIf Salary > 500000 And Salary <= 1000000 Then
    AK_TAX = 12500 + 0.2 * (Salary - 500000)
Else
    AK_TAX = 112500 + 0.3 * (Salary - 100000)
End If

End Function


Function AK_RoomArea(length As Single, Breadth As Single) As Single
Dim Area As Single

Area = length * Breadth
AK_RoomArea = Area

End Function

Function AK_Vote_Check(Age As Integer, Nationality As String) As String
Nationality = LCase(Nationality)
If Nationality = "indian" And Age >= 18 Then
AK_VoteCheck = "Eligible"
Else
AK_VoteCheck = "Not Eligible"
End If
End Function



Function AK_TAX(Salary As Currency) As Currency
If Salary <= 250000 Then
    AK_TAX = 0
ElseIf Salary > 250000 And Salary <= 500000 Then
    AK_TAX = 0.05 * (Salary - 250000)
ElseIf Salary > 500000 And Salary <= 1000000 Then
    AK_TAX = 12500 + 0.2 * (Salary - 500000)
Else
    AK_TAX = 112500 + 0.3 * (Salary - 100000)
End If

End Function

Function AK_Numsep(Data As String) As Long
    Dim length As Integer, count As Integer
    
    length = Len(Data)
    
    Dim extchar As String, result As String
    
    For count = 1 To length
        extchar = Mid(Data, count, 1)
        If Asc(extchar) >= 48 And Asc(extchar) <= 57 Then
        result = result & extchar
        End If
    Next count
    
    AK_Numsep = Val(result)
    
End Function

Function AK_Textsep(Data As String) As String
    Dim length As Integer, count As Integer
    
    length = Len(Data)
    
    Dim extchar As String, result As String
    
    For count = 1 To length
        extchar = Mid(Data, count, 1)
        If (Asc(extchar) >= 65 And Asc(extchar) <= 90) Or (Asc(extchar) >= 97 And Asc(extchar) <= 122) Then
        result = result & extchar
        End If
    Next count
    
    AK_Textsep = result
    
End Function

