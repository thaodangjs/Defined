Attribute VB_Name = "Module1"
Function Giamgia(soluong As Byte, ngay As Date, MaCT As String) As Double
If soluong > 10 Or Weekday(ngay) = 1 Then
Giamgia = 0.05
Else
If soluong >= 5 And soluong <= 10 And Left(MaCT, 1) = "A" Then
Giamgia = 0.07
Else
Giamgia = 0
End If
End If

End Function
