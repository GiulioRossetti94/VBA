Public Function taxPaid(I As Double) As Double

If I <= 15000 Then
    taxPaid = I * 0.23
ElseIf I > 15000 And I <= 28000 Then
    t1 = 15000 * 0.23
    taxPaid = t1 + (I - 15000) * 0.27
ElseIf I > 28000 And I <= 55000 Then
    t1 = 15000 * 0.23
    t2 = 13000 * 0.27
    taxPaid = t1 + t2 + (I - 28000) * 0.38
ElseIf I > 55000 And I <= 75000 Then
    t1 = 15000 * 0.23
    t2 = 13000 * 0.27
    t3 = 27000 * 0.38
    taxPaid = t1 + t2 + t3 + (I - 55000) * 0.41
Else
    t1 = 15000 * 0.23
    t2 = 13000 * 0.27
    t3 = 27000 * 0.38
    t4 = 20000 * 0.41
    taxPaid = t1 + t2 + t3 + t4 + (I - 75000) * 0.43
End If


End Function

Public Function taxPaidUK(I As Double) As Double

If I <= 11850 Then
    taxPaidUK = 0
ElseIf I > 11850 And I <= 46350 Then
    t1 = 0
    taxPaidUK = t1 + (I - 11850) * 0.2
ElseIf I > 46351 And I <= 150000 Then
    t1 = 0
    t2 = 34500 * 0.2
    taxPaidUK = t1 + t2 + (I - 46351) * 0.4
Else
    t1 = 0
    t2 = 34500 * 0.2
    t3 = 103646 * 0.4
    taxPaidUK = t1 + t2 + t3 + (I - 150000) * 0.45
End If


End Function
