Private Sub CheckBox10_Change()
If Me.CheckBox10.Value Then Me.OLEObjects("CheckBox15").Object.Value = False

End Sub

Private Sub CheckBox11_Change()
If Me.CheckBox11.Value Then Me.OLEObjects("CheckBox15").Object.Value = False

End Sub

Private Sub CheckBox12_Change()
If Me.CheckBox12.Value Then Me.OLEObjects("CheckBox15").Object.Value = False

End Sub

Private Sub CheckBox15_Change()
Dim x As Long
If Me.CheckBox15.Value Then
    For x = 7 To 13
        Me.OLEObjects("CheckBox" & x).Object.Value = False
    Next x

End If

End Sub

Private Sub CheckBox13_Click()
If Me.CheckBox13.Value Then Me.OLEObjects("CheckBox15").Object.Value = False
End Sub

Private Sub CheckBox3_Click()

End Sub

Private Sub CheckBox7_Change()
If Me.CheckBox7.Value Then Me.OLEObjects("CheckBox15").Object.Value = False

End Sub

Private Sub CheckBox8_Change()
If Me.CheckBox8.Value Then Me.OLEObjects("CheckBox15").Object.Value = False
End Sub

Private Sub CheckBox9_Change()
If Me.CheckBox9.Value Then Me.OLEObjects("CheckBox15").Object.Value = False

End Sub

Private Sub CommandButton1_Click()
Call DoYouWannaBeMyPPT
End Sub

Private Sub CommandButton10_Click()
Call calcu
End Sub

Private Sub CommandButton11_Click()
Call bloombergExcel
End Sub

Private Sub CommandButton12_Click()
Call table_comparison
End Sub

Private Sub CommandButton13_Click()
Call table_perf
End Sub

Private Sub CommandButton14_Click()
Call PreparingData
End Sub

Private Sub CommandButton15_Click()
Call PreparingDataAllocation
End Sub

Private Sub CommandButton2_Click()
Call WeDontNeedNoEduMail
End Sub

Private Sub CommandButton3_Click()
Call NewReport
End Sub

Private Sub CommandButton4_Click()
Call grabSend
End Sub

Private Sub CommandButton5_Click()
Call SendingToFolder
End Sub

Private Sub CommandButton6_Click()
Call SaveAttach
End Sub

Private Sub CommandButton7_Click()
Call CTG
End Sub

Private Sub CommandButton8_Click()
Call ExcelFileBancaFinInt
End Sub

Private Sub CommandButton9_Click()
Call MorningstarExcel
End Sub