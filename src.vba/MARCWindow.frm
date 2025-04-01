Attribute VB_Name = "MARCWindow"
Attribute VB_Base = "0{04FA21DE-EAF8-4FCD-996D-12E2FD6E3CE1}{CA89C330-AF81-4E80-8A90-1568CFD6E311}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub Label15_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub UserForm_Initialize()
    Excel2MARC.bEvents = True
End Sub

Private Sub AddProfileButton_Click()
    sDefault000 = "$Lnam#a22$S5u#4500"
    sDefault008 = "$DsDATE####cc######r#########0#chi#d"
    sNewProfileName = MARCWindow.NewProfileNameTextBox.Value
    If StrComp(sNewProfileName, "") = 0 Then
        MsgBox ("Please give the new profile a name")
    Else
        With Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2")
            iMaxProfileRow = 1
            Do While Len(Workbooks("MARC.xlam").Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
                iMaxProfileRow = iMaxProfileRow + 1
            Loop
            iLastRow = 0
            For i = 1 To iMaxProfileRow - 1
              iLastRow = i
              sProfile = .Cells(i, 1).Value
              If StrComp(sProfile, "") = 0 Then
                 Exit For
              End If
            Next
            .Cells(iLastRow, 1) = sNewProfileName
            .Cells(iLastRow, 2) = "000"
            .Cells(iLastRow, 3) = "1"
            .Cells(iLastRow, 6) = sDefault000
            .Cells(iLastRow + 1, 1) = sNewProfileName
            .Cells(iLastRow + 1, 2) = "008"
            .Cells(iLastRow + 1, 3) = "1"
            .Cells(iLastRow + 1, 6) = sDefault008
            Excel2MARC.UpdateProfileList
        End With
        c = ProfileComboBox.ListCount
        For i = 0 To c - 1
            If StrComp(sNewProfileName, ProfileComboBox.List(i)) = 0 Then
                ProfileComboBox.ListIndex = i
                UpdateProfileWindow
            End If
        Next
    End If
    Workbooks("MARC.xlam").Save
End Sub

Private Sub CancelButton_Click()
    MARCWindow.Hide
End Sub


Private Sub ConvertButton_Click()
    Excel2MARC.ConvertToMARC
End Sub

Private Sub DeleteEntryButton_Click()
    iSel = MARCWindow.ProfileListBox.ListIndex
    sField = MARCWindow.ProfileListBox.List(iSel, 0)
    sSeq = MARCWindow.ProfileListBox.List(iSel, 1)
    sInd1 = MARCWindow.ProfileListBox.List(iSel, 2)
    sInd2 = MARCWindow.ProfileListBox.List(iSel, 3)
    sValue = MARCWindow.ProfileListBox.List(iSel, 4)
    
    sSelProfile = MARCWindow.ProfileComboBox.Value
    iMaxProfileRow = 1
    Do While Len(Workbooks("MARC.xlam").Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
        iMaxProfileRow = iMaxProfileRow + 1
    Loop
    For i = 1 To iMaxProfileRow - 1
        sProfile = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 1).Value
        sProfileField = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 2).Value
        sProfileSeq = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 3).Value
        sProfileInd1 = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 4).Value
        sProfileInd2 = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 5).Value
        sProfileValue = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 6).Value
        If StrComp(sProfile, sSelProfile, 1) = 0 And _
            StrComp(sProfileSeq, sSeq, 1) = 0 And _
            StrComp(sProfileField, sField, 1) = 0 And _
            StrComp(sProfileInd1, sInd1, 1) = 0 And _
            StrComp(sProfileInd2, sInd2, 1) = 0 And _
            StrComp(sProfileValue, sValue, 1) = 0 _
        Then
            Workbooks("MARC.xlam").Worksheets("Profiles").Rows(i + 1).EntireRow.Delete xlShiftUp
        End If
        If iSel > 0 Then
            MARCWindow.ProfileListBox.ListIndex = iSel - 1
        Else
           MARCWindow.ProfileListBox.ListIndex = 0
        End If
        Excel2MARC.UpdateProfileWindow
    Next
    Workbooks("MARC.xlam").Save
    FieldTextBox.SetFocus
End Sub




Private Sub DeleteProfileButton_Click()
    sSelProfile = MARCWindow.ProfileComboBox.Value
    iMaxProfileRow = 1
    Do While Len(Workbooks("MARC.xlam").Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
        iMaxProfileRow = iMaxProfileRow + 1
    Loop
    i = 1
    Do While i < iMaxProfileRow + 1
        If StrComp(Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 1).Value, sSelProfile, 1) = 0 Then
            Workbooks("MARC.xlam").Worksheets("Profiles").Rows(i + 1).EntireRow.Delete xlShiftUp
            iMaxProfileRow = iMaxProfileRow - 1
        Else
            i = i + 1
        End If
    Loop
    MARCWindow.ProfileComboBox.ListIndex = 0
    Excel2MARC.UpdateProfileList
    Workbooks("MARC.xlam").Save
End Sub

Private Sub FieldTextBox_Change()

End Sub

Private Sub Frame1_Click()

End Sub


Private Sub MARCPreviewBox_Click()

End Sub

Private Sub PreviewListBox_Change()
    If Excel2MARC.bEvents Then
        Excel2MARC.UpdateMARCPreview
    End If
End Sub

Private Sub ProfileComboBox_Click()
    Excel2MARC.UpdateProfileList
End Sub


Private Sub ProfileListBox_Click()
    iSel = MARCWindow.ProfileListBox.ListIndex
    sField = MARCWindow.ProfileListBox.List(iSel, 0)
    sSeq = MARCWindow.ProfileListBox.List(iSel, 1)
    sInd1 = MARCWindow.ProfileListBox.List(iSel, 2)
    sInd2 = MARCWindow.ProfileListBox.List(iSel, 3)
    sValue = MARCWindow.ProfileListBox.List(iSel, 4)
    MARCWindow.FieldTextBox.Value = sField
    MARCWindow.SeqTextBox.Value = sSeq
    MARCWindow.Ind1TextBox.Value = sInd1
    MARCWindow.Ind2TextBox.Value = sInd2
    MARCWindow.ValueTextBox.Value = sValue
    
End Sub

Private Sub ProfilesComboBox_Change()

End Sub

Private Sub SelectAllButton_Click()
    Excel2MARC.bEvents = False
    For i = 0 To MARCWindow.PreviewListBox.ListCount - 1
        MARCWindow.PreviewListBox.Selected(i) = True
    Next
    Excel2MARC.bEvents = True
    Excel2MARC.UpdateMARCPreview
End Sub

Private Sub SelectAllButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

End Sub

Private Sub TitleRowCheckBox_Click()
    Excel2MARC.UpdateMARCPreview
End Sub

Private Sub UpdateEntryButton_Click()
    sSelProfile = MARCWindow.ProfileComboBox.Value
    sSeq = MARCWindow.SeqTextBox.Value
    sField = MARCWindow.FieldTextBox.Value
    sInd1 = MARCWindow.Ind1TextBox.Value
    sInd2 = MARCWindow.Ind2TextBox.Value
    sValue = MARCWindow.ValueTextBox.Value
    iMaxProfileRow = 1
    Do While Len(Workbooks("MARC.xlam").Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
        iMaxProfileRow = iMaxProfileRow + 1
    Loop
    
    bFound = False
    bGenerateSeq = False
    iMaxSeq = 1
    If Len(sSeq) = 0 Then
        bGenerateSeq = True
    End If
    
    iLastRow = 0
    For i = 1 To iMaxProfileRow - 1
        iLastRow = i
        sProfile = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 1).Value
        sProfileField = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 2).Value
        sProfileSeq = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 3).Value
        If StrComp(sProfile, "") = 0 Then
            Exit For
        End If
        If bGenerateSeq And _
            StrComp(sProfile, sSelProfile, 1) = 0 And _
            StrComp(Mid(sProfileField, 1, 3), Mid(sField, 1, 3), 1) = 0 And _
            Int(sProfileSeq) >= iMaxSeq _
        Then
            iMaxSeq = sProfileSeq
        End If
        
        If StrComp(sProfile, sSelProfile, 1) = 0 And _
            StrComp(sProfileField, sField, 1) = 0 And _
            StrComp(sProfileSeq, sSeq, 1) = 0 _
        Then
            Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 4).Value = sInd1
            Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 5).Value = sInd2
            Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 6).Value = sValue
            Excel2MARC.UpdateProfileWindow
            bFound = True
            Exit For
        End If
    Next
    If bGenerateSeq Then
        sSeq = iMaxSeq + 1
    End If
    If Not bFound Then
        Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(iLastRow, 1).Value = sSelProfile
        Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(iLastRow, 2).Value = sField
        Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(iLastRow, 3).Value = sSeq
        Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(iLastRow, 4).Value = sInd1
        Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(iLastRow, 5).Value = sInd2
        Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(iLastRow, 6).Value = sValue
        Excel2MARC.UpdateProfileWindow
    End If
    c = ProfileListBox.ListCount
    For i = 0 To c - 1
        If StrComp(ProfileListBox.List(i, 0), sField) = 0 And _
            StrComp(ProfileListBox.List(i, 1), sSeq) = 0 _
        Then
            ProfileListBox.ListIndex = i
            Exit For
        End If
    Next
    Workbooks("MARC.xlam").Save
    FieldTextBox.SetFocus
End Sub

Private Sub UserForm_Click()

End Sub