Private Sub ExportProfileButton_Click()
    ProfileSelector.OKButton.Caption = "Export"
    ProfileSelector.ProfileSelectorList.List = MARCWindow.ProfileComboBox.List
    ProfileSelector.Show
End Sub


Private Sub ImportProfileButton_Click()
    With ProfileSelector
        .sFileName = ""
        .sDefaultFileName = "ProfileExport.tab"
        .OKButton.Caption = "Import"
        .sFileName = Application.GetOpenFilename( _
            FileFilter:="Tab-delimited file (*.tab), *.tab")
        If .sFileName = "" Then
            MsgBox ("Error opening file: no filename given")
            Exit Sub
        End If
        Set .oInputLines = New Collection
        Set .oProfileNames = New Collection
        
        iFile = FreeFile
        Open .sFileName For Input As iFile
        sInputLine = ""
        Do While Not EOF(iFile)
            Line Input #iFile, sInputLine
            .oInputLines.Add sInputLine
            aFields = Split(sInputLine, Chr(9))
            On Error Resume Next
            .oProfileNames.Add aFields(0), aFields(0)
            On Error GoTo 0
        Loop
        
        .ProfileSelectorList.Clear
        For Each sProfileName In .oProfileNames
            .ProfileSelectorList.AddItem sProfileName
        Next sProfileName
        
        Close #iFile
        .Show
    End With
End Sub

Private Sub MARCPreviewBox_Click()

End Sub

Private Sub ProfileListBox_Change()
    With ProfileListBox
    iSelected = -1
    For i = 0 To UBound(.List)
        If .Selected(i) = True Then
            iSelected = i
            Exit For
        End If
    Next i
    If iSelected = 0 Then
        MARCWindow.MoveFieldUpButton.Enabled = False
        MARCWindow.MoveFieldDownButton.Enabled = True
    ElseIf iSelected = UBound(.List) Then
        MARCWindow.MoveFieldUpButton.Enabled = True
        MARCWindow.MoveFieldDownButton.Enabled = False
    ElseIf iSelected = -1 Then
        MARCWindow.MoveFieldUpButton.Enabled = False
        MARCWindow.MoveFieldDownButton.Enabled = False
    Else
        MARCWindow.MoveFieldUpButton.Enabled = True
        MARCWindow.MoveFieldDownButton.Enabled = True
    End If
    End With
 
End Sub

Private Sub MoveField(bDown As Boolean)
    sSelProfile = MARCWindow.ProfileComboBox.Value
    With MARCWindow.ProfileListBox
    iSelected = -1
    iLower = 0
    iUpper = UBound(.List) - 1
    If Not bDown Then
        iLower = iLower + 1
        iUpper = iUpper + 1
    End If
    
    For i = iLower To iUpper
        If .Selected(i) = True Then
            iSelected = i
            Exit For
        End If
    Next i
    End With
    If iSelected = -1 Then
        Exit Sub
    End If
    
    iSelected = iSelected + 1
    
    iMaxProfileRow = 1
    Do While Len(ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
        iMaxProfileRow = iMaxProfileRow + 1
    Loop
    
    iLastRow = 0
    iProfileIndex = 1
    
    For i = 1 To iMaxProfileRow - 1
        iLastRow = i
        sProfile = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 1).Value
        If StrComp(sProfile, "") = 0 Then
            Exit For
        End If
        
         If StrComp(sProfile, sSelProfile, 1) = 0 Then
            If iProfileIndex = iSelected Then
                If bDown Then
                    t1 = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i + 1, 8).Value
                Else
                    t1 = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i - 1, 8).Value
                End If
                t2 = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 8).Value
                Debug.Print CStr(t1) + " " + CStr(t2)
                
                If bDown Then
                    ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i + 1, 8).Value = t2
                Else
                    ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i - 1, 8).Value = t2
                End If
                ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 8).Value = t1
                
                Excel2MARC.UpdateProfileWindow
                If bDown Then
                    MARCWindow.ProfileListBox.Selected(iSelected) = True
                Else
                    MARCWindow.ProfileListBox.Selected(iSelected - 2) = True
                End If
                
                Exit For
            End If
            iProfileIndex = iProfileIndex + 1
        End If
    Next i
End Sub

Private Sub MoveFieldDownButton_Click()
    MoveField (True)
End Sub

Private Sub MoveFieldUpButton_Click()
    MoveField (False)
End Sub

Private Sub RenameProfileButton_Click()
    Application.ScreenUpdating = False
    With MARCWindow.ProfileComboBox
        sOldProfileName = .List(.ListIndex)
        sNewProfileName = MARCWindow.NewProfileNameTextBox.Value
        If StrComp(sNewProfileName, "") = 0 Then
            MsgBox ("Please enter the new name in 'New Profile Name'")
        ElseIf StrComp(sOldProfileName, sNewProfileName) = 0 Then
            MsgBox ("The old and new names are the same.")
        Else
            bExists = False
            iMaxProfileRow = 1
            Do While Len(ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
                iMaxProfileRow = iMaxProfileRow + 1
            Loop
            For i = 1 To iMaxProfileRow - 1
              sProfile = ThisWorkbook.Worksheets("Profiles").Cells(i, 1).Value
              If StrComp(sProfile, sNewProfileName) = 0 Then
                 MsgBox ("The profile name '" & sProfile & "' already exists")
                 bExists = True
                 Exit For
              End If
            Next i
            If Not bExists Then
                For i = 1 To iMaxProfileRow - 1
                    sProfile = ThisWorkbook.Worksheets("Profiles").Cells(i, 1).Value
                    If StrComp(sProfile, sOldProfileName) = 0 Then
                        ThisWorkbook.Worksheets("Profiles").Cells(i, 1).Value = sNewProfileName
                    End If
                Next i
                Excel2MARC.UpdateProfileList
                .Value = sNewProfileName
            End If
        End If
    End With
    ThisWorkbook.Save
    Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Initialize()
    Excel2MARC.bEvents = True
End Sub

Private Sub AddProfileButton_Click()
    Application.ScreenUpdating = False
    sDefault000 = "$Lnam#a22$S5u#4500"
    sDefault008 = "$DsDATE####cc######r#########0#chi#d"
    sNewProfileName = MARCWindow.NewProfileNameTextBox.Value
    If StrComp(sNewProfileName, "") = 0 Then
        MsgBox ("Please enter the new name in 'New Profile Name'")
    Else
        With ThisWorkbook.Worksheets("Profiles").Range("A2")
            iMaxProfileRow = 1
            Do While Len(ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
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
    ThisWorkbook.Save
    Application.ScreenUpdating = True
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
    Do While Len(ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
        iMaxProfileRow = iMaxProfileRow + 1
    Loop
    For i = 1 To iMaxProfileRow - 1
        sProfile = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 1).Value
        sProfileField = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 2).Value
        sProfileSeq = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 3).Value
        sProfileInd1 = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 4).Value
        sProfileInd2 = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 5).Value
        sProfileValue = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 6).Value
        If StrComp(sProfile, sSelProfile, 1) = 0 And _
            StrComp(sProfileSeq, sSeq, 1) = 0 And _
            StrComp(sProfileField, sField, 1) = 0 And _
            StrComp(sProfileInd1, sInd1, 1) = 0 And _
            StrComp(sProfileInd2, sInd2, 1) = 0 And _
            StrComp(sProfileValue, sValue, 1) = 0 _
        Then
            ThisWorkbook.Worksheets("Profiles").Rows(i + 1).EntireRow.Delete xlShiftUp
            If iSel > 0 Then
                MARCWindow.ProfileListBox.ListIndex = iSel - 1
            Else
                MARCWindow.ProfileListBox.ListIndex = 0
            End If
            ThisWorkbook.Save
            Excel2MARC.UpdateProfileWindow
        End If
    Next
    FieldTextBox.SetFocus
End Sub

Private Sub DeleteProfileButton_Click()
    Application.ScreenUpdating = False
    sSelProfile = MARCWindow.ProfileComboBox.Value
    iMaxProfileRow = 1
    Do While Len(ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
        iMaxProfileRow = iMaxProfileRow + 1
    Loop
    i = 1
    Do While i < iMaxProfileRow + 1
        If StrComp(ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 1).Value, sSelProfile, 1) = 0 Then
            ThisWorkbook.Worksheets("Profiles").Rows(i + 1).EntireRow.Delete xlShiftUp
            iMaxProfileRow = iMaxProfileRow - 1
        Else
            i = i + 1
        End If
    Loop
    MARCWindow.ProfileComboBox.ListIndex = 0
    Excel2MARC.UpdateProfileList
    ThisWorkbook.Save
    Application.ScreenUpdating = True
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


Private Sub SelectAllButton_Click()
    Excel2MARC.bEvents = False
    For i = 0 To MARCWindow.PreviewListBox.ListCount - 1
        MARCWindow.PreviewListBox.Selected(i) = True
    Next
    Excel2MARC.bEvents = True
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
    Do While Len(ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
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
        sProfile = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 1).Value
        sProfileField = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 2).Value
        sProfileSeq = ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 3).Value
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
            ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 4).Value = sInd1
            ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 5).Value = sInd2
            ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(i, 6).Value = sValue
            Excel2MARC.UpdateProfileWindow
            bFound = True
            Exit For
        End If
    Next
    If bGenerateSeq Then
        sSeq = iMaxSeq + 1
    End If
    If Not bFound Then
        ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(iLastRow, 1).Value = sSelProfile
        ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(iLastRow, 2).Value = sField
        ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(iLastRow, 3).Value = sSeq
        ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(iLastRow, 4).Value = sInd1
        ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(iLastRow, 5).Value = sInd2
        ThisWorkbook.Worksheets("Profiles").Range("A2").Cells(iLastRow, 6).Value = sValue
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
    ThisWorkbook.Save
    FieldTextBox.SetFocus
End Sub
