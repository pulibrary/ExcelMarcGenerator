Attribute VB_Name = "ProfileSelector"
Attribute VB_Base = "0{C5DC416D-B056-4EB8-A25B-65428A05A574}{5B0714DF-DCF7-4DB4-A733-F97EB6879C78}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public sFileName As String
Public sDefaultFileName As String
Public oInputLines As Object
Public oProfileNames As Object

Private Sub CancelButton_Click()
    Hide
End Sub

Private Sub OKButton_Click()
            For i = 0 To ProfileSelectorList.ListCount - 1
                Debug.Print i & " " & ProfileSelectorList.List(i)
                If ProfileSelectorList.Selected(i) Then
                    Debug.Print "*"
                End If
            Next i
    If ProfileSelector.OKButton.Caption = "Export" Then
        sFileName = Application.GetSaveAsFilename( _
            InitialFileName:=sDefaultFileName, _
            FileFilter:="Tab-delimited file (*.tab), *.tab")
        If sFileName = Null Then
            MsgBox ("Error saving file: no filename given")
            Exit Sub
        End If
            For i = 0 To ProfileSelectorList.ListCount - 1
                Debug.Print i & " " & ProfileSelectorList.List(i)
                If ProfileSelectorList.Selected(i) Then
                    Debug.Print "*"
                End If
            Next i
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(sFileName) Then
            iOverwrite = MsgBox("Overwrite the existing file '" & sFileName & "'?", vbOKCancel)
            If iOverwrite = vbCancel Then
                Exit Sub
            End If
        End If

        iFile = FreeFile
        Open sFileName For Output As iFile
        
        With ThisWorkbook.Worksheets("Profiles")
            iMaxProfileRow = 1
            Do While Len(.Cells(iMaxProfileRow, 1)) > 0
                iMaxProfileRow = iMaxProfileRow + 1
            Loop
            For i = 0 To ProfileSelectorList.ListCount - 1
                If ProfileSelectorList.Selected(i) Then
                    Debug.Print i
                    sSelectedProfile = ProfileSelectorList.List(i)
                    Debug.Print sSelectedProfile
                    For j = 1 To iMaxProfileRow - 1
                        sProfile = .Cells(j, 1).Value
                        If sProfile = sSelectedProfile Then
                            sOutputStr = .Cells(j, 1) & Chr(9) & .Cells(j, 2) & _
                                Chr(9) & .Cells(j, 3) & Chr(9) & .Cells(j, 4) & _
                                Chr(9) & .Cells(j, 5) & Chr(9) & .Cells(j, 6)
                            Print #iFile, sOutputStr
                        End If
                    Next j
                End If
            Next i
        End With
        Close #iFile
        MsgBox ("Export of profiles successful")
    Else 'Import
        bProfilesSelected = False
        iMaxProfileRow = 1
        Do While Len(ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
            iMaxProfileRow = iMaxProfileRow + 1
        Loop
        For i = 0 To ProfileSelectorList.ListCount - 1
            sProfileName = ProfileSelectorList.List(i)
            If ProfileSelectorList.Selected(i) Then
                bProfilesSelected = True
                With ThisWorkbook.Worksheets("Profiles")
                    iMaxProfileRow = 1
                    Do While Len(.Cells(iMaxProfileRow, 1)) > 0
                        iMaxProfileRow = iMaxProfileRow + 1
                    Loop
                    For j = 1 To iMaxProfileRow - 1
                        If sProfileName = .Cells(j, 1) Then
                            MsgBox ("The profile '" & sProfileName & "' already " & _
                                "exists.  Please rename the existing profile before " & _
                                "importing the new one.  Import aborted.")
                            Exit Sub
                        End If
                    Next j
                End With
            End If
        Next i
        
        iFile = FreeFile
        Open ProfileSelector.sFileName For Input As iFile
        sInputLine = ""
        Do While Not EOF(iFile)
            Line Input #iFile, sInputLine
            ProfileSelector.oInputLines.Add sInputLine
            aFields = Split(sInputLine, Chr(9))
            On Error Resume Next
            For i = 0 To ProfileSelectorList.ListCount - 1
                If ProfileSelectorList.Selected(i) Then
                    If ProfileSelectorList.List(i) = aFields(0) Then
                       ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 1) = aFields(0)
                       ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 2) = aFields(1)
                       ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 3) = aFields(2)
                       ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 4) = aFields(3)
                       ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 5) = aFields(4)
                       ThisWorkbook.Worksheets("Profiles").Cells(iMaxProfileRow, 6) = aFields(5)
                       iMaxProfileRow = iMaxProfileRow + 1
                       Exit For
                    End If
                End If
            Next i
            On Error GoTo 0
        Loop
        Close #iFile
        
        If Not (bProfilesSelected) Then
            MsgBox ("No profiles selected for import")
            Exit Sub
        Else
            Excel2MARC.UpdateProfileList
            MsgBox ("Import of Profiles Successful")
        End If
        ThisWorkbook.Save
    End If
    ProfileSelector.Hide
End Sub

Private Sub UserForm_Click()

End Sub