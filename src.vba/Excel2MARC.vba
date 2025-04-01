Attribute VB_Name = "Excel2MARC"
Public bEvents As Boolean
Public iMaxColumn As Integer
Public iMaxRow As Integer

Sub makeMARC(control As IRibbonControl)
    If Right(ActiveWorkbook.FullName, 4) = ".xls" Then
        iResult = MsgBox("File must be in XLSX format.  Convert Now?", vbYesNo, "Question")
        If iResult = vbYes Then
            sXLSname = ActiveWorkbook.FullName
            sXLSXname = Replace(sXLSname, ".xls", ".xlsx")
            ActiveWorkbook.SaveAs Filename:=sXLSXname, FileFormat:=xlOpenXMLWorkbook
            Kill sXLSname
            Workbooks.Open sXLSXname
        Else
            Exit Sub
        End If
    End If

    iMaxColumn = ActiveSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Column
    MARCWindow.PreviewListBox.ColumnCount = iMaxColumn
    
    Application.DisplayAlerts = False
    Workbooks("MARC.xlam").Worksheets("Scratch").Delete
    Workbooks("MARC.xlam").Worksheets.Add().Name = "Scratch"
    iMaxRow = 1
    
    bEvents = False
    UpdateProfileList (False)
    MARCWindow.ProfileComboBox.ListIndex = Workbooks("MARC.xlam").Worksheets("Profiles").Cells(1, 7)
    UpdateProfileWindow
    bEvents = True
    
    For i = 1 To iMaxColumn
        iLocalLastRow = ActiveSheet.Range(Cells(65534, i), Cells(65534, i)).End(xlUp).Row
        If iLocalLastRow > iMaxRow Then
            iMaxRow = iLocalLastRow
        End If
    Next
    
    If iMaxRow < Selection.Row + Selection.Rows.Count - 1 Then
        ActiveSheet.Range(Rows(Selection.Row), Rows(iMaxRow)).Select
    End If
    Selection.Cells.EntireRow.Copy
    Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    MARCWindow.PreviewListBox.RowSource = "[MARC.xlam]Scratch!2:" & (Selection.Rows.Count + 1)
    MARCWindow.PreviewListBox.ListIndex = 0
    For i = 0 To MARCWindow.PreviewListBox.ListCount - 1
        If MARCWindow.PreviewListBox.Selected(i) Then
            MARCWindow.PreviewListBox.Selected(i) = False
        End If
    Next i
    
    
    UpdateMARCPreview
    MARCWindow.ConvertButton.SetFocus
    MARCWindow.Show
        
End Sub

Sub UpdateColumnHeaders()
    Dim RegEx
    Set RegEx = CreateObject("vbscript.regexp")
    With RegEx
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
    End With
    
    sWidths = ""
    With Workbooks("MARC.xlam").Worksheets("Profiles")
        ReDim aColumns(1 To iMaxColumn) As String
        For iCol = 1 To iMaxColumn
            sResult = ""
            RegEx.Pattern = "\$" & iCol & "(?![0-9])"
            For i = 0 To MARCWindow.ProfileListBox.ListCount - 1
                sValue = .Cells(i + 2, 6)
                Set aMatches = RegEx.Execute(sValue)
                If aMatches.Count > 0 Then
                    If sResult = "" Then
                        sResult = .Cells(i + 2, 2)
                    Else
                        sResult = sResult & "," & .Cells(i + 2, 2)
                    End If
                    For j = 0 To MARCWindow.ProfileListBox.ListCount - 1
                       If .Cells(j + 2, 2) = .Cells(i + 2, 2) And .Cells(j + 2, 3) > 1 Then
                          sResult = sResult & "[" & .Cells(i + 2, 3) & "]"
                          Exit For
                       End If
                    Next j
                End If
            Next i
            If sResult <> "" Then
                sResult = "(" & sResult & ")"
            End If
            Workbooks("MARC.xlam").Worksheets("Scratch").Cells(1, iCol) = "Column " & iCol & " " & sResult
            sWidths = sWidths & Application.International(xlRowSeparator) & CStr(50 + (Len(sResult) * 4))
        Next iCol
    End With
    MARCWindow.PreviewListBox.RowSource = "[MARC.xlam]Scratch!2:" & (iMaxRow + 1)
    sWidths = Mid(sWidths, 2)
    MARCWindow.PreviewListBox.ColumnWidths = sWidths
End Sub



'GuiCtrlCreateLabel("$X: Column X", 10, 450, 275, 20, 0)
'GuiCtrlCreateLabel("$X[Y]: Substring of column X starting at character Y", 10, 470, 275, 15, 0)
'GuiCtrlCreateLabel("$X[Y,Z]: Characters Y-Z of Column X", 10, 485, 275, 15, 0)
'GuiCtrlCreateLabel("$D: Current date (6-digit)", 10, 500, 275, 15, 0)
'GuiCtrlCreateLabel("$L: Length of record", 10, 515, 275, 15, 0)
'GuiCtrlCreateLabel("$S: Start address of data", 10, 530, 275, 15, 0)
'

Function UpdateProfileList(Optional bUpdateWindow = True)
    With MARCWindow.ProfileComboBox
        c = .ListCount
        If c > 0 Then
            For i = 0 To c - 1
               .RemoveItem 0
            Next
        End If
        
       iMaxProfileRow = 1
        Do While Len(Workbooks("MARC.xlam").Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
            iMaxProfileRow = iMaxProfileRow + 1
        Loop
        Workbooks("MARC.xlam").Worksheets("Profiles").Range("A1:I" & iMaxProfileRow).Sort _
            Key1:=Workbooks("MARC.xlam").Worksheets("Profiles").Columns("A"), _
            Order1:=xlAscending, _
            Header:=xlYes
        For i = 1 To iMaxProfileRow
            bMatch = False
            sProfile = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 1)
            If sProfile <> sPrevProfile And sProfile <> "" Then
                For j = 0 To .ListCount - 1
                    If StrComp(.List(j), sProfile, vbTextCompare) = 1 Then
                        .AddItem sProfile, j
                        bMatch = True
                        Exit For
                    End If
                Next j
                If Not (bMatch) Then
                    .AddItem (sProfile)
                End If
            End If
            sPrevProfile = sProfile
        Next i
        If (.ListIndex < 0) Then
            .ListIndex = 0
        End If
        Workbooks("MARC.xlam").Worksheets("Profiles").Cells(1, 7) = .ListIndex
    End With
    If bUpdateWindow Then
        UpdateProfileWindow
    End If
End Function

Sub UpdateProfileWindow()
    With MARCWindow.ProfileListBox
        sSelProfile = MARCWindow.ProfileComboBox.Value
        iMaxProfileRow = 1
        Do While Len(Workbooks("MARC.xlam").Worksheets("Profiles").Cells(iMaxProfileRow, 1)) > 0
            iMaxProfileRow = iMaxProfileRow + 1
        Loop
        iProfileRows = 0
        For i = 1 To iMaxProfileRow - 1
            sProfile = Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 1).Value
            If sProfile <> "" Then
                If StrComp(sProfile, sSelProfile, 1) = 0 Then
                   iProfileRows = iProfileRows + 1
                   Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 7) = 1
                Else
                   Workbooks("MARC.xlam").Worksheets("Profiles").Range("A2").Cells(i, 7) = 0
                End If
            End If
        Next
        With Workbooks("MARC.xlam").Worksheets("Profiles")
        iMaxProfileRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row
        .Range("A1:H" & iMaxProfileRow).Sort _
            Key1:=.Columns("G"), _
            Order1:=xlDescending, _
            Key2:=.Columns("B"), _
            Order2:=xlAscending, _
            Key3:=.Columns("C"), _
            Order3:=xlAscending, _
            Header:=xlYes
        End With
        If (iProfileRows > 0) Then
            MARCWindow.ProfileListBox.RowSource = "[MARC.xlam]Profiles!B2:F" & iProfileRows + 1
        End If
    End With
    UpdateColumnHeaders
    UpdateMARCPreview
End Sub
Function UTF8Size(sString As String) As Integer
    Dim UTFStream As Object
    Set UTFStream = CreateObject("ADODB.stream")
    UTFStream.Type = 2
    UTFStream.Mode = 3
    UTFStream.Charset = "UTF-8"
    UTFStream.LineSeparator = 10
    UTFStream.Open
    UTFStream.WriteText sString, 0
    UTFStream.Flush
    UTF8Size = UTFStream.Size - 3
End Function

Function ResolveVariablesLeader(sLeader As String, sDirectory As String, sRecord As String) As String
    Dim RegEx
    Set RegEx = CreateObject("vbscript.regexp")
    With RegEx
        .MultiLine = False
        .Global = False
        .IgnoreCase = True
    End With
    
    RegEx.Pattern = "\$L"
    sRecLen = Format(UTF8Size(sRecord) + UTF8Size(sDirectory) + 24 + 2, "00000")
    sLeader = RegEx.Replace(sLeader, sRecLen)
    
    RegEx.Pattern = "\$S"
    sStartPos = Format(UTF8Size(sDirectory) + 24 + 1 + 1, "00000")
    sLeader = RegEx.Replace(sLeader, sStartPos)
    sLeader = Replace(sLeader, ChrW(-257), "")
    ResolveVariablesLeader = sLeader

End Function

Function ResolveVariables(sField As String, iRow As Integer) As String
    Dim RegEx
    Set RegEx = CreateObject("vbscript.regexp")
    With RegEx
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
    End With
    
    RegEx.Pattern = "\$([0-9]+)\[([0-9]+),([0-9]+)\]"
    Set aMatches = RegEx.Execute(sField)
    For Each sMatch In aMatches
        sString = sMatch.subMatches(0)
        iCol = Int(sString)
        sCol = Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").Cells(iRow + 1, iCol).Value
        sFormat = Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").Cells(iRow + 1, iCol).NumberFormat
        If StrComp(sFormat, "General") <> 0 Then
            sCol = Format(sCol, sFormat)
        End If
        sString = sMatch
        sString = Replace(sString, "$", "\$")
        sString = Replace(sString, "[", "\[")
        sString = Replace(sString, "]", "\]")
        sVal = Mid(sCol, sMatch.subMatches(1), sMatch.subMatches(2) - sMatch.subMatches(1) + 1)
        sVal = CleanUpColumn(sVal)
        RegEx.Pattern = sString
        sField = RegEx.Replace(sField, sVal)
    Next
    
    RegEx.Pattern = "\$([0-9]+)\[-([0-9]+)\]"
    Set aMatches = RegEx.Execute(sField)
    For Each sMatch In aMatches
        sString = sMatch.subMatches(0)
        iCol = Int(sString)
        sCol = Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").Cells(iRow + 1, iCol).Value
        sFormat = Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").Cells(iRow + 1, iCol).NumberFormat
        If StrComp(sFormat, "General") <> 0 Then
            sCol = Format(sCol, sFormat)
        End If
        sString = sMatch
        sString = Replace(sString, "$", "\$")
        sString = Replace(sString, "[", "\[")
        sString = Replace(sString, "]", "\]")
        sVal = Right(sCol, sMatch.subMatches(1))
        sVal = CleanUpColumn(sVal)
        sField = RegEx.Replace(sField, sVal)
    Next
    
    RegEx.Pattern = "\$([0-9]+)\[([0-9]+)\]"
    Set aMatches = RegEx.Execute(sField)
    For Each sMatch In aMatches
        sString = sMatch.subMatches(0)
        iCol = Int(sString)
        sCol = Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").Cells(iRow + 1, iCol).Value
        sFormat = Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").Cells(iRow + 1, iCol).NumberFormat
        If StrComp(sFormat, "General") <> 0 Then
            sCol = Format(sCol, sFormat)
        End If
        sString = sMatch
        sString = Replace(sString, "$", "\$")
        sString = Replace(sString, "[", "\[")
        sString = Replace(sString, "]", "\]")
        sVal = Mid(sCol, sMatch.subMatches(1))
        sVal = CleanUpColumn(sVal)
        sField = RegEx.Replace(sField, sVal)
    Next

    RegEx.Pattern = "\$([0-9]+)"
    Set aMatches = RegEx.Execute(sField)
    For Each sMatch In aMatches
        sString = sMatch.subMatches(0)
        iCol = Int(sString)
        sCol = Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").Cells(iRow + 1, iCol).Value
        sFormat = Workbooks("MARC.xlam").Worksheets("Scratch").Range("A2").Cells(iRow + 1, iCol).NumberFormat
        If StrComp(sFormat, "General") <> 0 Then
            sCol = Format(sCol, sFormat)
        End If
        sCol = CleanUpColumn(sCol)
        RegEx.Pattern = "\" & sMatch & "(?=[^0-9])"
        sField = RegEx.Replace(sField, sCol)
        RegEx.Pattern = "\" & sMatch & "$"
        sField = RegEx.Replace(sField, sCol)
    Next

    RegEx.Pattern = "\$D"
    sField = RegEx.Replace(sField, Format(Date, "yymmdd"))
    
    sField = Replace(sField, ChrW(-257), "")
    sField = Replace(sField, "{$}", "$")
    
    RegEx.Pattern = "{=([^}]+)}"
    Set aMatches = RegEx.Execute(sField)
    While aMatches.Count > 0
        sFormula = aMatches(aMatches.Count - 1).subMatches(0)
        sResult = Application.Evaluate(sFormula)
        sField = Replace(sField, "{=" & sFormula & "}", sResult)
        Set aMatches = RegEx.Execute(sField)
    Wend
    
    ResolveVariables = sField
    Exit Function
    
End Function

Function CleanUpColumn(ByVal sStr As String) As String
    sStr = Trim(sStr)
    If Right(sStr, 2) = "_)" Then
        sStr = Left(sStr, Len(sStr) - 2)
    End If
    If Right(sStr, 1) = "_" Then
        sStr = Left(sStr, Len(sStr) - 1)
    End If
    CleanUpColumn = sStr
End Function


Sub ConvertToMARC()
    On Error GoTo ErrHandler:
    Dim RegEx
    Set RegEx = CreateObject("vbscript.regexp")
    With RegEx
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "\.[^.]+$"
        .Global = True
    End With

    With MARCWindow.PreviewListBox
        bHasSelected = False
        iNumSelected = 0
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                bHasSelected = True
                iNumSelected = iNumSelected + 1
            End If
        Next
        
        If iNumSelected = 1 Then
            iChoice = MsgBox("Only row is selected.  Do you want to select all rows?", vbYesNo, "Question")
            If iChoice = vbYes Then
                bHasSelected = False
            End If
        End If
        
        bEvents = False
        If Not bHasSelected Then
            For i = 0 To .ListCount - 1
                .Selected(i) = True
            Next
        End If
        bEvents = True
        
        Dim sAllRecords As String
        sAllRecords = ""
        
        If MARCWindow.TitleRowCheckBox.Value = True Then
            .Selected(0) = False
        End If
        
        
        sDefaultFileName = ActiveWorkbook.FullName
        sDefaultFileName = RegEx.Replace(sDefaultFileName, "")
        sDefaultFileName = sDefaultFileName & "_" & ActiveSheet.Name & ".mrc"
        Application.DisplayAlerts = True
        sFileSaveName = Application.GetSaveAsFilename( _
            InitialFileName:=sDefaultFileName, _
            FileFilter:="MARC records (*.mrc), *.mrc")
        If sFileSaveName = False Then
            Exit Sub
        End If
    
        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(sFileSaveName) Then
            iOverwrite = MsgBox("Overwrite the existing file '" & sFileSaveName & "'?", vbOKCancel)
            If iOverwrite = vbCancel Then
                Exit Sub
            End If
        End If
        
        
        iCount = 0
        
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                iCount = iCount + 1
                Dim sLeader As String, sDirectory As String
                Dim sFieldValue As String, sRecord As String
                sLeader = ""
                sDirectory = ""
                iPos = 0
                sRecord = ""
                sField = ""
                sSeq = ""
                For j = 0 To MARCWindow.ProfileListBox.ListCount - 1
                    sField = MARCWindow.ProfileListBox.List(j, 0)
                    sInd1 = Left(MARCWindow.ProfileListBox.List(j, 2), 1)
                    If Len(sInd1) = 0 Then
                        sInd1 = " "
                    End If
                    sInd2 = Left(MARCWindow.ProfileListBox.List(j, 3), 1)
                    If Len(sInd2) = 0 Then
                        sInd2 = " "
                    End If
                    Dim sValue As String
                    sValue = MARCWindow.ProfileListBox.List(j, 4)
                    If StrComp(sField, "000") = 0 Then
                        sLeader = ResolveVariables(sValue, Int(i))
                        sLeader = Replace(sLeader, "#", " ")
                    Else
                        If StrComp(Left(sField, 2), "00") = 0 Then
                            sValue = Replace(sValue, "#", " ")
                        End If
                        sValue = Replace(sValue, "|", Chr(31))
                        sValue = ResolveVariables(sValue, Int(i))
                        If Len(sValue) > 0 Then
                            If StrComp(Left(sField, 2), "00") = 0 Then
                                sValue = Chr(30) & sValue
                            Else
                                sValue = Chr(30) & sInd1 & sInd2 & sValue
                            End If
                            iLen = UTF8Size(sValue)
                            sDirectory = sDirectory & sField & Format(iLen, "0000") & Format(iPos, "00000")
                            iPos = iPos + iLen
                            sRecord = sRecord & sValue
                        End If
                    End If
                Next
                sLeader = ResolveVariablesLeader(sLeader, sDirectory, sRecord)
                sAllRecords = sAllRecords & sLeader & sDirectory & sRecord & Chr(30) & Chr(29)
            End If
        Next
    End With
     
    Dim UTFStream As Object
    Set UTFStream = CreateObject("ADODB.stream")
    UTFStream.Type = 2
    UTFStream.Mode = 3
    UTFStream.Charset = "UTF-8"
    UTFStream.LineSeparator = 10
    UTFStream.Open
    UTFStream.WriteText sAllRecords, 0

    UTFStream.Position = 3 'skip BOM

    Dim BinaryStream As Object
    Set BinaryStream = CreateObject("ADODB.stream")
    BinaryStream.Type = 1
    BinaryStream.Mode = 3
    BinaryStream.Open

    'Strips BOM (first 3 bytes)
    UTFStream.CopyTo BinaryStream

    UTFStream.Flush
    UTFStream.Close

    BinaryStream.SaveToFile sFileSaveName, 2
    BinaryStream.Flush
    BinaryStream.Close
    
    MsgBox ("Conversion was successful!" & Chr(10) & iCount & " records created.")
    Exit Sub
ErrHandler:
    iResult = MsgBox("An Error Occurred while trying to apply the pattern '" & sValue _
        & "' to record #" & i + 1 & ". No records generated.", vbExclamation, "Error")
    Exit Sub
End Sub

Sub UpdateMARCPreview()
    On Error GoTo ErrHandler:
    Dim RegEx
    Set RegEx = CreateObject("vbscript.regexp")
    With RegEx
        .IgnoreCase = True
        .MultiLine = False
        .Global = True
    End With

    For i = 0 To MARCWindow.MARCPreviewBox.ListCount - 1
        MARCWindow.MARCPreviewBox.RemoveItem (0)
    Next


    iSelected = 0
    For i = 0 To MARCWindow.PreviewListBox.ListCount - 1
        If MARCWindow.PreviewListBox.Selected(i) Then
            iSelected = i
            Exit For
        End If
    Next
    If iSelected = 0 And MARCWindow.TitleRowCheckBox.Value = True Then
        iSelected = 1
    End If
    
    iLeaderIndex = -1
    Dim sLeader As String
    sLeader = ""
    Dim sDirectory As String
    sDirectory = ""
    iPos = 0
    Dim sRecord As String
    sRecord = ""
    
    j = 0
    
    For i = 0 To MARCWindow.ProfileListBox.ListCount - 1
        Dim sValue As String
        sValue = MARCWindow.ProfileListBox.List(i, 4)
        If StrComp(MARCWindow.ProfileListBox.List(i, 0), "000") = 0 Then
            iLeaderIndex = i
            sLeader = ResolveVariables(sValue, Int(iSelected))
            MARCWindow.MARCPreviewBox.AddItem
            MARCWindow.MARCPreviewBox.List(j, 0) = "000"
            j = j + 1
        Else
            sValue = ResolveVariables(sValue, Int(iSelected))
            RegEx.Pattern = "(.)\|"
            sValue = RegEx.Replace(sValue, "$1 |")
            
            RegEx.Pattern = "\|(.)"
            sValue = RegEx.Replace(sValue, "|$1 ")

            If Len(sValue) > 0 Then
                iLen = UTF8Size(sValue)
                sDirectory = sDirectory & sField & Format(iLen, "0000") & Format(iPos, "00000")
                iPos = iPos + iLen
                sRecord = sRecord & sValue
                MARCWindow.MARCPreviewBox.AddItem
                MARCWindow.MARCPreviewBox.List(j, 0) = MARCWindow.ProfileListBox.List(i, 0)
                MARCWindow.MARCPreviewBox.List(j, 1) = Left(MARCWindow.ProfileListBox.List(i, 2), 1)
                MARCWindow.MARCPreviewBox.List(j, 2) = Left(MARCWindow.ProfileListBox.List(i, 3), 1)
                MARCWindow.MARCPreviewBox.List(j, 3) = sValue
                j = j + 1
            End If
        End If
    Next
    If iLeaderIndex > -1 Then
        sValue = ResolveVariablesLeader(sLeader, sDirectory, sRecord)
        MARCWindow.MARCPreviewBox.List(iLeaderIndex, 3) = sValue
    End If
    Exit Sub
ErrHandler:
    If bEvents Then
        iResult = MsgBox("An Error Occurred while trying to apply the pattern '" & sValue _
            & "' to record #" & iSelected + 1, vbExclamation, "Error")
    End If
    Resume Next
End Sub


