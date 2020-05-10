Attribute VB_Name = "MergeAndSavePDFs"
' Mail Merge PDF Splitter version 1.06
' Author JNCressey
' Ensure reference to "Microsoft VBScript Regular Expressions 5.5" is checked from Tools>References...
    Public ConnectionFailed As Boolean
    
    
Sub Main()
' summon user dialoge to choose data source
    ' if no file chosen -> exit sub
    ' if .xlsx file not chosen -> exit sub
' summon user dialoge to choose save folder
    ' if folder not chosen -> exit sub
' summon user dialogue to choose group number/range
    ' if no group number chosen -> exit sub
' give user oportunity to cancel <- prompt with file/folder addresses, and group number
    'if canceled -> exit sub
' MailMerge.OpenDataSource...
' Freeze application view
' go to first record
' if prieview is off -> toggle preview
' run mainloop
' run finish
'''''''''''''''''''''''''''''''''''

    Dim DataSource As String
    Dim SaveFolder As String
    Dim GoAhead As Integer
    Dim Group As Integer
    ConnectionFailed = False
    
' summon user dialoge to choose data source
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select your data source."
        .Show
        If .SelectedItems.Count = 0 Then
    ' if no file chosen -> exit sub
            MsgBox "No data source chosen." + vbCrLf + "Procedure halted."
            Exit Sub
        ElseIf Right(.SelectedItems(1), 5) <> ".xlsx" Then
    ' if chosen file is not .xlsx -> exit sub
            MsgBox "Data source need to be an Excel file. (.xlsx)" + vbCrLf _
            + "You picked " + .SelectedItems(1) + vbCrLf _
            + "Procedure halted."
            Exit Sub
        Else
            DataSource = .SelectedItems(1)
        End If
    End With
    
' summon user dialoge to choose save folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select your save folder."
        .Show
        If .SelectedItems.Count = 0 Then
    ' if folder not chosen -> exit sub
            MsgBox "No folder chosen. Procedure halted."
            Exit Sub
        Else
            SaveFolder = .SelectedItems(1)
        End If
    End With
' summon user dialogue to choose group number/range
    Dim GroupInput As String
    Dim AcceptedGroupRange As Boolean
    AcceptedGroupRange = False
    Do While Not AcceptedGroupRange
        GroupInput = InputBox("Which group of records?" + vbCrLf + vbCrLf _
        + "(You may type just an integer or a range like 2-4, 6, 9)")
        If GroupInput = "" Then
        ' if no group number chosen -> exit sub
            MsgBox "No group chosen. Procedure halted."
            Exit Sub
        ElseIf IsMultiRange(GroupInput) Then
            AcceptedGroupRange = True
            Dim GroupRange As String
            GroupRange = MultiRangeTrimWhiteSpace(GroupInput)
        Else
            MsgBox ("Invalid range." + vbCrLf + vbCrLf _
            + "You may type just an integer or a range like 2-4, 6, 9")
        End If
    Loop
' give user oportunity to cancel <- prompt with file/folder addresses
    
    GoAhead = MsgBox("Data Source: " + DataSource + vbCrLf + vbCrLf _
    + "Save Folder: " + SaveFolder + vbCrLf + vbCrLf _
    + "Group(s): " + GroupRange + vbCrLf + vbCrLf _
    + "Continue?", vbYesNo)
    If GoAhead = vbNo Then
    'if canceled -> exit sub
        Exit Sub
    End If
    

' MailMerge.OpenDataSource...
    OpenDataSource DataSource
    If ConnectionFailed Then
        Exit Sub
    End If
' Freeze application view
    Application.ScreenUpdating = False
    
' go to first record
    ActiveDocument.MailMerge.DataSource.ActiveRecord = wdFirstRecord
    
' if prieview is off -> toggle preview
    If ActiveDocument.MailMerge.ViewMailMergeFieldCodes = -1 Then
        ActiveDocument.MailMerge.ViewMailMergeFieldCodes = wdToggle
    End If

' run mainloop
    MainLoop DataSource, SaveFolder, GroupRange
    
' run finish
    Finish

End Sub


Sub OpenDataSource(DataSource As String)
'
' Uses Select-Recipients and picks the data Excel document
'
'
    On Error GoTo handler
    ActiveDocument.MailMerge.OpenDataSource Name:=DataSource, SQLStatement:="SELECT * FROM `MailMergeData$`"
    Exit Sub
handler:
    On Error GoTo 0
    Dim TryAgain As Integer
    TryAgain = MsgBox("Data source doesn't contain a sheet named 'MailMergeData'." + vbCrLf + vbCrLf + "Do you want the program to try again without assuming the name of the sheet?", vbYesNo)
    If TryAgain = vbYes Then
        ActiveDocument.MailMerge.OpenDataSource Name:=DataSource
    Else
        ConnectionFailed = True
    End If
End Sub


Sub MainLoop(DataSource As String, SaveFolder As String, GroupRange As String)
' for each record
    ' if group number correct, then
        ' save pdf <- FileName field, SaveFolder
    ' if not last record -> move to next record
' end for
'''''''''''''''''''''''''''''''''''

' for each record
    Dim RecordIndex As Integer
    For RecordIndex = 1 To ActiveDocument.MailMerge.DataSource.RecordCount
    ' if group number correct, then
    If NumInRange(ActiveDocument.MailMerge.DataSource.DataFields("Group").Value, GroupRange) Then
        ' save pdf <- FileName field, SaveFolder
        SavePDF ActiveDocument.MailMerge.DataSource.DataFields("FileName").Value, SaveFolder
    End If
    
    ' if not last record -> move to next record
        If RecordIndex < ActiveDocument.MailMerge.DataSource.RecordCount Then
            ActiveDocument.MailMerge.DataSource.ActiveRecord = wdNextRecord
        End If
    Next RecordIndex
' end for
End Sub

Function IsMultiRange(CandidateString As String) As Boolean
' checks if CandidateString is of the form of a range selection like what is used for selecting pages on a print dialog
' https://stackoverflow.com/a/22542835
    Dim MultiRangePattern As String
    MultiRangePattern = "^\s*\d+(?:\s*-\s*\d+)?(?:\s*,\s*\d+(?:\s*-\s*\d+)?)*\s*$" ' composed using https://regex101.com/
    Dim regEx As New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = MultiRangePattern
    End With
    If regEx.Test(CandidateString) Then
        IsMultiRange = True
        Exit Function
    Else
        IsMultiRange = False
        Exit Function
    End If
End Function

Function MultiRangeTrimWhiteSpace(InputString As String) As String
'cleans up the multirange string so that there are no spaces expect for just one after each comma.
    MultiRangeTrimWhiteSpace = RegExReplace(RegExReplace(InputString, "\s*", ""), ",", ", ")
End Function

Function RegExReplace(InputString As String, strPattern As String, strReplace As String) As String
    Dim regEx As New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    RegExReplace = regEx.Replace(InputString, strReplace)
End Function

Sub SavePDF(FileName As String, SaveFolder As String)
'
' Saves current record as pdf
'
'

Dim FullFileName As String
FullFileName = SaveFolder + "\" + FileName + ".pdf"

    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        FullFileName, ExportFormat:= _
        wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    
End Sub


Sub SaveDOCX(FileName As String, SaveFolder As String)
' saves the current record as docx
' when first opened it will try to connect to data. select no then save, to stop it doing that.

Dim FullFileName As String
FullFileName = SaveFolder + "\" + FileName + ".docx"

    ActiveDocument.SaveAs FileName:=FullFileName, FileFormat:=wdFormatDocumentDefault
    
End Sub

Function NumInRange(Num As Integer, MultiRange As String) As Boolean
    ' takes a user written range in a form similar to 'page range' you get on print dialogues. eg "2-4, 5, 9"
    ' (must have no spaces beside the dash, and must have a space after each comma)
    ' tests is Num is in that custom range
    NumInRange = False 'initialise as False, to turn to true when found a valid inclusion
    
    Dim RangePieces() As String
    RangePieces = Split(MultiRange, ", ")
    
    Dim PieceIndex As Integer
    Dim TempPiece() As String
    For PieceIndex = 0 To UBound(RangePieces)
        TempPiece = Split(RangePieces(PieceIndex), "-")
        ' TempPiece is now {x1} or {x1,x2}
        If UBound(TempPiece) = 0 Then
            'TempPiece is {x1}
            If Num = CInt(TempPiece(0)) Then
                NumInRange = True
                Exit For
            End If
        Else
            'TempPiece is {x1,x2}
            If Num >= CInt(TempPiece(0)) And Num <= CInt(TempPiece(1)) Then
                NumInRange = True
                Exit For
            End If
        End If
    Next PieceIndex
    
    
End Function



Sub Finish()
' go to first record
' toggle preview
' unfreeze application view
' alert box 'success!'
'''''''''''''''''''''''''''''''''''

' go to first record
    ActiveDocument.MailMerge.DataSource.ActiveRecord = wdFirstRecord
' toggle preview
    ActiveDocument.MailMerge.ViewMailMergeFieldCodes = wdToggle
' unfreeze application view
    Application.ScreenUpdating = True
' alert box 'success!'
    MsgBox "Mail merge complete."
End Sub



