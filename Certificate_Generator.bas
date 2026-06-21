Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateCertificatesPDF()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long

    Dim TemplatePath As String
    Dim SaveFolder As String

    Dim CandidateID As String
    Dim CandidateName As String
    Dim CertDate As String
    Dim CertificateID As String

    Set ws = ActiveSheet

    'Select Word Template
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Certificate Template"
        .Filters.Clear
        .Filters.Add "Word Files", "*.docx;*.docm"
        
        If .Show <> -1 Then Exit Sub
        
        TemplatePath = .SelectedItems(1)
    End With

    'Select Output Folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Output Folder"
        
        If .Show <> -1 Then Exit Sub
        
        SaveFolder = .SelectedItems(1)
    End With

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To LastRow

        CandidateID = ws.Cells(i, 1).Value
        CandidateName = ws.Cells(i, 2).Value
        CertDate = ws.Cells(i, 3).Text
        CertificateID = ws.Cells(i, 4).Value

        Set wdDoc = wdApp.Documents.Open(TemplatePath)

        ReplaceEverywhere wdDoc, "<<CANDIDATE NAME>>", CandidateName
        ReplaceEverywhere wdDoc, "<<DATE>>", CertDate
        ReplaceEverywhere wdDoc, "<<CERTIFICATE ID>>", CertificateID

        wdDoc.ExportAsFixedFormat _
            OutputFileName:=SaveFolder & "\" & CandidateID & "_Certificate.pdf", _
            ExportFormat:=17

        wdDoc.Close False

    Next i

    wdApp.Quit

    Set wdDoc = Nothing
    Set wdApp = Nothing

    MsgBox "All Certificates Generated Successfully!", vbInformation

End Sub

Sub ReplaceEverywhere(doc As Object, FindTxt As String, ReplaceTxt As String)

    Dim shp As Object
    Dim sec As Object
    Dim hdr As Object
    Dim ftr As Object

    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = FindTxt
        .Replacement.Text = ReplaceTxt
        .Wrap = 1
        .Execute Replace:=2
    End With

    For Each shp In doc.Shapes
        On Error Resume Next
        If shp.TextFrame.HasText Then
            shp.TextFrame.TextRange.Text = _
            Replace(shp.TextFrame.TextRange.Text, FindTxt, ReplaceTxt)
        End If
        On Error GoTo 0
    Next shp

    For Each sec In doc.Sections

        For Each hdr In sec.Headers
            For Each shp In hdr.Range.ShapeRange
                On Error Resume Next
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Text = _
                    Replace(shp.TextFrame.TextRange.Text, FindTxt, ReplaceTxt)
                End If
                On Error GoTo 0
            Next shp
        Next hdr

        For Each ftr In sec.Footers
            For Each shp In ftr.Range.ShapeRange
                On Error Resume Next
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Text = _
                    Replace(shp.TextFrame.TextRange.Text, FindTxt, ReplaceTxt)
                End If
                On Error GoTo 0
            Next shp
        Next ftr

    Next sec

End Sub

