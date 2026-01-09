Attribute VB_Name = "Export"
Option Explicit

' =========================
' PUBLIC: reusable PDF export
' =========================
Public Sub ExportRangeToPdf(ByVal ws As Worksheet, ByVal rng As Range, ByVal filePath As String)
    On Error GoTo CleanFail

    Dim oldPrintArea As String
    oldPrintArea = ws.PageSetup.PrintArea

    Dim oldOrientation As XlPageOrientation
    oldOrientation = ws.PageSetup.Orientation

    Dim oldZoom As Variant
    oldZoom = ws.PageSetup.Zoom

    Dim oldFitWide As Variant, oldFitTall As Variant
    oldFitWide = ws.PageSetup.FitToPagesWide
    oldFitTall = ws.PageSetup.FitToPagesTall

    Dim oldCenterH As Boolean, oldCenterV As Boolean
    oldCenterH = ws.PageSetup.CenterHorizontally
    oldCenterV = ws.PageSetup.CenterVertically

    Dim oldPaper As Variant
    oldPaper = ws.PageSetup.PaperSize

    ' Margins backup
    Dim oldLM As Double, oldRM As Double, oldTM As Double, oldBM As Double, oldHM As Double, oldFM As Double
    oldLM = ws.PageSetup.LeftMargin
    oldRM = ws.PageSetup.RightMargin
    oldTM = ws.PageSetup.TopMargin
    oldBM = ws.PageSetup.BottomMargin
    oldHM = ws.PageSetup.HeaderMargin
    oldFM = ws.PageSetup.FooterMargin

    With ws.PageSetup
        .PrintArea = rng.Address
        .PaperSize = xlPaperA4

        ' Choose orientation based on aspect ratio
        If rng.Width > rng.Height Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If

        ' Minimal safe margins (in inches -> points)
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.2)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)

        .CenterHorizontally = True
        .CenterVertically = True

        ' Fit to single page
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

CleanExit:
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = oldPrintArea
        .Orientation = oldOrientation
        .Zoom = oldZoom
        .FitToPagesWide = oldFitWide
        .FitToPagesTall = oldFitTall
        .CenterHorizontally = oldCenterH
        .CenterVertically = oldCenterV
        .PaperSize = oldPaper

        .LeftMargin = oldLM
        .RightMargin = oldRM
        .TopMargin = oldTM
        .BottomMargin = oldBM
        .HeaderMargin = oldHM
        .FooterMargin = oldFM
    End With
    On Error GoTo 0
    Exit Sub

CleanFail:
    Resume CleanExit
End Sub


' =========================
' PUBLIC: your Pokedex export (Mon -> PDF)
' =========================
Public Sub ExportMon()
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = Pokedex   ' CodeName

    Dim exportRng As Range
    Set exportRng = ws.Range("B4:AF34")

    Dim monName As String
    monName = CleanFileName(CStr(ws.Range("PKMN_DEX").Value2))
    If Len(monName) = 0 Then monName = "Pokemon"

    If Not ExportWithPrompt(ws, exportRng, monName & ".pdf", "ExportMon") Then Exit Sub

    Exit Sub

CleanFail:
    MsgBox "ExportMon failed:" & vbCrLf & Err.Description, vbExclamation, "ExportMon"
End Sub

Public Sub ExportTypeChart()
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = TypeChart  ' CodeName

    Dim exportRng As Range
    Set exportRng = ws.Range("B3:Y28")

    If Not ExportWithPrompt(ws, exportRng, "TypeChart.pdf", "ExportTypeChart") Then Exit Sub

    Exit Sub

CleanFail:
    MsgBox "ExportTypeChart failed:" & vbCrLf & Err.Description, vbExclamation, "ExportTypeChart"
End Sub

Public Sub ExportNatures()
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = Naturedex   ' CodeName

    Dim exportRng As Range
    Set exportRng = ws.Range("B4:L35")

    Dim baseName As String
    baseName = "Natures"

    If Not ExportWithPrompt(ws, exportRng, baseName & ".pdf", "ExportNatures") Then Exit Sub

    Exit Sub

CleanFail:
    MsgBox "ExportNatures failed:" & vbCrLf & Err.Description, vbExclamation, "ExportNatures"
End Sub


' =========================
' HELPERS
' =========================
Private Function CleanFileName(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function

    Dim invalidChars As Variant, i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(invalidChars) To UBound(invalidChars)
        s = Replace$(s, CStr(invalidChars(i)), "_")
    Next i

    Do While Right$(s, 1) = "." Or Right$(s, 1) = " "
        s = Left$(s, Len(s) - 1)
        If Len(s) = 0 Then Exit Function
    Loop

    CleanFileName = s
End Function

Private Function ExportWithPrompt( _
    ByVal ws As Worksheet, _
    ByVal exportRng As Range, _
    ByVal defaultFileName As String, _
    ByVal dialogTitle As String _
) As Boolean

    Dim filePath As Variant
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFileName, _
        FileFilter:="PDF files (*.pdf), *.pdf" _
    )
    If VarType(filePath) = vbBoolean And filePath = False Then Exit Function

    Dim resolvedPath As String
    resolvedPath = CStr(filePath)
    If Len(resolvedPath) = 0 Then Exit Function

    If Len(Dir$(resolvedPath)) > 0 Then
        Dim overwriteResp As VbMsgBoxResult
        overwriteResp = MsgBox( _
            "File already exists. Overwrite?", _
            vbExclamation + vbYesNo + vbDefaultButton2, _
            dialogTitle _
        )
        If overwriteResp <> vbYes Then Exit Function
    End If

    ExportRangeToPdf ws, exportRng, resolvedPath
    ExportWithPrompt = True

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "PDF exported successfully." & vbCrLf & vbCrLf & _
        "Do you wish to open it now?", _
        vbQuestion + vbYesNo, _
        dialogTitle _
    )

    If resp = vbYes Then
        ThisWorkbook.FollowHyperlink Address:=resolvedPath, NewWindow:=True
    End If
End Function


