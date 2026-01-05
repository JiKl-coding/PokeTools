Attribute VB_Name = "VbaExport"
Option Explicit

' Late-bound constants (so you don't need reference to VBIDE)
Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_CLASSMODULE As Long = 2
Private Const VBEXT_CT_MSFORM As Long = 3
Private Const VBEXT_CT_DOCUMENT As Long = 100

' === Public API ===
Public Sub ExportAllVba()
    Dim msg As String
    msg = "Are you sure you want to export all VBA code to:" & vbCrLf & vbCrLf & _
            "  src\  (bas / cls / frm)" & vbCrLf & _
            "  src\_internal\  (ThisWorkbook and worksheet modules)" & vbCrLf & vbCrLf & _
            "All existing files will be overwritten."

    If MsgBox(msg, vbQuestion + vbYesNo, "Export VBA") <> vbYes Then Exit Sub

    Dim basePath As String, srcPath As String, internalPath As String
    basePath = ThisWorkbook.path & Application.PathSeparator
    srcPath = basePath & "src" & Application.PathSeparator
    internalPath = srcPath & "_internal" & Application.PathSeparator

    EnsureFolderExists srcPath
    EnsureFolderExists internalPath

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim vbComp As Object
    For Each vbComp In vbProj.VBComponents
        ExportOneComponent vbComp, srcPath, internalPath
    Next vbComp

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Done. VBA exported.", vbInformation
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "ExportAllVba failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' === Core ===
Private Sub ExportOneComponent(ByVal vbComp As Object, ByVal srcPath As String, ByVal internalPath As String)
    Dim outDir As String
    Dim outFile As String

    Select Case vbComp.Type
        Case VBEXT_CT_DOCUMENT
            ' ThisWorkbook + Worksheets -> do _internal
            outDir = internalPath
            outFile = outDir & SanitizeFileName(vbComp.name) & ".cls"

        Case VBEXT_CT_STDMODULE
            outDir = srcPath
            outFile = outDir & SanitizeFileName(vbComp.name) & ".bas"

        Case VBEXT_CT_CLASSMODULE
            outDir = srcPath
            outFile = outDir & SanitizeFileName(vbComp.name) & ".cls"

        Case VBEXT_CT_MSFORM
            outDir = srcPath
            outFile = outDir & SanitizeFileName(vbComp.name) & ".frm"

        Case Else
            ' ignore unknown types
            Exit Sub
    End Select

    ' Delete existing file(s) to avoid stale content
    SafeKill outFile

    ' For UserForms there is also .frx; kill it too so it doesn't go stale
    If vbComp.Type = VBEXT_CT_MSFORM Then
        SafeKill Replace$(outFile, ".frm", ".frx")
    End If

    vbComp.Export outFile
End Sub

' === Helpers ===
Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = vbNullString Then
        MkDir folderPath
    End If
End Sub

Private Sub SafeKill(ByVal filePath As String)
    On Error Resume Next
    If Len(Dir(filePath)) > 0 Then Kill filePath
    On Error GoTo 0
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    ' Just in case (VBA component names should be safe anyway)
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next i
    SanitizeFileName = s
End Function


