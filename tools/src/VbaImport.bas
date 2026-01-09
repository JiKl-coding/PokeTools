Attribute VB_Name = "VbaImport"
Option Explicit

' Modules that must NEVER be imported nor deleted
Private Const IGNORE_IMPORT_MODULES As String = "vbaImport,vbaExport"

' VBIDE component types (numeric, to avoid VBIDE reference)
Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_CLASSMODULE As Long = 2
Private Const VBEXT_CT_MSFORM As Long = 3
Private Const VBEXT_CT_DOCUMENT As Long = 100

' ...

' === Public API ===============================================================

' Imports all VBA modules (.bas, .cls, .frm) from /src
' - Deletes existing modules/forms first
' - Never touches ThisWorkbook or Worksheets (document modules)
' - Intended for VS Code / external editor workflow
Public Sub ImportAllVba()

    Dim srcPath As String
    srcPath = ThisWorkbook.path & Application.PathSeparator & "src" & Application.PathSeparator

    If Dir(srcPath, vbDirectory) = vbNullString Then
        MsgBox "Folder not found:" & vbCrLf & srcPath, vbExclamation
        Exit Sub
    End If

    ' --- Destructive action confirmation ---
    Dim msg As String
    msg = "Are you sure you want to IMPORT all VBA modules from:" & vbCrLf & vbCrLf & _
          "  " & srcPath & vbCrLf & vbCrLf & _
          "This will DELETE and REPLACE existing:" & vbCrLf & _
          "  ? Standard modules (.bas)" & vbCrLf & _
          "  ? Class modules (.cls)" & vbCrLf & _
          "  ? UserForms (.frm)" & vbCrLf & vbCrLf & _
          "Ignored modules:" & vbCrLf & _
          "  " & IGNORE_IMPORT_MODULES & vbCrLf & vbCrLf & _
          "Continue?"

    If MsgBox(msg, vbQuestion + vbYesNo, "Import VBA") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    ' 1) Remove existing importable components
    DeleteAllImportableComponents ThisWorkbook.VBProject

    ' 2) Import fresh modules from /src
    ImportFolderIntoThisWorkbook srcPath
    
    ' Clean possible duplicates if they ever existed from older runs
    CleanupIgnoredDuplicates ThisWorkbook.VBProject

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "ImportAllVba failed:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
End Sub


' === Core ====================================================================

' Deletes all existing .bas / .cls / .frm components
' except those explicitly ignored
Private Sub DeleteAllImportableComponents(ByVal vbProj As Object)

    Dim i As Long
    Dim vbComp As Object

    ' Iterate backwards (safe removal)
    For i = vbProj.VBComponents.count To 1 Step -1
        Set vbComp = vbProj.VBComponents(i)

        Select Case vbComp.Type
            Case VBEXT_CT_STDMODULE, VBEXT_CT_CLASSMODULE, VBEXT_CT_MSFORM

                ' Skip protected modules
                If IsNameInCsvList(vbComp.name, IGNORE_IMPORT_MODULES) Then
                    ' Also clean possible numeric duplicates (vbaImport1, vbaImport2...)
                    DeleteComponentsByPrefix vbProj, vbComp.name
                Else
                    vbProj.VBComponents.Remove vbComp
                End If

            Case Else
                ' Document modules (ThisWorkbook / Worksheets) are ignored by design
        End Select
    Next i
End Sub


' Imports all .bas / .cls / .frm files from a folder
' Skips ignored modules by VB_Name (or filename fallback)
Private Sub ImportFolderIntoThisWorkbook(ByVal folderPath As String)

    Dim fileName As String
    fileName = Dir(folderPath & "*.*")

    Do While fileName <> vbNullString

        Dim fullPath As String
        fullPath = folderPath & fileName

        If (GetAttr(fullPath) And vbDirectory) = 0 Then

            Dim ext As String
            ext = LCase$(GetFileExtension(fileName))

            Select Case ext
                Case "bas", "cls", "frm"

                    Dim compName As String
                    compName = GuessComponentName(fullPath)

                    ' Skip protected modules completely (prevents VbaImport1/VbaExport1)
                    If Not IsNameInCsvList(compName, IGNORE_IMPORT_MODULES) Then
                        RemoveComponentByExactName ThisWorkbook.VBProject, compName
                        ThisWorkbook.VBProject.VBComponents.Import fullPath
                    End If

            End Select
        End If

        fileName = Dir()
    Loop
End Sub

' Removes VbaImport1/VbaExport1 if they exist (digits suffix)
Private Sub CleanupIgnoredDuplicates(ByVal vbProj As Object)
    Dim arr() As String, i As Long
    arr = Split(IGNORE_IMPORT_MODULES, ",")

    For i = LBound(arr) To UBound(arr)
        DeleteComponentsByPrefix vbProj, Trim$(arr(i))
    Next i
End Sub


' === Helpers =================================================================

Private Function GetFileExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then GetFileExtension = Mid$(fileName, p + 1)
End Function


Private Function IsNameInCsvList(ByVal name As String, ByVal csv As String) As Boolean
    Dim arr() As String, i As Long
    arr = Split(LCase$(csv), ",")
    For i = LBound(arr) To UBound(arr)
        If LCase$(Trim$(arr(i))) = LCase$(Trim$(name)) Then
            IsNameInCsvList = True
            Exit Function
        End If
    Next i
End Function


' Removes duplicate components created by failed imports:
'   e.g. vbaImport1, vbaImport2...
Private Sub DeleteComponentsByPrefix(ByVal vbProj As Object, ByVal baseName As String)

    Dim i As Long
    Dim vbComp As Object

    For i = vbProj.VBComponents.count To 1 Step -1
        Set vbComp = vbProj.VBComponents(i)

        If vbComp.Type <> VBEXT_CT_DOCUMENT Then
            If IsDuplicateNameOfBase(vbComp.name, baseName) Then
                vbProj.VBComponents.Remove vbComp
            End If
        End If
    Next i
End Sub


Private Function IsDuplicateNameOfBase(ByVal candidate As String, ByVal baseName As String) As Boolean

    If Len(candidate) <= Len(baseName) Then Exit Function
    If StrComp(Left$(candidate, Len(baseName)), baseName, vbTextCompare) <> 0 Then Exit Function

    IsDuplicateNameOfBase = IsAllDigits(Mid$(candidate, Len(baseName) + 1))
End Function


Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As Integer
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next i
    IsAllDigits = True
End Function

' Best effort:
' - for .bas/.cls/.frm it's usually: Attribute VB_Name = "X"
' - fallback: filename without extension
Private Function GuessComponentName(ByVal filePath As String) As String
    Dim nameFromAttr As String
    nameFromAttr = ReadVbNameAttribute(filePath)

    If Len(nameFromAttr) > 0 Then
        GuessComponentName = nameFromAttr
    Else
        GuessComponentName = StripExtension(GetFileName(filePath))
    End If
End Function

Private Function ReadVbNameAttribute(ByVal filePath As String) As String
    On Error GoTo Fail

    Dim f As Integer: f = FreeFile
    Open filePath For Input As #f

    Dim line As String
    Do While Not EOF(f)
        Line Input #f, line

        If InStr(1, line, "Attribute VB_Name", vbTextCompare) > 0 Then
            Dim p1 As Long, p2 As Long
            p1 = InStr(1, line, """")
            If p1 > 0 Then
                p2 = InStr(p1 + 1, line, """")
                If p2 > p1 Then
                    ReadVbNameAttribute = Mid$(line, p1 + 1, p2 - p1 - 1)
                    Close #f
                    Exit Function
                End If
            End If
        End If
    Loop

    Close #f
Fail:
    ReadVbNameAttribute = vbNullString
    On Error Resume Next
    Close #f
End Function

Private Function StripExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then StripExtension = Left$(fileName, p - 1) Else StripExtension = fileName
End Function

Private Function GetFileName(ByVal filePath As String) As String
    Dim p As Long
    p = InStrRev(filePath, Application.PathSeparator)
    If p > 0 Then GetFileName = Mid$(filePath, p + 1) Else GetFileName = filePath
End Function

' Removes an existing component whose name matches targetName (case-insensitive)
' Useful to prevent Excel from creating Function1 / Module1 duplicates on import
Private Sub RemoveComponentByExactName(ByVal vbProj As Object, ByVal targetName As String)

    If Len(targetName) = 0 Then Exit Sub

    Dim i As Long
    Dim vbComp As Object

    For i = vbProj.VBComponents.count To 1 Step -1
        Set vbComp = vbProj.VBComponents(i)

        If vbComp.Type <> VBEXT_CT_DOCUMENT Then
            If StrComp(vbComp.name, targetName, vbTextCompare) = 0 Then
                vbProj.VBComponents.Remove vbComp
                Exit Sub
            End If
        End If
    Next i
End Sub

