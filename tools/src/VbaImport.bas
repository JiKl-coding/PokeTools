Attribute VB_Name = "VbaImport"
Option Explicit
Private Const IGNORE_IMPORT_MODULES As String = "vbaImport,vbaExport"

' === Public API ===
' Imports all VBA from tools/src/*
' Used to be able to develop in editor (vsCode)
' Does not import from internal
Public Sub ImportAllVba()
    Dim srcPath As String
    srcPath = ThisWorkbook.path & Application.PathSeparator & "src" & Application.PathSeparator

    If Dir(srcPath, vbDirectory) = vbNullString Then
        MsgBox "Folder not found: " & srcPath, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    ImportFolderIntoThisWorkbook srcPath

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "ImportAllVba failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' === Core ===
Private Sub ImportFolderIntoThisWorkbook(ByVal folderPath As String)
    Dim fileName As String
    fileName = Dir(folderPath & "*.*")

    Do While fileName <> vbNullString
        Dim fullPath As String
        fullPath = folderPath & fileName

        If (GetAttr(fullPath) And vbDirectory) = 0 Then
            Dim ext As String
            ext = LCase$(GetFileExtension(fileName))

            ' Skip txt and everything else except bas/cls/frm
            Select Case ext
                Case "bas", "cls", "frm"
                    ImportOneComponent fullPath
                Case Else
                    ' ignore
            End Select
        End If

        fileName = Dir()
    Loop
End Sub

Private Sub ImportOneComponent(ByVal filePath As String)
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim compName As String
    compName = GuessComponentName(filePath)

    ' 1) Skip ignored modules (keep them stable inside the workbook)
    If IsNameInCsvList(compName, IGNORE_IMPORT_MODULES) Then
        ' Optional: also delete any duplicates like vbaImport1, vbaImport2
        DeleteComponentsByPrefix vbProj, compName
        Exit Sub
    End If

    ' 2) Delete exact match (if any) but never touch document modules
    DeleteComponentIfExists vbProj, compName

    ' 3) Also delete duplicates by prefix (Module1, Module11...) created by past failed imports
    DeleteComponentsByPrefix vbProj, compName

    ' 4) Import fresh
    vbProj.VBComponents.Import filePath
End Sub

' === Helpers ===

' Best effort:
' - for .bas/.cls it's usually "Attribute VB_Name = ""X"""
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
        ' Example: Attribute VB_Name = "Module1"
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

Private Sub DeleteComponentIfExists(ByVal vbProj As Object, ByVal compName As String)
    Dim vbComp As Object
    For Each vbComp In vbProj.VBComponents
        If StrComp(vbComp.name, compName, vbTextCompare) = 0 Then
            ' vbext_ct_Document = 100 (ThisWorkbook + Worksheets)
            If vbComp.Type = 100 Then Exit Sub
            vbProj.VBComponents.Remove vbComp
            Exit Sub
        End If
    Next vbComp
End Sub

Private Function GetFileExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then GetFileExtension = Mid$(fileName, p + 1) Else GetFileExtension = ""
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

' Deletes components like:
'   vbaImport1, vbaImport2...
'   Module11 (if base is Module1) is tricky, but VBIDE duplicates typically append digits.
Private Sub DeleteComponentsByPrefix(ByVal vbProj As Object, ByVal baseName As String)
    Dim i As Long
    Dim vbComp As Object

    ' Iterate backwards (safe removal)
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set vbComp = vbProj.VBComponents(i)

        ' Never touch document modules (ThisWorkbook/Worksheets)
        If vbComp.Type <> 100 Then
            If StrComp(vbComp.name, baseName, vbTextCompare) <> 0 Then
                If IsDuplicateNameOfBase(vbComp.name, baseName) Then
                    vbProj.VBComponents.Remove vbComp
                End If
            End If
        End If
    Next i
End Sub

Private Function IsDuplicateNameOfBase(ByVal candidate As String, ByVal baseName As String) As Boolean
    ' True for baseName + digits, e.g. vbaImport1, vbaImport2
    Dim prefix As String
    prefix = baseName

    If Len(candidate) <= Len(prefix) Then Exit Function
    If StrComp(Left$(candidate, Len(prefix)), prefix, vbTextCompare) <> 0 Then Exit Function

    Dim suffix As String
    suffix = Mid$(candidate, Len(prefix) + 1)

    ' must be digits-only
    IsDuplicateNameOfBase = IsAllDigits(suffix)
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

