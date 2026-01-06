VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Movelist 
   Caption         =   "UserForm1"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18660
   OleObjectBlob   =   "Movelist.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Movelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================
' UserForm: Movelist
'===============================
Option Explicit

' Learnsets: column letter where MOVE NAME is stored (NOT method!)
' Method is in column E and Level is in column F.
Private Const LEARNSETS_MOVE_COL As String = "D"   ' <-- CHANGE this to the correct column letter!

Private pdWB As Workbook





Private Sub lbMoves_Click()

End Sub

Private Sub lbMoves_Change()
    If Me.lbMoves.ListCount = 0 Then
        Me.txtDescription.Visible = False
        Exit Sub
    End If


    Me.txtDescription.Visible = True
    Me.txtDescription.value = Me.lbMoves.List(Me.lbMoves.ListIndex, 6)
End Sub

Private Sub UserForm_Initialize()

    'styling
    Me.BackColor = RGB(204, 0, 0)
    Me.lbMoves.BackColor = RGB(173, 216, 230)

    ' Configure the ListBox (multi-column "table")
    SetupMovesListBox
    
    ' Configure fake header labels above the ListBox (optional)
    SetupHeaderLabels
    
    ' Get cached/hidden pokedata workbook (your implementation)
    Set pdWB = GetPokedataWb()
    
    ' Load moves from Lists!P and fill details from pokedata.xlsx
    LoadMovesWithDetails
    
    ' Set info label + form title
    SetInfoLabel
    Me.lbMoves.ListIndex = 0
End Sub

' -----------------------------
' UI: ListBox
' -----------------------------
Private Sub SetupMovesListBox()
    With Me.lbMoves
        .Clear
        .ColumnCount = 8
        
        ' Column order:
        ' 0 Move | 1 Category | 2 Power | 3 Accuracy | 4 PP | 5 Priority | 6 Description | 7 Method
        .ColumnWidths = _
            "120;70;50;70;40;60;400;120"
        
        .BoundColumn = 0
        .MultiSelect = False
        .IntegralHeight = False
    End With
End Sub

' -----------------------------
' UI: Fake header (labels above ListBox)
' -----------------------------
Private Sub SetupHeaderLabels()
    ' Add labels to the UserForm and name them like this (or adjust names below):
    ' lblHMove, lblHCategory, lblHPower, lblHAccuracy, lblHPP, lblHPriority, lblHDescription, lblHMethod
    
    ' We ignore errors so the form still works even if you didn't add the labels yet.
    On Error Resume Next
    
    Me.lblHMove.Caption = "Move"
    Me.lblHCategory.Caption = "Category"
    Me.lblHPower.Caption = "Power"
    Me.lblHAccuracy.Caption = "Accuracy"
    Me.lblHPP.Caption = "PP"
    Me.lblHPriority.Caption = "Priority"
    Me.lblHDescription.Caption = "Description"
    Me.lblHMethod.Caption = "Method"
    
    On Error GoTo 0
    
    ' Tip:
    ' Align label widths/positions manually in the designer to match ColumnWidths.
End Sub

Private Sub SetInfoLabel()
    Dim pkmnDex As String, game As String
    
    ' Pokedex is the CODE NAME of the sheet in THIS workbook.
    pkmnDex = Trim$(CStr(Pokedex.Range("PKMN_DEX").value))
    game = Trim$(CStr(Pokedex.Range("GAME").value))
    
    Me.lblInfo.Caption = "Movelist of " & pkmnDex & " (" & game & ")"
    Me.Caption = Me.lblInfo.Caption
End Sub

' -----------------------------
' Main load
' -----------------------------
Private Sub LoadMovesWithDetails()
    Dim lastRow As Long, r As Long, i As Long
    Dim moveName As String
    
    Dim pkmnDex As String, game As String, gameNorm As String
    
    Dim dictMoves As Object    ' moveName -> Array(category,power,accuracy,pp,priority,description)
    Dim dictMethod As Object   ' key(pkmn|gameNorm|move) -> "method" or "method[level]"
    
    ' Read current context from Pokedex sheet (code name)
    pkmnDex = Trim$(CStr(Pokedex.Range("PKMN_DEX").value))
    game = Trim$(CStr(Pokedex.Range("GAME").value))
    gameNorm = DexLogic.NormalizeGameVersion(game)
    
    ' Build fast lookup dictionaries
    Set dictMoves = BuildMovesDict(pdWB.Worksheets("Moves"))
    Set dictMethod = BuildLearnsetsMethodDict(pdWB.Worksheets("Learnsets"), pkmnDex, gameNorm)
    
    ' Lists is the CODE NAME of the sheet that contains move list in column P
    lastRow = Lists.Cells(Lists.Rows.Count, "P").End(xlUp).Row
    
    Me.lbMoves.Clear
    
    For r = 2 To lastRow
        moveName = Trim$(CStr(Lists.Cells(r, "P").value))
        If Len(moveName) > 0 Then
            i = Me.lbMoves.ListCount
            Me.lbMoves.AddItem moveName
            
            FillMoveDetails i, moveName, dictMoves, dictMethod, pkmnDex, gameNorm
        End If
    Next r
End Sub

Private Sub FillMoveDetails(ByVal rowIndex As Long, ByVal moveName As String, _
                            ByVal dictMoves As Object, ByVal dictMethod As Object, _
                            ByVal pkmnDex As String, ByVal gameNorm As String)
    Dim arr As Variant
    Dim key As String
    
    ' Fill columns from Moves sheet
    If dictMoves.Exists(LCase$(moveName)) Then
        arr = dictMoves(LCase$(moveName))
        Me.lbMoves.List(rowIndex, 1) = Nz(arr(0)) ' Category (Moves!D)
        Me.lbMoves.List(rowIndex, 2) = Nz(arr(1)) ' Power (Moves!E)
        Me.lbMoves.List(rowIndex, 3) = Nz(arr(2)) ' Accuracy (Moves!F)
        Me.lbMoves.List(rowIndex, 4) = Nz(arr(3)) ' PP (Moves!G)
        Me.lbMoves.List(rowIndex, 5) = Nz(arr(4)) ' Priority (Moves!H)
        Me.lbMoves.List(rowIndex, 6) = Nz(arr(5)) ' Description (Moves!I)
    Else
        ' Move not found in Moves sheet
        Me.lbMoves.List(rowIndex, 1) = "?"
        Me.lbMoves.List(rowIndex, 2) = ""
        Me.lbMoves.List(rowIndex, 3) = ""
        Me.lbMoves.List(rowIndex, 4) = ""
        Me.lbMoves.List(rowIndex, 5) = ""
        Me.lbMoves.List(rowIndex, 6) = ""
    End If
    
    ' Fill Method column from Learnsets (method[level])
    key = LearnKey(pkmnDex, gameNorm, moveName)
    If dictMethod.Exists(key) Then
        Me.lbMoves.List(rowIndex, 7) = CStr(dictMethod(key))
    Else
        Me.lbMoves.List(rowIndex, 7) = "-"
    End If
End Sub

' -----------------------------
' Dictionary builders
' -----------------------------
Private Function BuildMovesDict(ByVal wsMoves As Worksheet) As Object
    ' Moves sheet columns:
    ' B = Move name
    ' D = Category
    ' E = Power
    ' F = Accuracy
    ' G = PP
    ' H = Priority
    ' I = Description
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, r As Long
    Dim nameKey As String
    Dim arr(0 To 5) As Variant
    
    lastRow = wsMoves.Cells(wsMoves.Rows.Count, "B").End(xlUp).Row
    
    For r = 2 To lastRow
        nameKey = LCase$(Trim$(CStr(wsMoves.Cells(r, "B").value)))
        If Len(nameKey) > 0 Then
            arr(0) = wsMoves.Cells(r, "D").value
            arr(1) = wsMoves.Cells(r, "E").value
            arr(2) = wsMoves.Cells(r, "F").value
            arr(3) = wsMoves.Cells(r, "G").value
            arr(4) = wsMoves.Cells(r, "H").value
            arr(5) = wsMoves.Cells(r, "I").value
            dict(nameKey) = arr
        End If
    Next r
    
    Set BuildMovesDict = dict
End Function

Private Function BuildLearnsetsMethodDict(ByVal wsLearn As Worksheet, _
                                         ByVal pkmnDex As String, _
                                         ByVal gameNorm As String) As Object
    ' Learnsets sheet columns:
    ' B = Pokemon
    ' C = Version (needs normalization!)
    ' [LEARNSETS_MOVE_COL] = Move name
    ' E = Method
    ' F = Level
    '
    ' Output:
    ' key(pokemon|gameNorm|move) -> method OR method[level]
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, r As Long
    Dim poke As String, ver As String, verNorm As String
    Dim move As String, method As String, lvl As String
    Dim key As String, outMethod As String
    
    lastRow = wsLearn.Cells(wsLearn.Rows.Count, "B").End(xlUp).Row
    
    For r = 2 To lastRow
        poke = Trim$(CStr(wsLearn.Cells(r, "B").value))
        If Len(poke) = 0 Then GoTo ContinueRow
        
        ver = Trim$(CStr(wsLearn.Cells(r, "C").value))
        verNorm = DexLogic.NormalizeGameVersion(ver)
        
        ' Filter only rows for current pokemon + current game version
        If StrComp(poke, pkmnDex, vbTextCompare) <> 0 Then GoTo ContinueRow
        If StrComp(verNorm, gameNorm, vbTextCompare) <> 0 Then GoTo ContinueRow
        
        move = Trim$(CStr(wsLearn.Cells(r, LEARNSETS_MOVE_COL).value))
        If Len(move) = 0 Then GoTo ContinueRow
        
        method = Trim$(CStr(wsLearn.Cells(r, "E").value))
        lvl = Trim$(CStr(wsLearn.Cells(r, "F").value))
        
        If Len(method) = 0 Then method = "-"
        
        If Len(lvl) > 0 Then
            outMethod = method & " [" & lvl & "]"
        Else
            outMethod = method
        End If
        
        key = LearnKey(pkmnDex, gameNorm, move)
        
        ' If duplicates exist, last one wins (adjust if you want a different rule)
        dict(key) = outMethod
        
ContinueRow:
    Next r
    
    Set BuildLearnsetsMethodDict = dict
End Function

Private Function LearnKey(ByVal pkmnDex As String, ByVal gameNorm As String, ByVal moveName As String) As String
    LearnKey = LCase$(Trim$(pkmnDex)) & "|" & LCase$(Trim$(gameNorm)) & "|" & LCase$(Trim$(moveName))
End Function

' -----------------------------
' Helpers
' -----------------------------
Private Function Nz(ByVal v As Variant) As String
    If IsError(v) Then
        Nz = ""
    Else
        Nz = Trim$(CStr(v))
    End If
End Function


