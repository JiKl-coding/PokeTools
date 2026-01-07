VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Movelist 
   Caption         =   "Movelist"
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
' UserForm: Movelist2
' Custom Grid (Label + Frame)
'===============================
Option Explicit

' Learnsets: column letter where MOVE NAME is stored (NOT method!)
' Method is in column E and Level is in column F.
Private Const LEARNSETS_MOVE_COL As String = "D"   ' move name col in Learnsets

Private Const FILTER_ALL As String = "All"

' Visual layout (points)
Private Const PAD As Single = 6
Private Const HEADER_H As Single = 18
Private Const ROW_MIN_H As Single = 18

' Column indices
Private Enum GridCol
    gcMove = 0
    gcType = 1
    gcCategory = 2
    gcPower = 3
    gcAccuracy = 4
    gcPP = 5
    gcPriority = 6
    gcDescription = 7
    gcMethod = 8
End Enum

' Column widths (match old listbox proportions)
Private mColW(0 To 8) As Single

' Runtime UI
Private mLblInfo As MSForms.Label
Private mCboPokemon As MSForms.ComboBox
Private mCboType As MSForms.ComboBox
Private mCboMethod As MSForms.ComboBox
Private mCboGame As MSForms.ComboBox
Private mFraHeader As MSForms.Frame
Private mFraGrid As MSForms.Frame

' Event handlers for dynamic controls
Private mHeaderEvents As Collection
Private mFilterEvents As Collection
Private mRowEvents As Collection

' Data
Private pdWB As Workbook

Private Type MoveRow
    moveName As String
    MoveType As String
    Category As String
    Power As String
    Accuracy As String
    PP As String
    Priority As String
    Description As String
    method As String
End Type

Private mRows() As MoveRow
Private mRowCount As Long

' Sort state
Private mSortCol As Long
Private mSortAsc As Boolean

' Last selections to detect changes
Private mLastPokemon As String
Private mLastGameSel As String
Private mInFilterUpdate As Boolean

' Typed combo caches
Private mAllPokemon() As String
Private mAllTypes() As String
Private mAllMethods() As String
Private mSuppressTyped As Boolean

' =============================
' Form init
' =============================
Private Sub UserForm_Initialize()
    ' styling (keep consistent with Movelist)
    Me.BackColor = RGB(204, 0, 0)

    ' If opened outside Pokedex, reset context to All-moves view
    Dim onPokedex As Boolean
    onPokedex = False
    On Error Resume Next
    onPokedex = (StrComp(ActiveSheet.CodeName, "Pokedex", vbTextCompare) = 0)
    On Error GoTo 0
    If Not onPokedex Then
        On Error Resume Next
        Application.EnableEvents = False
        Pokedex.Range("PKMN_DEX").value = ""
        Pokedex.Range("GAME").value = FILTER_ALL
        Application.EnableEvents = True
        On Error GoTo 0
    End If

    InitColumnWidths
    BuildRuntimeUI

    ' Ensure GAME context is initialized and dependent lists are populated
    Dim g0 As String
    g0 = Trim$(CStr(Pokedex.Range("GAME").value))
    If Len(g0) = 0 Then
        Pokedex.Range("GAME").value = FILTER_ALL
    End If
    On Error Resume Next
    Application.Wait Now + TimeValue("00:00:01")
    Application.Calculate
    On Error GoTo 0

    Set pdWB = Functions.GetPokedataWb()

    ' Populate filters first (guard events), then load data
    mInFilterUpdate = True
    PopulateFilters
    ' Explicitly set Pokemon dropdown to current context cell
    SetPokemonSelectionFromContext
    mInFilterUpdate = False

    LoadData
    SetInfoLabel

    ' default sort: by Move asc
    mSortCol = gcMove
    mSortAsc = True

    RenderGrid
End Sub

Private Sub InitColumnWidths()
    ' Move | Type | Category | Power | Accuracy | PP | Priority | Description | Method
    mColW(gcMove) = 120
    mColW(gcType) = 70
    mColW(gcCategory) = 70
    mColW(gcPower) = 50
    mColW(gcAccuracy) = 70
    mColW(gcPP) = 40
    mColW(gcPriority) = 60
    mColW(gcDescription) = 300
    mColW(gcMethod) = 120
End Sub

Private Sub BuildRuntimeUI()
    Dim x As Single, y As Single

    ' Clean any pre-existing designer controls (if blob contains old ones)
    HideIfExists "lbMoves"
    HideIfExists "txtDescription"
    HideIfExists "lblInfo"

    ' Info label
    Set mLblInfo = Me.Controls.Add("Forms.Label.1", "lblInfo2", True)
    With mLblInfo
        .Left = PAD
        .Top = PAD
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = 18
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .caption = "Movelist"
    End With

    ' Filters
    y = mLblInfo.Top + mLblInfo.Height + PAD
    x = PAD

    ' Pokemon filter (first)
    Dim lblP As MSForms.Label
    Set lblP = Me.Controls.Add("Forms.Label.1", "lblPokemon", True)
    With lblP
        .Left = x
        .Top = y + 2
        .Width = 60
        .Height = 16
        .caption = "Pokemon"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
    End With
    x = x + lblP.Width + 4
    Set mCboPokemon = Me.Controls.Add("Forms.ComboBox.1", "cboPokemon", True)
    With mCboPokemon
        .Left = x
        .Top = y
        .Width = 150
        .Height = 18
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryNone
    End With

    x = mCboPokemon.Left + mCboPokemon.Width + 12

    ' Type filter
    Dim lblT As MSForms.Label
    Set lblT = Me.Controls.Add("Forms.Label.1", "lblType", True)
    With lblT
        .Left = x
        .Top = y + 2
        .Width = 35
        .Height = 16
        .caption = "Type"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
    End With
    x = x + lblT.Width + 4
    Set mCboType = Me.Controls.Add("Forms.ComboBox.1", "cboType", True)
    With mCboType
        .Left = x
        .Top = y
        .Width = 120
        .Height = 18
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryNone
    End With

    x = mCboType.Left + mCboType.Width + 12

    ' Method filter (physical/special/status)
    Dim lblM As MSForms.Label
    Set lblM = Me.Controls.Add("Forms.Label.1", "lblMethod", True)
    With lblM
        .Left = x
        .Top = y + 2
        .Width = 55
        .Height = 16
        .caption = "Method"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
    End With
    x = x + lblM.Width + 4
    Set mCboMethod = Me.Controls.Add("Forms.ComboBox.1", "cboMethod", True)
    With mCboMethod
        .Left = x
        .Top = y
        .Width = 120
        .Height = 18
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryNone
    End With

    x = mCboMethod.Left + mCboMethod.Width + 12

    ' Game filter (visible only in All/All)
    Dim lblG As MSForms.Label
    Set lblG = Me.Controls.Add("Forms.Label.1", "lblGame", True)
    With lblG
        .Left = x
        .Top = y + 2
        .Width = 45
        .Height = 16
        .caption = "Game"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Visible = True
    End With
    x = x + lblG.Width + 4
    Set mCboGame = Me.Controls.Add("Forms.ComboBox.1", "cboGame", True)
    With mCboGame
        .Left = x
        .Top = y
        .Width = 140
        .Height = 18
        .Style = fmStyleDropDownList
        .Visible = True
    End With

    ' Events (dynamic)
    Set mFilterEvents = New Collection
    Dim e1 As CGridComboEvents, e2 As CGridComboEvents, e3 As CGridComboEvents, e4 As CGridComboEvents
    Set e1 = New CGridComboEvents
    e1.Init Me, mCboPokemon
    mFilterEvents.Add e1
    Set e2 = New CGridComboEvents
    e2.Init Me, mCboType
    mFilterEvents.Add e2
    Set e3 = New CGridComboEvents
    e3.Init Me, mCboMethod
    mFilterEvents.Add e3
    Set e4 = New CGridComboEvents
    e4.Init Me, mCboGame
    mFilterEvents.Add e4

    ' Header frame
    y = y + 22 + PAD
    Set mFraHeader = Me.Controls.Add("Forms.Frame.1", "fraHeader", True)
    With mFraHeader
        .Left = PAD
        .Top = y
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = HEADER_H + 4
        .caption = vbNullString
        .BackColor = RGB(173, 216, 230)
        .SpecialEffect = fmSpecialEffectFlat
    End With

    BuildHeaderLabels

    ' Grid frame (scroll)
    y = mFraHeader.Top + mFraHeader.Height
    Set mFraGrid = Me.Controls.Add("Forms.Frame.1", "fraGrid", True)
    With mFraGrid
        .Left = PAD
        .Top = y
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = Me.InsideHeight - y - PAD
        .caption = vbNullString
        .BackColor = RGB(173, 216, 230)
        .SpecialEffect = fmSpecialEffectFlat
        .ScrollBars = fmScrollBarsVertical
        .ScrollTop = 0
    End With
End Sub

Private Sub BuildHeaderLabels()
    Set mHeaderEvents = New Collection

    Dim captions(0 To 8) As String
    captions(gcMove) = "Move"
    captions(gcType) = "Type"
    captions(gcCategory) = "Category"
    captions(gcPower) = "Power"
    captions(gcAccuracy) = "Accuracy"
    captions(gcPP) = "PP"
    captions(gcPriority) = "Priority"
    captions(gcDescription) = "Description"
    captions(gcMethod) = "Method"

    Dim i As Long
    Dim x As Single
    x = 2

    For i = 0 To 8
        Dim h As MSForms.Label
        Set h = mFraHeader.Controls.Add("Forms.Label.1", "h" & CStr(i), True)
        With h
            .Left = x
            .Top = 2
            .Width = mColW(i)
            .Height = HEADER_H
            .caption = captions(i)
            .Font.Bold = True
            .BackStyle = fmBackStyleTransparent
            .ForeColor = vbBlack
            .Tag = CStr(i)
        End With

        Dim ev As CGridHeaderLabel
        Set ev = New CGridHeaderLabel
        ev.Init Me, h, i
        mHeaderEvents.Add ev

        x = x + mColW(i)
    Next i
End Sub

Private Sub HideIfExists(ByVal ctrlName As String)
    On Error Resume Next
    Me.Controls(ctrlName).Visible = False
    On Error GoTo 0
End Sub

' =============================
' Public callbacks for event classes
' =============================
Public Sub HeaderClicked(ByVal colIndex As Long)
    If mSortCol = colIndex Then
        mSortAsc = Not mSortAsc
    Else
        mSortCol = colIndex
        mSortAsc = True
    End If

    RenderGrid
End Sub

Public Sub FiltersChanged()
    If mInFilterUpdate Then Exit Sub
    mInFilterUpdate = True
    Dim evWasEnabled As Variant
    On Error Resume Next
    evWasEnabled = Application.EnableEvents
    On Error GoTo 0
    On Error GoTo CleanUp

    Dim curPokemon As String
    Dim curGame As String
    curPokemon = Trim$(CStr(mCboPokemon.value))
    curGame = Trim$(CStr(mCboGame.value))

    Dim pokemonChanged As Boolean
    Dim gameChanged As Boolean
    pokemonChanged = (StrComp(curPokemon, mLastPokemon, vbTextCompare) <> 0)
    gameChanged = (StrComp(curGame, mLastGameSel, vbTextCompare) <> 0)

    ' Update context cells first
    If pokemonChanged Then
        If Len(curPokemon) = 0 Then
            MsgBox "Choose pokemon first", vbExclamation, "Movelist"
            GoTo CleanUp
        End If
        On Error Resume Next
        Pokedex.Range("PKMN_DEX").value = curPokemon
        On Error GoTo 0
    End If

    If gameChanged Then
        If Len(curGame) > 0 And curGame <> "0" Then
            On Error Resume Next
            Pokedex.Range("GAME").value = curGame
            On Error GoTo 0
        End If
    End If

    UpdateGameVisibility

    If pokemonChanged Or gameChanged Then
        On Error Resume Next
        Application.Wait Now + TimeValue("00:00:01")
        Application.Calculate
        On Error GoTo 0

        ' After recalc, read normalized values back from UI cells
        Dim normPkmn As String, normGame As String
        normPkmn = Trim$(CStr(Pokedex.Range("PKMN_DEX").value))
        normGame = Trim$(CStr(Pokedex.Range("GAME").value))

        ' Refresh Pokemon dropdown from Lists!O (depends on GAME)
        mInFilterUpdate = True
        RefreshPokemonListFromLists normPkmn
        ' Ensure Game reflects normalized value
        EnsureComboHasValue mCboGame, normGame, False
        mCboGame.value = normGame
        mInFilterUpdate = False

        LoadData
        SetInfoLabel
    End If

    RenderGrid
    
CleanUp:
    On Error Resume Next
    Application.EnableEvents = evWasEnabled
    On Error GoTo 0
    mLastPokemon = Trim$(CStr(mCboPokemon.value))
    mLastGameSel = Trim$(CStr(mCboGame.value))
    mInFilterUpdate = False
End Sub

' =============================
' Context + info
' =============================
Private Sub SetInfoLabel()
    Dim pkmnDex As String, game As String
    pkmnDex = GetContextPokemon()
    game = Trim$(CStr(Pokedex.Range("GAME").value))

    If Len(pkmnDex) = 0 Or StrComp(pkmnDex, FILTER_ALL, vbTextCompare) = 0 Then
        mLblInfo.caption = "Movelist (All Moves) (" & game & ")"
        Me.caption = mLblInfo.caption
    Else
        mLblInfo.caption = "Movelist of " & pkmnDex & " (" & game & ")"
        Me.caption = mLblInfo.caption
    End If
End Sub

Private Function GetContextPokemon() As String
    On Error GoTo Fallback

    If Not ActiveSheet Is Nothing Then
        If StrComp(ActiveSheet.CodeName, "Pokedex", vbTextCompare) = 0 Then
            GetContextPokemon = Trim$(CStr(Pokedex.Range("PKMN_DEX").value))
            Exit Function
        End If
    End If

Fallback:
    ' For now: default to Pokedex sheet. (Future: per-sheet mapping)
    On Error Resume Next
    GetContextPokemon = Trim$(CStr(Pokedex.Range("PKMN_DEX").value))
    On Error GoTo 0
End Function

' =============================
' Load + filters
' =============================
Private Sub LoadData()
    Dim pkmnDex As String
    Dim game As String
    Dim gameNorm As String

    ' Prefer Pokemon from dropdown when provided
    Dim selP As String
    selP = ""
    On Error Resume Next
    selP = Trim$(CStr(mCboPokemon.value))
    On Error GoTo 0

    If Len(selP) > 0 And StrComp(selP, FILTER_ALL, vbTextCompare) <> 0 Then
        pkmnDex = selP
    Else
        pkmnDex = GetContextPokemon()
    End If
    game = Trim$(CStr(Pokedex.Range("GAME").value))
    gameNorm = DexLogic.NormalizeGameVersion(game)

    Dim allMovesMode As Boolean
    allMovesMode = (Len(pkmnDex) = 0 Or StrComp(pkmnDex, FILTER_ALL, vbTextCompare) = 0)

    Dim dictMoves As Object
    Set dictMoves = BuildMovesDict(pdWB.Worksheets("Moves"))

    Dim dictMethod As Object
    If allMovesMode Then
        Set dictMethod = Nothing
    Else
        Set dictMethod = BuildLearnsetsMethodDict(pdWB.Worksheets("Learnsets"), pkmnDex, gameNorm)
    End If

    Dim names As Variant
    If allMovesMode Then
        names = GetAllMoveNames(pdWB.Worksheets("Moves"))
    Else
        names = GetMoveNamesFromLists()
    End If

    BuildRowsFromNames names, dictMoves, dictMethod, pkmnDex, gameNorm, allMovesMode
End Sub

Private Sub UpdateGameVisibility()
    On Error Resume Next
    Dim lblG As MSForms.Label
    Set lblG = Me.Controls("lblGame")
    If Not lblG Is Nothing Then lblG.Visible = True
    mCboGame.Visible = True
    On Error GoTo 0
End Sub

Private Sub PopulateFilters()
    ' Pokemon: source depends on GAME
    RefreshPokemonListFromLists ""
    ' Type: Lists A (skip 0/empty)
    PopulateComboUniqueFromColumn mCboType, Lists, "A", FILTER_ALL, True
    ' Method: physical/special/status from Lists G (skip 0/empty)
    PopulateComboUniqueFromColumn mCboMethod, Lists, "G", FILTER_ALL, True
    ' Game: Lists F (no All, skip 0/empty)
    PopulateComboUniqueFromColumn mCboGame, Lists, "F", "", True

    ' All first for relevant dropdowns
    SetPokemonSelectionFromContext
    mCboType.ListIndex = 0
    mCboMethod.ListIndex = 0
    SetGameSelectionFromContext

    ' Remember last selections
    mLastPokemon = Trim$(CStr(mCboPokemon.value))
    mLastGameSel = Trim$(CStr(mCboGame.value))

    UpdateGameVisibility

    ' Capture typed caches
    CaptureComboItemsToArray mCboPokemon, mAllPokemon
    CaptureComboItemsToArray mCboType, mAllTypes
    CaptureComboItemsToArray mCboMethod, mAllMethods
End Sub

Private Sub PopulateComboUniqueFromColumn(ByVal cbo As MSForms.ComboBox, ByVal ws As Worksheet, ByVal colLetter As String, ByVal allValue As String, Optional skipZero As Boolean = False)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).row

    For r = 2 To lastRow
        Dim v As String
        v = Trim$(CStr(ws.Cells(r, colLetter).value))
        If skipZero And (v = "0" Or v = "") Then GoTo NextR
        If Len(v) > 0 Then
            If Not dict.Exists(v) Then dict.Add v, True
        End If
NextR:
    Next r

    cbo.Clear
    If allValue <> "" Then cbo.AddItem allValue

    Dim k As Variant
    For Each k In dict.Keys
        cbo.AddItem CStr(k)
    Next k
End Sub

Private Sub SetPokemonSelectionFromContext()
    Dim ctx As String
    ctx = Trim$(CStr(Pokedex.Range("PKMN_DEX").value))
    If Len(ctx) = 0 Then Exit Sub

    RefreshPokemonListFromLists ctx
End Sub

Private Sub EnsureComboHasValue(ByRef cbo As MSForms.ComboBox, ByVal val As String, Optional ByVal insertAtTop As Boolean = False)
    If Len(val) = 0 Then Exit Sub
    If Not ComboContains(cbo, val) Then
        If insertAtTop Then
            cbo.AddItem val, 0
        Else
            cbo.AddItem val
        End If
    End If
End Sub

Private Function ComboContains(ByVal cbo As MSForms.ComboBox, ByVal val As String) As Boolean
    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If StrComp(CStr(cbo.List(i)), val, vbTextCompare) = 0 Then
            ComboContains = True
            Exit Function
        End If
    Next i
End Function

Private Sub SetGameSelectionFromContext()
    Dim g As String
    g = Trim$(CStr(Pokedex.Range("GAME").value))
    If Len(g) = 0 Or g = "0" Then
        If mCboGame.ListCount > 0 Then mCboGame.ListIndex = 0
        Exit Sub
    End If

    EnsureComboHasValue mCboGame, g, False
    mCboGame.value = g
End Sub

Private Sub RefreshPokemonListFromLists(ByVal desiredSelection As String)
    Dim keepSel As String
    keepSel = desiredSelection

    ' Rebuild list: for GAME=All use Lists!B, otherwise Lists!O
    Dim g As String
    g = Trim$(CStr(Pokedex.Range("GAME").value))
    Dim colLetter As String
    If StrComp(g, FILTER_ALL, vbTextCompare) = 0 Then
        colLetter = "B"
    Else
        colLetter = "O"
    End If
    PopulateComboUniqueFromColumn mCboPokemon, Lists, colLetter, "", False

    If Len(keepSel) > 0 Then
        EnsureComboHasValue mCboPokemon, keepSel, True
        mCboPokemon.value = keepSel
    ElseIf mCboPokemon.ListCount > 0 Then
        mCboPokemon.ListIndex = 0
    End If

    ' Update cache after rebuild
    CaptureComboItemsToArray mCboPokemon, mAllPokemon
End Sub

Private Function GetMoveNamesFromLists() As Variant
    ' Lists is the CodeName of sheet containing move list in column P
    Dim lastRow As Long, r As Long, n As Long
    lastRow = Lists.Cells(Lists.Rows.Count, "P").End(xlUp).row

    Dim out() As String
    n = 0

    For r = 2 To lastRow
        Dim mv As String
        mv = Trim$(CStr(Lists.Cells(r, "P").value))
        If Len(mv) > 0 Then
            n = n + 1
            ReDim Preserve out(1 To n)
            out(n) = mv
        End If
    Next r

    If n = 0 Then
        Dim defArr(1 To 1) As String
        defArr(1) = "-"
        GetMoveNamesFromLists = defArr
    Else
        GetMoveNamesFromLists = out
    End If
End Function

Private Function GetAllMoveNames(ByVal wsMoves As Worksheet) As Variant
    Dim lastRow As Long, r As Long, n As Long
    lastRow = SafeLastDataRow(wsMoves, "B")

    Dim out() As String
    n = 0

    For r = 2 To lastRow
        Dim mv As String
        mv = Trim$(CStr(wsMoves.Cells(r, "B").value))
        If Len(mv) > 0 Then
            n = n + 1
            ReDim Preserve out(1 To n)
            out(n) = mv
        End If
    Next r

    If n = 0 Then
        Dim defArr(1 To 1) As String
        defArr(1) = "-"
        GetAllMoveNames = defArr
    Else
        GetAllMoveNames = out
    End If
End Function

Private Sub BuildRowsFromNames(ByVal names As Variant, ByVal dictMoves As Object, ByVal dictMethod As Object, _
                              ByVal pkmnDex As String, ByVal gameNorm As String, ByVal allMovesMode As Boolean)
    Dim i As Long
    Dim n As Long

    On Error GoTo SafeDefault
    n = UBound(names) - LBound(names) + 1
    If n <= 0 Then GoTo SafeDefault

    ReDim mRows(1 To n)
    mRowCount = 0

    For i = LBound(names) To UBound(names)
        Dim moveName As String
        moveName = Trim$(CStr(names(i)))
        If Len(moveName) = 0 Or moveName = "-" Then GoTo ContinueI

        mRowCount = mRowCount + 1

        Dim row As MoveRow
        FillMoveRow row, moveName, dictMoves, dictMethod, pkmnDex, gameNorm, allMovesMode
        mRows(mRowCount) = row

ContinueI:
    Next i

    If mRowCount = 0 Then GoTo SafeDefault

    If mRowCount < n Then
        ReDim Preserve mRows(1 To mRowCount)
    End If

    Exit Sub

SafeDefault:
    ReDim mRows(1 To 1)
    mRowCount = 0
End Sub

Private Sub FillMoveRow(ByRef row As MoveRow, ByVal moveName As String, _
                        ByVal dictMoves As Object, ByVal dictMethod As Object, _
                        ByVal pkmnDex As String, ByVal gameNorm As String, ByVal allMovesMode As Boolean)
    row.moveName = moveName

    ' Moves sheet details
    If dictMoves.Exists(LCase$(moveName)) Then
        Dim arr As Variant
        arr = dictMoves(LCase$(moveName))
        row.MoveType = Nz(arr(0))
        row.Category = Nz(arr(1))
        row.Power = Nz(arr(2))
        row.Accuracy = Nz(arr(3))
        row.PP = Nz(arr(4))
        row.Priority = Nz(arr(5))
        row.Description = Nz(arr(6))
    Else
        row.MoveType = ""
        row.Category = "?"
        row.Power = ""
        row.Accuracy = ""
        row.PP = ""
        row.Priority = ""
        row.Description = ""
    End If

    ' Method (only when bound to specific pokemon)
    If allMovesMode Then
        row.method = "-"
    Else
        Dim key As String
        key = LearnKey(pkmnDex, gameNorm, moveName)
        If Not dictMethod Is Nothing Then
            If dictMethod.Exists(key) Then
                row.method = CStr(dictMethod(key))
            Else
                row.method = "-"
            End If
        Else
            row.method = "-"
        End If
    End If
End Sub

' =============================
' Rendering
' =============================
Private Sub RenderGrid()
    Dim prevSU As Boolean
    prevSU = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False
    On Error GoTo 0

    On Error Resume Next
    mFraGrid.Visible = False
    On Error GoTo 0

    ClearGridRows

    ' reset row event handlers for fresh render
    Set mRowEvents = New Collection

    If mRowCount <= 0 Then GoTo FinishRender

    Dim filteredIdx() As Long
    filteredIdx = GetFilteredIndices()
    If UBound(filteredIdx) = 0 Then Exit Sub

    SortIndices filteredIdx, 1, UBound(filteredIdx)

    Dim y As Single
    y = 2

    Dim i As Long
    For i = 1 To UBound(filteredIdx)
        Dim idx As Long
        idx = filteredIdx(i)

        Dim rh As Single
        rh = CalcRowHeight(mRows(idx).Description)

        AddGridRow idx, y, rh
        y = y + rh
    Next i

    mFraGrid.ScrollHeight = y + 4

    On Error Resume Next
    mFraGrid.Visible = True
    On Error GoTo 0

FinishRender:
    On Error Resume Next
    Application.ScreenUpdating = prevSU
    On Error GoTo 0
End Sub

Private Sub ClearGridRows()
    On Error Resume Next
    Dim c As MSForms.Control
    Dim namesToRemove As Collection
    Set namesToRemove = New Collection

    For Each c In mFraGrid.Controls
        If Left$(c.name, 3) = "r__" Then
            namesToRemove.Add c.name
        End If
    Next c

    Dim n As Variant
    For Each n In namesToRemove
        mFraGrid.Controls.Remove CStr(n)
    Next n
    On Error GoTo 0
End Sub

Private Sub AddGridRow(ByVal rowIndex As Long, ByVal topY As Single, ByVal rowH As Single)
    Dim x As Single
    x = 2

    Dim row As MoveRow
    row = mRows(rowIndex)

    AddCellLabel "r__m" & rowIndex, row.moveName, x, topY, mColW(gcMove), rowH, False
    AttachRowEvent "r__m" & rowIndex, rowIndex
    x = x + mColW(gcMove)

    AddCellLabel "r__t" & rowIndex, row.MoveType, x, topY, mColW(gcType), rowH, False
    AttachRowEvent "r__t" & rowIndex, rowIndex
    x = x + mColW(gcType)

    AddCellLabel "r__c" & rowIndex, row.Category, x, topY, mColW(gcCategory), rowH, False
    AttachRowEvent "r__c" & rowIndex, rowIndex
    x = x + mColW(gcCategory)

    AddCellLabel "r__p" & rowIndex, row.Power, x, topY, mColW(gcPower), rowH, True
    AttachRowEvent "r__p" & rowIndex, rowIndex
    x = x + mColW(gcPower)

    AddCellLabel "r__a" & rowIndex, row.Accuracy, x, topY, mColW(gcAccuracy), rowH, True
    AttachRowEvent "r__a" & rowIndex, rowIndex
    x = x + mColW(gcAccuracy)

    AddCellLabel "r__pp" & rowIndex, row.PP, x, topY, mColW(gcPP), rowH, True
    AttachRowEvent "r__pp" & rowIndex, rowIndex
    x = x + mColW(gcPP)

    AddCellLabel "r__pr" & rowIndex, row.Priority, x, topY, mColW(gcPriority), rowH, True
    AttachRowEvent "r__pr" & rowIndex, rowIndex
    x = x + mColW(gcPriority)

    AddCellLabel "r__d" & rowIndex, row.Description, x, topY, mColW(gcDescription), rowH, False, True
    AttachRowEvent "r__d" & rowIndex, rowIndex
    x = x + mColW(gcDescription)

    AddCellLabel "r__me" & rowIndex, row.method, x, topY, mColW(gcMethod), rowH, False
    AttachRowEvent "r__me" & rowIndex, rowIndex
End Sub

Private Sub AddCellLabel(ByVal name As String, ByVal caption As String, _
                         ByVal leftX As Single, ByVal topY As Single, _
                         ByVal w As Single, ByVal h As Single, _
                         ByVal center As Boolean, Optional ByVal wrap As Boolean = False)
    Dim lbl As MSForms.Label
    Set lbl = mFraGrid.Controls.Add("Forms.Label.1", name, True)

    With lbl
        .Left = leftX
        .Top = topY
        .Width = w
        .Height = h
        .caption = caption
        .BackStyle = fmBackStyleTransparent
        .ForeColor = vbBlack
        .WordWrap = wrap
        .AutoSize = False
        If center Then
            .TextAlign = fmTextAlignCenter
        Else
            .TextAlign = fmTextAlignLeft
        End If
    End With
End Sub

Private Sub AttachRowEvent(ByVal ctrlName As String, ByVal rowIndex As Long)
    On Error Resume Next
    Dim lbl As MSForms.Label
    Set lbl = mFraGrid.Controls(ctrlName)
    On Error GoTo 0
    If lbl Is Nothing Then Exit Sub

    Dim ev As CGridRowLabel
    Set ev = New CGridRowLabel
    ev.Init Me, lbl, rowIndex
    mRowEvents.Add ev
End Sub

Public Sub OnRowDoubleClick(ByVal rowIndex As Long)
    If rowIndex <= 0 Or rowIndex > mRowCount Then Exit Sub
    Dim mv As String
    mv = mRows(rowIndex).moveName
    On Error Resume Next
    Pokedex.Range("PKMN_MOVELIST").value = mv
    On Error GoTo 0
    ' Close the form after selection
    Unload Me
End Sub

Private Function CalcRowHeight(ByVal desc As String) As Single
    ' Lightweight wrap estimation to avoid txtDescription.
    Const CHARS_PER_LINE As Long = 65

    Dim lines As Long
    If Len(desc) = 0 Then
        lines = 1
    Else
        lines = (Len(desc) + CHARS_PER_LINE - 1) \ CHARS_PER_LINE
        If lines < 1 Then lines = 1
        If lines > 6 Then lines = 6 ' avoid runaway tall rows
    End If

    CalcRowHeight = Application.WorksheetFunction.Max(ROW_MIN_H, (ROW_MIN_H - 2) * lines)
End Function

' =============================
' Filtering + sorting
' =============================
Private Function GetFilteredIndices() As Long()
    Dim pokeSel As String, typeSel As String, methodSel As String, gameSel As String
    pokeSel = Trim$(CStr(mCboPokemon.value))
    typeSel = Trim$(CStr(mCboType.value))
    methodSel = Trim$(CStr(mCboMethod.value))
    gameSel = Trim$(CStr(mCboGame.value))

    Dim tmp() As Long
    ReDim tmp(1 To mRowCount)

    Dim n As Long
    n = 0

    Dim i As Long
    For i = 1 To mRowCount
        Dim ok As Boolean
        ok = True

        ' Pokemon filter: handled by data reload; no per-row check here.

        ' Type filter (skip empty/0)
        If ok Then
            If Len(typeSel) > 0 And StrComp(typeSel, FILTER_ALL, vbTextCompare) <> 0 Then
                If mRows(i).MoveType = "" Or mRows(i).MoveType = "0" Or StrComp(mRows(i).MoveType, typeSel, vbTextCompare) <> 0 Then ok = False
            End If
        End If

        ' Method filter (physical/special/status => Category)
        If ok Then
            If Len(methodSel) > 0 And StrComp(methodSel, FILTER_ALL, vbTextCompare) <> 0 Then
                If mRows(i).Category = "" Or StrComp(mRows(i).Category, methodSel, vbTextCompare) <> 0 Then ok = False
            End If
        End If

        ' Game filter only when visible (All/All mode). Without per-move availability mapping, skip or implement later.
        ' Keep placeholder: no-op unless future data ties moves to game.

        If ok Then
            n = n + 1
            tmp(n) = i
        End If
    Next i

    If n = 0 Then
        Dim emptyArr(0 To 0) As Long
        GetFilteredIndices = emptyArr
    Else
        ReDim Preserve tmp(1 To n)
        GetFilteredIndices = tmp
    End If
End Function

Private Sub SortIndices(ByRef idx() As Long, ByVal lo As Long, ByVal hi As Long)
    ' QuickSort on indices based on current sort column
    Dim i As Long, j As Long
    i = lo
    j = hi

    Dim pivot As Variant
    pivot = SortKeyForIndex(idx((lo + hi) \ 2))

    Do While i <= j
        Do While CompareKeys(SortKeyForIndex(idx(i)), pivot) < 0
            i = i + 1
        Loop
        Do While CompareKeys(SortKeyForIndex(idx(j)), pivot) > 0
            j = j - 1
        Loop

        If i <= j Then
            Dim t As Long
            t = idx(i)
            idx(i) = idx(j)
            idx(j) = t
            i = i + 1
            j = j - 1
        End If
    Loop

    If lo < j Then SortIndices idx, lo, j
    If i < hi Then SortIndices idx, i, hi
End Sub

Private Function SortKeyForIndex(ByVal i As Long) As Variant
    Select Case mSortCol
        Case gcMove: SortKeyForIndex = LCase$(mRows(i).moveName)
        Case gcType: SortKeyForIndex = LCase$(mRows(i).MoveType)
        Case gcCategory: SortKeyForIndex = LCase$(mRows(i).Category)
        Case gcPower: SortKeyForIndex = CLng(val(mRows(i).Power))
        Case gcAccuracy: SortKeyForIndex = CLng(val(mRows(i).Accuracy))
        Case gcPP: SortKeyForIndex = CLng(val(mRows(i).PP))
        Case gcPriority: SortKeyForIndex = CLng(val(mRows(i).Priority))
        Case gcDescription: SortKeyForIndex = LCase$(mRows(i).Description)
        Case gcMethod: SortKeyForIndex = LCase$(mRows(i).method)
        Case Else: SortKeyForIndex = LCase$(mRows(i).moveName)
    End Select
End Function

Private Function CompareKeys(ByVal a As Variant, ByVal b As Variant) As Long
    Dim r As Long

    If IsNumeric(a) And IsNumeric(b) Then
        If CLng(a) < CLng(b) Then
            r = -1
        ElseIf CLng(a) > CLng(b) Then
            r = 1
        Else
            r = 0
        End If
    Else
        r = StrComp(CStr(a), CStr(b), vbTextCompare)
    End If

    If mSortAsc Then
        CompareKeys = r
    Else
        CompareKeys = -r
    End If
End Function

' =============================
' Dictionary builders
' =============================
Private Function BuildMovesDict(ByVal wsMoves As Worksheet) As Object
    ' Moves sheet columns:
    ' B = Move name
    ' C = Type
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
    Dim arr(0 To 6) As Variant

    lastRow = SafeLastDataRow(wsMoves, "B")

    For r = 2 To lastRow
        nameKey = LCase$(Trim$(CStr(wsMoves.Cells(r, "B").value)))
        If Len(nameKey) > 0 Then
            arr(0) = wsMoves.Cells(r, "C").value ' Type
            arr(1) = wsMoves.Cells(r, "D").value ' Category
            arr(2) = wsMoves.Cells(r, "E").value ' Power
            arr(3) = wsMoves.Cells(r, "F").value ' Accuracy
            arr(4) = wsMoves.Cells(r, "G").value ' PP
            arr(5) = wsMoves.Cells(r, "H").value ' Priority
            arr(6) = wsMoves.Cells(r, "I").value ' Description
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

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long
    Dim poke As String, ver As String, verNorm As String
    Dim move As String, method As String, lvl As String
    Dim key As String, outMethod As String

    lastRow = SafeLastDataRow(wsLearn, "B")

    For r = 2 To lastRow
        poke = Trim$(CStr(wsLearn.Cells(r, "B").value))
        If Len(poke) = 0 Then GoTo ContinueRow

        ver = Trim$(CStr(wsLearn.Cells(r, "C").value))
        verNorm = DexLogic.NormalizeGameVersion(ver)

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
        dict(key) = outMethod

ContinueRow:
    Next r

    Set BuildLearnsetsMethodDict = dict
End Function

Private Function LearnKey(ByVal pkmnDex As String, ByVal gameNorm As String, ByVal moveName As String) As String
    LearnKey = LCase$(Trim$(pkmnDex)) & "|" & LCase$(Trim$(gameNorm)) & "|" & LCase$(Trim$(moveName))
End Function
' Safe last-row helper (avoids huge ranges if column has stray content)
Private Function SafeLastDataRow(ByVal ws As Worksheet, ByVal colLetter As String) As Long
    On Error GoTo CleanFail
    Dim rng As Range, f As Range
    Set rng = ws.Columns(colLetter)
    Set f = rng.Find(What:="*", After:=rng.Cells(1, 1), LookIn:=xlValues, _
                     LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then
        SafeLastDataRow = 1
    Else
        SafeLastDataRow = f.Row
    End If
    Exit Function
CleanFail:
    SafeLastDataRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
End Function

' =============================
' Helpers
' =============================
Private Function Nz(ByVal v As Variant) As String
    If IsError(v) Then
        Nz = ""
    Else
        Nz = Trim$(CStr(v))
    End If
End Function

' =============================
' Typed ComboBox filtering (UI-only)
' =============================
Private Sub ComboTyped(ByVal ctrlName As String, ByVal typed As String)
    If mSuppressTyped Then Exit Sub
    Select Case LCase$(ctrlName)
        Case "cbopokemon"
            FilterComboDropdown mCboPokemon, mAllPokemon, typed
        Case "cbotype"
            FilterComboDropdown mCboType, mAllTypes, typed
        Case "cbomethod"
            FilterComboDropdown mCboMethod, mAllMethods, typed
        Case Else
            ' ignore others
    End Select
End Sub

Private Sub FilterComboDropdown(ByRef cbo As MSForms.ComboBox, ByRef cache() As String, ByVal prefix As String)
    On Error GoTo CleanExit
    mSuppressTyped = True
    Dim i As Long
    Dim pfx As String: pfx = LCase$(prefix)
    cbo.Clear
    If Not (Not cache) Then ' array is initialized
        For i = LBound(cache) To UBound(cache)
            If pfx = "" Or LCase$(cache(i)) Like pfx & "*" Then
                cbo.AddItem cache(i)
            End If
        Next i
    End If
    If cbo.ListCount = 0 Then cbo.AddItem "(no match)"
    cbo.DropDown
    cbo.Text = prefix
    cbo.SelStart = Len(prefix)
CleanExit:
    mSuppressTyped = False
End Sub

Private Sub CaptureComboItemsToArray(ByRef cbo As MSForms.ComboBox, ByRef arr() As String)
    Dim n As Long: n = cbo.ListCount
    If n <= 0 Then
        Erase arr
        Exit Sub
    End If
    ReDim arr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        arr(i) = CStr(cbo.List(i))
    Next i
End Sub
