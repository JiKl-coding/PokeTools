VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Movelist 
   Caption         =   "Movelist"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23760
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
' Custom Grid (Label + Frame)
'===============================
Option Explicit

Private Const FILTER_ALL As String = "All"
Private Const GAME_KEY_ALL As String = "__all__"
Private Const TMP_MOVE_HEADER As String = "TmpMovelist"
Private Const UI_FONT_NAME As String = "Aptos Narrow"
Private Const UI_FONT_SIZE As Integer = 12

' Visual layout (points)
Private Const PAD As Single = 6
Private Const HEADER_H As Single = 18
Private Const ROW_MIN_H As Single = 22

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
End Enum

' Column widths (match old listbox proportions)
Private mColW(0 To 7) As Single

' Runtime UI
Private mLblInfo As MSForms.label
Private mCboPokemon As MSForms.ComboBox
Private mCboType As MSForms.ComboBox
Private mCboGame As MSForms.ComboBox
Private WithEvents mBtnApply As MSForms.CommandButton
Attribute mBtnApply.VB_VarHelpID = -1
Private WithEvents mBtnClear As MSForms.CommandButton
Attribute mBtnClear.VB_VarHelpID = -1
Private WithEvents mBtnClose As MSForms.CommandButton
Attribute mBtnClose.VB_VarHelpID = -1
Private mFraHeader As MSForms.Frame
Private mFraGrid As MSForms.Frame
Private mMeasureLabel As MSForms.label

' Event handlers for dynamic controls
Private mHeaderEvents As Collection
Private mFilterEvents As Collection
Private mRowEvents As Collection

' Data
Private Type MoveRow
    moveName As String
    MoveType As String
    Category As String
    Power As String
    Accuracy As String
    PP As String
    Priority As String
    Description As String
End Type

Private mRows() As MoveRow
Private mRowCount As Long

' Sort + state
Private mSortCol As Long
Private mSortAsc As Boolean
Private mLastPokemon As String
Private mLastGameSel As String
Private mInFilterUpdate As Boolean


' Cached data (global tables)
Private mMoveMetaByKey As Object
Private mMoveKeyByName As Object
Private mMovesByPokemonGame As Object
Private mPokemonOptionsCache As Object
Private mGameKeyToLabel As Object
Private mTypeOptions As Variant
Private mGameOptions As Variant
Private mAllMoveKeysCache As Object
Private mCachesReady As Boolean

' =============================
' Form lifecycle
' =============================
Private Sub UserForm_Initialize()
    On Error GoTo CleanFail

    mInFilterUpdate = True
    Me.BackColor = RGB(204, 0, 0)
    On Error Resume Next
    Me.Font.name = UI_FONT_NAME
    On Error GoTo 0

    InitColumnWidths

    BuildRuntimeUI

    Dim defaultGame As String
    Dim defaultPokemon As String
    defaultGame = DefaultGameValue()
    defaultPokemon = DefaultPokemonValue()

    PopulateFilters defaultGame, defaultPokemon

    LoadData

    SetInfoLabel

    mSortCol = gcMove
    mSortAsc = True

    RenderGrid

    mInFilterUpdate = False
    Exit Sub

CleanFail:
    mInFilterUpdate = False
    MsgBox "Unable to initialize Movelist: " & Err.Description, vbExclamation
End Sub

Private Sub InitColumnWidths()
    mColW(gcMove) = 185
    mColW(gcType) = 85
    mColW(gcCategory) = 80
    mColW(gcPower) = 60
    mColW(gcAccuracy) = 70
    mColW(gcPP) = 55
    mColW(gcPriority) = 60
    mColW(gcDescription) = 550
End Sub

Private Sub BuildRuntimeUI()
    Dim x As Single, y As Single

    HideIfExists "lbMoves"
    HideIfExists "txtDescription"
    HideIfExists "lblInfo"

    Set mLblInfo = Me.Controls.Add("Forms.Label.1", "lblInfo2", True)
    With mLblInfo
        .Left = PAD
        .Top = PAD
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = 25
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .caption = "Movelist"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE + 5
        .Font.Bold = True
    End With

    y = mLblInfo.Top + mLblInfo.Height + PAD
    x = PAD

    Dim lblP As MSForms.label
    Set lblP = Me.Controls.Add("Forms.Label.1", "lblPokemon", True)
    With lblP
        .Left = x
        .Top = y + 2
        .Width = 60
        .Height = 16
        .caption = "Pokemon"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + lblP.Width + 4

    Set mCboPokemon = Me.Controls.Add("Forms.ComboBox.1", "cboPokemon", True)
    With mCboPokemon
        .Left = x
        .Top = y
        .Width = 150
        .Height = 20
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .Font.name = UI_FONT_NAME
        On Error Resume Next
        .Font.Size = UI_FONT_SIZE
        On Error GoTo 0
    End With

    x = mCboPokemon.Left + mCboPokemon.Width + 12

    Dim lblT As MSForms.label
    Set lblT = Me.Controls.Add("Forms.Label.1", "lblType", True)
    With lblT
        .Left = x
        .Top = y + 2
        .Width = 35
        .Height = 16
        .caption = "Type"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + lblT.Width + 4

    Set mCboType = Me.Controls.Add("Forms.ComboBox.1", "cboType", True)
    With mCboType
        .Left = x
        .Top = y
        .Width = 120
        .Height = 20
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .Font.name = UI_FONT_NAME
        On Error Resume Next
        .Font.Size = UI_FONT_SIZE
        On Error GoTo 0
    End With

    x = mCboType.Left + mCboType.Width + 12

    Dim lblG As MSForms.label
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
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + lblG.Width + 4

    Set mCboGame = Me.Controls.Add("Forms.ComboBox.1", "cboGame", True)
    With mCboGame
        .Left = x
        .Top = y
        .Width = 140
        .Height = 20
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .Visible = True
        .Font.name = UI_FONT_NAME
        On Error Resume Next
        .Font.Size = UI_FONT_SIZE
        On Error GoTo 0
    End With

    x = mCboGame.Left + mCboGame.Width + 12

    Set mBtnApply = Me.Controls.Add("Forms.CommandButton.1", "btnApplyFilters", True)
    With mBtnApply
        .Left = x
        .Top = y - 1
        .Width = 60
        .Height = 24
        .caption = "Apply"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With

    x = mBtnApply.Left + mBtnApply.Width + 6

    Set mBtnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClearFilters", True)
    With mBtnClear
        .Left = x
        .Top = y - 1
        .Width = 90
        .Height = 24
        .caption = "Clear Filters"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With

    Dim closeWidth As Single
    closeWidth = 70
    Set mBtnClose = Me.Controls.Add("Forms.CommandButton.1", "btnCloseML", True)
    With mBtnClose
        .Width = closeWidth
        .Height = 24
        .Top = y - 1
        .Left = Me.InsideWidth - PAD - closeWidth
        .caption = "Close"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With

    Set mFilterEvents = New Collection
    Dim e1 As CGridComboEvents, e2 As CGridComboEvents, e3 As CGridComboEvents
    Set e1 = New CGridComboEvents: e1.Init Me, mCboPokemon: mFilterEvents.Add e1
    Set e2 = New CGridComboEvents: e2.Init Me, mCboType: mFilterEvents.Add e2
    Set e3 = New CGridComboEvents: e3.Init Me, mCboGame: mFilterEvents.Add e3

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

    EnsureMeasureLabel
End Sub

Private Sub mBtnClear_Click()
    On Error GoTo CleanFail
    mInFilterUpdate = True

    mCboGame.value = FILTER_ALL
    PopulatePokemonCombo FILTER_ALL, FILTER_ALL
    mCboPokemon.value = FILTER_ALL
    mCboType.value = FILTER_ALL

    mInFilterUpdate = False
    Exit Sub

CleanFail:
    mInFilterUpdate = False
End Sub

Private Sub mBtnApply_Click()
    FiltersChanged
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub

Private Sub BuildHeaderLabels()
    Set mHeaderEvents = New Collection

    Dim captions(0 To 7) As String
    captions(gcMove) = "Move"
    captions(gcType) = "Type"
    captions(gcCategory) = "Category"
    captions(gcPower) = "Power"
    captions(gcAccuracy) = "Accuracy"
    captions(gcPP) = "PP"
    captions(gcPriority) = "Priority"
    captions(gcDescription) = "Description"

    Dim i As Long
    Dim x As Single
    x = 2

    For i = 0 To 7
        Dim h As MSForms.label
        Set h = mFraHeader.Controls.Add("Forms.Label.1", "h" & CStr(i), True)
        With h
            .Left = x
            .Top = 2
            .Width = mColW(i)
            .Height = HEADER_H
            .caption = captions(i)
            .Font.name = UI_FONT_NAME
            .Font.Size = UI_FONT_SIZE
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
    On Error GoTo CleanFail

    Dim curPokemon As String
    Dim curGame As String
    curPokemon = CleanSelection(mCboPokemon.value, FILTER_ALL)
    curGame = CleanSelection(mCboGame.value, FILTER_ALL)

    Dim pokemonChanged As Boolean
    Dim gameChanged As Boolean
    pokemonChanged = (StrComp(curPokemon, mLastPokemon, vbTextCompare) <> 0)
    gameChanged = (StrComp(curGame, mLastGameSel, vbTextCompare) <> 0)

    If gameChanged Then
        PopulatePokemonCombo curGame, curPokemon
        curPokemon = CleanSelection(mCboPokemon.value, FILTER_ALL)
    End If

    UpdateContextCells curPokemon, curGame

    If pokemonChanged Or gameChanged Then
        LoadData
        SetInfoLabel
    End If

    mLastPokemon = curPokemon
    mLastGameSel = curGame
    mInFilterUpdate = False

    RenderGrid
    Exit Sub

CleanFail:
    mInFilterUpdate = False
End Sub

Public Sub ComboTyped(ByVal ctrlName As String, ByVal typed As String)
    ' Typed filtering disabled for Movelist.
End Sub

Public Sub FilterControlChanged(ByVal ctrlName As String)
    ' Filter changes are applied explicitly via the Apply button.
End Sub

Public Sub ComboClicked(ByVal ctrlName As String)
    Dim target As MSForms.ComboBox
    Select Case LCase$(ctrlName)
        Case "cbopokemon": Set target = mCboPokemon
        Case "cbotype": Set target = mCboType
        Case "cbogame": Set target = mCboGame
        Case Else: Exit Sub
    End Select
    HighlightComboText target
End Sub

' =============================
' Context + info
' =============================
Private Sub SetInfoLabel()
    Dim pkmnDex As String
    Dim game As String
    pkmnDex = CleanSelection(mCboPokemon.value, FILTER_ALL)
    game = CleanSelection(mCboGame.value, FILTER_ALL)

    If StrComp(pkmnDex, FILTER_ALL, vbTextCompare) = 0 Then
        mLblInfo.caption = "Movelist (All Moves) (" & game & ")"
    Else
        mLblInfo.caption = "Movelist of " & pkmnDex & " (" & game & ")"
    End If
    Me.caption = mLblInfo.caption
End Sub

' =============================
' Load + filters
' =============================
Private Sub LoadData()
    EnsureDataCaches

    Dim pkmnDex As String
    Dim gameSel As String
    pkmnDex = CleanSelection(mCboPokemon.value, FILTER_ALL)
    gameSel = CleanSelection(mCboGame.value, FILTER_ALL)

    Dim allMovesMode As Boolean
    allMovesMode = (StrComp(pkmnDex, FILTER_ALL, vbTextCompare) = 0)

    Dim moveKeys As Variant
    If allMovesMode Then
        moveKeys = GetAllMoveKeysSorted(gameSel)
        BuildRowsFromMoveKeys moveKeys, FILTER_ALL, gameSel, True
    Else
        moveKeys = GetMoveKeysForPokemon(pkmnDex, gameSel)
        BuildRowsFromMoveKeys moveKeys, pkmnDex, gameSel, False
    End If
End Sub

Private Sub PopulateGameCombo(ByVal desiredSelection As String)
    Dim target As String
    target = ResolveGameLabel(desiredSelection)

    mCboGame.Clear
    mCboGame.AddItem FILTER_ALL

    If Not IsEmpty(mGameOptions) Then
        Dim i As Long
        Dim optionText As String
        For i = LBound(mGameOptions) To UBound(mGameOptions)
            optionText = Nz(mGameOptions(i))
            If Len(optionText) > 0 Then
                If StrComp(optionText, FILTER_ALL, vbTextCompare) <> 0 Then
                    mCboGame.AddItem optionText
                End If
            End If
        Next i
    End If

    EnsureComboSelection mCboGame, target
End Sub

Private Sub PopulatePokemonCombo(ByVal gameSelection As String, ByVal desiredSelection As String)
    Dim options As Variant
    options = GetPokemonOptionsForGame(gameSelection)

    mCboPokemon.Clear
    mCboPokemon.AddItem FILTER_ALL

    If Not IsEmpty(options) Then
        Dim i As Long
        Dim optionText As String
        For i = LBound(options) To UBound(options)
            optionText = Nz(options(i))
            If Len(optionText) > 0 Then
                mCboPokemon.AddItem optionText
            End If
        Next i
    End If

    EnsureComboSelection mCboPokemon, CleanSelection(desiredSelection, FILTER_ALL)
End Sub

Private Sub PopulateTypeCombo()
    mCboType.Clear
    mCboType.AddItem FILTER_ALL

    If Not IsEmpty(mTypeOptions) Then
        Dim i As Long
        For i = LBound(mTypeOptions) To UBound(mTypeOptions)
            If Len(mTypeOptions(i)) > 0 Then mCboType.AddItem mTypeOptions(i)
        Next i
    End If

    EnsureComboSelection mCboType, FILTER_ALL
End Sub

Private Function DefaultGameValue() As String
    On Error GoTo CleanFail
    DefaultGameValue = CleanSelection(Pokedex.Range("GAME").value, FILTER_ALL)
    Exit Function
CleanFail:
    DefaultGameValue = FILTER_ALL
End Function

Private Function DefaultPokemonValue() As String
    On Error GoTo CleanFail
    DefaultPokemonValue = CleanSelection(Pokedex.Range("PKMN_DEX").value, FILTER_ALL)
    Exit Function
CleanFail:
    DefaultPokemonValue = FILTER_ALL
End Function

Private Sub UpdateContextCells(ByVal Pokemon As String, ByVal game As String)
    On Error Resume Next
    Pokedex.Range("PKMN_DEX").value = IIf(StrComp(Pokemon, FILTER_ALL, vbTextCompare) = 0, FILTER_ALL, Pokemon)
    Pokedex.Range("GAME").value = game
    On Error GoTo 0
End Sub

Private Function ResolveGameLabel(ByVal rawValue As String) As String
    Dim cleaned As String
    cleaned = CleanSelection(rawValue, FILTER_ALL)
    If StrComp(cleaned, FILTER_ALL, vbTextCompare) = 0 Then
        ResolveGameLabel = FILTER_ALL
        Exit Function
    End If

    Dim key As String
    key = GameVersionKey(cleaned)
    If StrComp(key, GAME_KEY_ALL, vbTextCompare) = 0 Then
        ResolveGameLabel = FILTER_ALL
        Exit Function
    End If

    If Not mGameKeyToLabel Is Nothing Then
        If mGameKeyToLabel.Exists(key) Then
            ResolveGameLabel = CStr(mGameKeyToLabel(key))
            Exit Function
        End If
    End If

    ResolveGameLabel = cleaned
End Function

Private Sub PopulateFilters(ByVal defaultGame As String, ByVal defaultPokemon As String)
    EnsureDataCaches

    PopulateGameCombo defaultGame

    PopulatePokemonCombo CleanSelection(mCboGame.value, FILTER_ALL), defaultPokemon

    PopulateTypeCombo

    mCboType.value = FILTER_ALL

    mLastPokemon = CleanSelection(mCboPokemon.value, FILTER_ALL)
    mLastGameSel = CleanSelection(mCboGame.value, FILTER_ALL)
End Sub


' =============================
' Global table helpers
' =============================
Private Sub BuildRowsFromMoveKeys(ByVal moveKeys As Variant, ByVal pokemonName As String, _
                                  ByVal gameSelection As String, ByVal allMovesMode As Boolean)
    On Error GoTo SafeDefault
    If IsEmpty(moveKeys) Then GoTo SafeDefault

    Dim lb As Long, ub As Long
    lb = LBound(moveKeys)
    ub = UBound(moveKeys)
    If ub < lb Then GoTo SafeDefault

    ReDim mRows(1 To ub - lb + 1)
    mRowCount = 0

    Dim i As Long
    For i = lb To ub
        Dim moveKey As String
        moveKey = Nz(moveKeys(i))
        If Len(moveKey) = 0 Then GoTo ContinueLoop

        Dim meta As Variant
        meta = GetMoveMeta(moveKey)
        If IsEmpty(meta) Then GoTo ContinueLoop

        Dim row As MoveRow
        row.moveName = Nz(meta(1))
        Dim moveTypeText As String
        moveTypeText = FormatTypeName(meta(2))
        If Len(moveTypeText) = 0 Then
            moveTypeText = Nz(meta(2))
        End If
        row.MoveType = moveTypeText
        row.Category = Nz(meta(3))
        row.Power = Nz(meta(4))
        row.Accuracy = Nz(meta(5))
        row.PP = Nz(meta(6))
        row.Priority = Nz(meta(7))
        row.Description = Nz(meta(8))

        mRowCount = mRowCount + 1
        mRows(mRowCount) = row

ContinueLoop:
    Next i

    If mRowCount = 0 Then GoTo SafeDefault
    If mRowCount < UBound(mRows) Then
        ReDim Preserve mRows(1 To mRowCount)
    End If
    Exit Sub

SafeDefault:
    ReDim mRows(1 To 1)
    mRowCount = 0
End Sub

Private Sub EnsureDataCaches()
    If mCachesReady Then Exit Sub

    GlobalTables.LoadMovesTable
    GlobalTables.LoadPokemonTable
    GlobalTables.LoadGameversionsTable
    GlobalTables.LoadAssetsTable

    BuildMoveMetaIndex
    BuildPokemonMoveIndex
    BuildGameOptions

    Set mPokemonOptionsCache = CreateObject("Scripting.Dictionary")
    mPokemonOptionsCache.CompareMode = vbTextCompare

    Set mAllMoveKeysCache = CreateObject("Scripting.Dictionary")
    mAllMoveKeysCache.CompareMode = vbTextCompare
    mCachesReady = True
End Sub

Private Sub BuildMoveMetaIndex()
    mTypeOptions = Empty

    Set mMoveMetaByKey = CreateObject("Scripting.Dictionary")
    mMoveMetaByKey.CompareMode = vbTextCompare
    Set mMoveKeyByName = CreateObject("Scripting.Dictionary")
    mMoveKeyByName.CompareMode = vbTextCompare

    Dim tbl As Variant
    tbl = GlobalTables.movesTable
    If IsEmpty(tbl) Then Exit Sub

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)

    Dim moveKeyCol As Long
    Dim nameCol As Long
    Dim typeCol As Long
    Dim categoryCol As Long
    Dim powerCol As Long
    Dim accuracyCol As Long
    Dim ppCol As Long
    Dim priorityCol As Long
    Dim descCol As Long

    moveKeyCol = GlobalTables.FindHeaderColumn(tbl, "MOVE_KEY")
    nameCol = GlobalTables.FindHeaderColumn(tbl, "DISPLAY_NAME")
    typeCol = GlobalTables.FindHeaderColumn(tbl, "TYPE")
    categoryCol = GlobalTables.FindHeaderColumn(tbl, "CATEGORY")
    powerCol = GlobalTables.FindHeaderColumn(tbl, "POWER")
    accuracyCol = GlobalTables.FindHeaderColumn(tbl, "ACCURACY")
    ppCol = GlobalTables.FindHeaderColumn(tbl, "PP")
    priorityCol = GlobalTables.FindHeaderColumn(tbl, "PRIORITY")
    descCol = GlobalTables.FindHeaderColumn(tbl, "EFFECT_SHORT")

    If moveKeyCol = 0 Or nameCol = 0 Then Exit Sub

    Dim firstRow As Long
    firstRow = headerRow + 1

    Dim r As Long

    For r = firstRow To UBound(tbl, 1)
        Dim moveKey As String
        moveKey = Nz(tbl(r, moveKeyCol))
        If Len(moveKey) = 0 Then GoTo ContinueRow

        Dim meta(1 To 8) As Variant
        meta(1) = Trim$(Nz(tbl(r, nameCol)))
        Dim typeText As String
        If typeCol > 0 Then typeText = FormatTypeName(tbl(r, typeCol))
        meta(2) = typeText
        If categoryCol > 0 Then meta(3) = Trim$(Nz(tbl(r, categoryCol)))
        If powerCol > 0 Then meta(4) = Trim$(Nz(tbl(r, powerCol)))
        If accuracyCol > 0 Then meta(5) = Trim$(Nz(tbl(r, accuracyCol)))
        If ppCol > 0 Then meta(6) = Trim$(Nz(tbl(r, ppCol)))
        If priorityCol > 0 Then meta(7) = Trim$(Nz(tbl(r, priorityCol)))
        If descCol > 0 Then meta(8) = Trim$(Nz(tbl(r, descCol)))

        mMoveMetaByKey(moveKey) = meta

        Dim nameKey As String
        nameKey = NormalizeMoveNameKey(meta(1))
        If Len(nameKey) > 0 Then
            If Not mMoveKeyByName.Exists(nameKey) Then mMoveKeyByName(nameKey) = moveKey
        End If

ContinueRow:
    Next r

    mTypeOptions = CollectMoveTypeOptions(tbl)
End Sub

Private Sub BuildPokemonMoveIndex()
    Set mMovesByPokemonGame = CreateObject("Scripting.Dictionary")
    mMovesByPokemonGame.CompareMode = vbTextCompare

    Dim tbl As Variant
    tbl = GlobalTables.PokemonTable
    If IsEmpty(tbl) Then Exit Sub

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstRow As Long
    firstRow = headerRow + 1
    Dim lastRow As Long
    lastRow = UBound(tbl, 1)

    Dim nameCol As Long
    nameCol = GlobalTables.FindHeaderColumn(tbl, "DISPLAY_NAME")
    If nameCol = 0 Then Exit Sub

    Dim movesetCols As Object
    Set movesetCols = CreateObject("Scripting.Dictionary")
    movesetCols.CompareMode = vbTextCompare

    Dim c As Long
    For c = LBound(tbl, 2) To UBound(tbl, 2)
        Dim header As String
        header = Nz(tbl(headerRow, c))
        If StrComp(Left$(header, 8), "MOVESET_", vbTextCompare) = 0 Then
            Dim suffix As String
            suffix = Mid$(header, 9)
            If Len(suffix) > 0 Then
                Dim bucketKey As String
                bucketKey = MovesetSuffixToKey(suffix)
                If Len(bucketKey) > 0 Then
                    If Not movesetCols.Exists(bucketKey) Then
                        movesetCols.Add bucketKey, c
                    End If
                End If
            End If
        End If
    Next c

    If movesetCols.count = 0 Then Exit Sub

    Dim r As Long
    For r = firstRow To lastRow
        Dim pokemonName As String
        pokemonName = Nz(tbl(r, nameCol))
        If Len(pokemonName) = 0 Then GoTo ContinueRow

        Dim pk As String
        pk = pokemonKey(pokemonName)

        Dim colKey As Variant
        For Each colKey In movesetCols.keys
            Dim colIndex As Long
            colIndex = CLng(movesetCols(colKey))
            Dim movesetRaw As String
            movesetRaw = Nz(tbl(r, colIndex))
            If Len(movesetRaw) = 0 Then GoTo NextColumn

            AddMovesFromMoveset pk, CStr(colKey), movesetRaw

NextColumn:
        Next colKey

ContinueRow:
    Next r
End Sub

Private Sub BuildGameOptions()
    mGameOptions = Empty
    Set mGameKeyToLabel = Nothing

    If BuildGameOptionsFromAssetsTable() Then Exit Sub

    Dim tbl As Variant
    tbl = GlobalTables.GameversionsTable
    If IsEmpty(tbl) Then Exit Sub

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstCol As Long
    Dim lastCol As Long
    firstCol = LBound(tbl, 2)
    lastCol = UBound(tbl, 2)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim c As Long
    For c = firstCol To lastCol
        Dim header As String
        header = Nz(tbl(headerRow, c))
        If Len(header) = 0 Then GoTo ContinueCol
        If StrComp(header, "POKEMON_ALL", vbTextCompare) = 0 Then GoTo ContinueCol
        If Left$(header, 8) = "POKEMON_" Then
            Dim suffix As String
            suffix = Mid$(header, 9)
            If Len(suffix) > 0 Then
                If Not dict.Exists(suffix) Then dict.Add suffix, True
            End If
        End If
ContinueCol:
    Next c

    Dim arr As Variant
    arr = DictionaryToSortedArray(dict)
    If IsEmpty(arr) Then
        mGameOptions = Empty
    Else
        mGameOptions = arr
    End If
End Sub

Private Function BuildGameOptionsFromAssetsTable() As Boolean
    GlobalTables.LoadAssetsTable
    Dim tbl As Variant
    tbl = GlobalTables.AssetsTable
    If IsEmpty(tbl) Then Exit Function

    Dim gamesCol As Long
    gamesCol = GlobalTables.FindHeaderColumn(tbl, "GAMES")
    If gamesCol = 0 Then Exit Function

    Dim values As Variant
    values = GlobalTables.ExtractColumnValues(tbl, gamesCol, True)
    If IsEmpty(values) Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set mGameKeyToLabel = CreateObject("Scripting.Dictionary")
    mGameKeyToLabel.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(values) To UBound(values)
        Dim labelText As String
        labelText = Trim$(Nz(values(i)))
        If Len(labelText) = 0 Then GoTo ContinueRow
        If Not dict.Exists(labelText) Then
            dict.Add labelText, True

            Dim versionKey As String
            versionKey = GameVersionKey(labelText)
            If Len(versionKey) > 0 Then
                If StrComp(versionKey, GAME_KEY_ALL, vbTextCompare) <> 0 Then
                    If Not mGameKeyToLabel.Exists(versionKey) Then
                        mGameKeyToLabel.Add versionKey, labelText
                    End If
                End If
            End If
        End If
ContinueRow:
    Next i

    If dict.count = 0 Then
        Set mGameKeyToLabel = Nothing
        Exit Function
    End If

    Dim arr As Variant
    arr = DictionaryToSortedArray(dict)
    If IsEmpty(arr) Then
        Set mGameKeyToLabel = Nothing
        Exit Function
    End If

    mGameOptions = arr
    BuildGameOptionsFromAssetsTable = True
End Function

Private Function GetPokemonOptionsForGame(ByVal gameSelection As String) As Variant
    EnsureDataCaches
    If mPokemonOptionsCache Is Nothing Then
        Set mPokemonOptionsCache = CreateObject("Scripting.Dictionary")
        mPokemonOptionsCache.CompareMode = vbTextCompare
    End If

    Dim cacheKey As String
    cacheKey = GameVersionKey(gameSelection)

    If mPokemonOptionsCache.Exists(cacheKey) Then
        GetPokemonOptionsForGame = mPokemonOptionsCache(cacheKey)
        Exit Function
    End If

    Dim headerName As String
    Dim suffix As String
    suffix = DexLogic.NormalizeGameVersion(CleanSelection(gameSelection, FILTER_ALL))
    If Len(suffix) = 0 Or StrComp(suffix, FILTER_ALL, vbTextCompare) = 0 Then
        headerName = "POKEMON_ALL"
    Else
        headerName = "POKEMON_" & suffix
    End If

    Dim colIndex As Long
    colIndex = GlobalTables.FindHeaderColumn(GlobalTables.GameversionsTable, headerName)
    If colIndex = 0 Then Exit Function

    Dim columnValues As Variant
    columnValues = GlobalTables.ExtractColumnValues(GlobalTables.GameversionsTable, colIndex, True)
    If IsEmpty(columnValues) Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(columnValues) To UBound(columnValues)
        Dim valueText As String
        valueText = Nz(columnValues(i))
        If Len(valueText) > 0 And valueText <> "0" Then
            If Not dict.Exists(valueText) Then dict.Add valueText, True
        End If
    Next i

    Dim arr As Variant
    arr = DictionaryToSortedArray(dict)
    If Not IsEmpty(arr) Then
        mPokemonOptionsCache(cacheKey) = arr
        GetPokemonOptionsForGame = arr
    End If
End Function

Private Function MovesetSuffixToKey(ByVal suffix As String) As String
    Dim valueText As String
    valueText = Trim$(CStr(suffix))
    If Len(valueText) = 0 Then Exit Function

    Dim normalized As String
    normalized = DexLogic.NormalizeGameVersion(valueText)
    If Len(normalized) = 0 Then normalized = valueText

    If StrComp(normalized, FILTER_ALL, vbTextCompare) = 0 Then
        MovesetSuffixToKey = GAME_KEY_ALL
    Else
        MovesetSuffixToKey = LCase$(normalized)
    End If
End Function

Private Sub AddMovesFromMoveset(ByVal pokemonKey As String, ByVal versionKey As String, ByVal movesetRaw As String)
    If mMovesByPokemonGame Is Nothing Then Exit Sub

    Dim flattened As String
    flattened = Replace(movesetRaw, vbCrLf, ";")
    flattened = Replace(flattened, vbCr, ";")
    flattened = Replace(flattened, vbLf, ";")

    Dim tokens() As String
    tokens = Split(flattened, ";")

    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        Dim moveKey As String
        moveKey = ResolveMoveKey(tokens(i))
        If Len(moveKey) = 0 Then GoTo ContinueLoop

        AddMoveForPokemon pokemonKey, versionKey, moveKey
        If versionKey <> GAME_KEY_ALL Then
            AddMoveForPokemon pokemonKey, GAME_KEY_ALL, moveKey
        End If

ContinueLoop:
    Next i
End Sub

Private Function ResolveMoveKey(ByVal tokenValue As Variant) As String
    Dim tokenText As String
    tokenText = StripMoveTokenNotes(Nz(tokenValue))
    If Len(tokenText) = 0 Then Exit Function

    If Not mMoveMetaByKey Is Nothing Then
        If mMoveMetaByKey.Exists(tokenText) Then
            ResolveMoveKey = CStr(tokenText)
            Exit Function
        End If
    End If

    Dim nameKey As String
    nameKey = NormalizeMoveNameKey(tokenText)
    If Len(nameKey) = 0 Then Exit Function

    If Not mMoveKeyByName Is Nothing Then
        If mMoveKeyByName.Exists(nameKey) Then
            ResolveMoveKey = CStr(mMoveKeyByName(nameKey))
        End If
    End If
End Function

Private Function StripMoveTokenNotes(ByVal token As String) As String
    Dim t As String
    t = Trim$(token)
    If Len(t) = 0 Then Exit Function

    Dim parenPos As Long
    parenPos = InStr(t, "(")
    If parenPos > 0 Then
        t = Left$(t, parenPos - 1)
    End If

    StripMoveTokenNotes = Trim$(t)
End Function

Private Function NormalizeMoveNameKey(ByVal textValue As Variant) As String
    Dim t As String
    t = LCase$(Trim$(CStr(textValue)))
    If Len(t) = 0 Then Exit Function

    t = Replace(t, ChrW(&H2019), "'")
    t = Replace(t, "'", "")
    t = Replace(t, "-", " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    NormalizeMoveNameKey = Trim$(t)
End Function

Private Sub AddMoveForPokemon(ByVal pokemonKey As String, ByVal versionKey As String, ByVal moveKey As String)
    If Len(pokemonKey) = 0 Or Len(moveKey) = 0 Then Exit Sub

    Dim bucketKey As String
    bucketKey = PokemonGameKey(pokemonKey, versionKey)

    Dim dictMoves As Object
    If mMovesByPokemonGame.Exists(bucketKey) Then
        Set dictMoves = mMovesByPokemonGame(bucketKey)
    Else
        Set dictMoves = CreateObject("Scripting.Dictionary")
        dictMoves.CompareMode = vbTextCompare
        Set mMovesByPokemonGame(bucketKey) = dictMoves
    End If

    If Not dictMoves.Exists(moveKey) Then dictMoves.Add moveKey, True
End Sub

Private Function pokemonKey(ByVal name As String) As String
    pokemonKey = LCase$(Trim$(name))
End Function

Private Function PokemonGameKey(ByVal pokemonKey As String, ByVal versionKey As String) As String
    PokemonGameKey = pokemonKey & "|" & versionKey
End Function

Private Function GetMoveKeysForPokemon(ByVal pokemonName As String, ByVal gameSelection As String) As Variant
    EnsureDataCaches

    If StrComp(pokemonName, FILTER_ALL, vbTextCompare) <> 0 Then
        Dim tmpMoves As Variant
        tmpMoves = FetchTmpMoveListValues()
        If Not IsEmpty(tmpMoves) Then
            Dim tmpKeys As Variant
            tmpKeys = MoveNamesToKeys(tmpMoves)
            If Not IsEmpty(tmpKeys) Then
                GetMoveKeysForPokemon = tmpKeys
                Exit Function
            End If
        End If
    End If

    If mMovesByPokemonGame Is Nothing Then Exit Function

    Dim pKey As String
    pKey = pokemonKey(pokemonName)

    Dim gKey As String
    gKey = GameVersionKey(gameSelection)

    Dim dictMoves As Object
    Dim bucketKey As String
    bucketKey = PokemonGameKey(pKey, gKey)

    If mMovesByPokemonGame.Exists(bucketKey) Then
        Set dictMoves = mMovesByPokemonGame(bucketKey)
    ElseIf gKey <> GAME_KEY_ALL Then
        bucketKey = PokemonGameKey(pKey, GAME_KEY_ALL)
        If mMovesByPokemonGame.Exists(bucketKey) Then
            Set dictMoves = mMovesByPokemonGame(bucketKey)
        End If
    End If

    If dictMoves Is Nothing Then Exit Function

    Dim keys As Variant
    keys = dictMoves.keys
    If Not IsArray(keys) Then Exit Function

    Dim lb As Long, ub As Long
    lb = LBound(keys)
    ub = UBound(keys)
    If ub < lb Then Exit Function

    Dim arr() As String
    Dim names() As String
    ReDim arr(1 To ub - lb + 1)
    ReDim names(1 To ub - lb + 1)

    Dim idx As Long
    Dim i As Long
    For i = lb To ub
        Dim mk As String
        mk = CStr(keys(i))
        Dim meta As Variant
        meta = GetMoveMeta(mk)
        If Not IsEmpty(meta) Then
            idx = idx + 1
            arr(idx) = mk
            names(idx) = LCase$(CStr(meta(1)))
        End If
    Next i

    If idx = 0 Then Exit Function

    If idx < UBound(arr) Then
        ReDim Preserve arr(1 To idx)
        ReDim Preserve names(1 To idx)
    End If

    SortParallelByNames arr, names, 1, idx
    GetMoveKeysForPokemon = arr
End Function

Private Function GetAllMoveKeysSorted(ByVal gameSelection As String) As Variant
    EnsureDataCaches

    Dim cacheKey As String
    cacheKey = GameVersionKey(gameSelection)

    If Not mAllMoveKeysCache Is Nothing Then
        If mAllMoveKeysCache.Exists(cacheKey) Then
            GetAllMoveKeysSorted = mAllMoveKeysCache(cacheKey)
            Exit Function
        End If
    End If

    Dim headerName As String
    headerName = MovesHeaderForGame(cacheKey)

    Dim moveNames As Variant
    moveNames = GetGameversionsColumnValues(headerName)
    If IsEmpty(moveNames) And StrComp(headerName, "MOVES_ALL", vbTextCompare) <> 0 Then
        moveNames = GetGameversionsColumnValues("MOVES_ALL")
    End If
    If IsEmpty(moveNames) Then Exit Function

    Dim arr() As String
    Dim names() As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim idx As Long
    Dim i As Long
    For i = LBound(moveNames) To UBound(moveNames)
        Dim displayText As String
        displayText = Nz(moveNames(i))
        If Len(displayText) = 0 Then GoTo ContinueAllLoop

        Dim moveKey As String
        moveKey = ResolveMoveKey(displayText)
        If Len(moveKey) = 0 Then GoTo ContinueAllLoop
        If dict.Exists(moveKey) Then GoTo ContinueAllLoop

        Dim meta As Variant
        meta = GetMoveMeta(moveKey)
        If IsEmpty(meta) Then GoTo ContinueAllLoop

        dict.Add moveKey, True
        idx = idx + 1
        If idx = 1 Then
            ReDim arr(1 To 1)
            ReDim names(1 To 1)
        Else
            ReDim Preserve arr(1 To idx)
            ReDim Preserve names(1 To idx)
        End If
        arr(idx) = moveKey
        names(idx) = LCase$(CStr(meta(1)))

ContinueAllLoop:
    Next i

    If idx = 0 Then Exit Function

    SortParallelByNames arr, names, 1, idx

    If Not mAllMoveKeysCache Is Nothing Then
        mAllMoveKeysCache(cacheKey) = arr
    End If
    GetAllMoveKeysSorted = arr
End Function

Private Function GetGameversionsColumnValues(ByVal headerName As String) As Variant
    GlobalTables.LoadGameversionsTable
    If IsEmpty(GlobalTables.GameversionsTable) Then Exit Function

    Dim colIndex As Long
    colIndex = GlobalTables.FindHeaderColumn(GlobalTables.GameversionsTable, headerName)
    If colIndex = 0 Then Exit Function

    GetGameversionsColumnValues = GlobalTables.ExtractColumnValues(GlobalTables.GameversionsTable, colIndex, True)
End Function

Private Function MovesHeaderForGame(ByVal gameKey As String) As String
    If Len(gameKey) = 0 Or StrComp(gameKey, GAME_KEY_ALL, vbTextCompare) = 0 Then
        MovesHeaderForGame = "MOVES_ALL"
    Else
        MovesHeaderForGame = "MOVES_" & gameKey
    End If
End Function

Private Function FetchTmpMoveListValues() As Variant
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = Lists

    Dim headerCell As Range
    Set headerCell = ws.Rows(1).Find(What:=TMP_MOVE_HEADER, LookIn:=xlValues, _
                                     LookAt:=xlWhole, MatchCase:=False)
    If headerCell Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, headerCell.Column).End(xlUp).row
    If lastRow <= headerCell.row Then Exit Function

    Dim values() As String
    Dim count As Long
    Dim r As Long
    For r = headerCell.row + 1 To lastRow
        Dim textValue As String
        textValue = Nz(ws.Cells(r, headerCell.Column).value)
        If Len(textValue) > 0 Then
            count = count + 1
            If count = 1 Then
                ReDim values(1 To 1)
            Else
                ReDim Preserve values(1 To count)
            End If
            values(count) = textValue
        End If
    Next r

    If count > 0 Then
        FetchTmpMoveListValues = values
    End If
    Exit Function

CleanFail:
    ' fallback to Empty
End Function

Private Function MoveNamesToKeys(ByVal moveNames As Variant) As Variant
    If IsEmpty(moveNames) Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim keys() As String
    Dim names() As String
    Dim idx As Long
    Dim i As Long

    For i = LBound(moveNames) To UBound(moveNames)
        Dim nameText As String
        nameText = Nz(moveNames(i))
        If Len(nameText) = 0 Then GoTo ContinueLoop

        Dim moveKey As String
        moveKey = ResolveMoveKey(nameText)
        If Len(moveKey) = 0 Then GoTo ContinueLoop
        If dict.Exists(moveKey) Then GoTo ContinueLoop

        Dim meta As Variant
        meta = GetMoveMeta(moveKey)
        If IsEmpty(meta) Then GoTo ContinueLoop

        dict.Add moveKey, True
        idx = idx + 1
        If idx = 1 Then
            ReDim keys(1 To 1)
            ReDim names(1 To 1)
        Else
            ReDim Preserve keys(1 To idx)
            ReDim Preserve names(1 To idx)
        End If
        keys(idx) = moveKey
        names(idx) = LCase$(CStr(meta(1)))

ContinueLoop:
    Next i

    If idx = 0 Then Exit Function

    SortParallelByNames keys, names, 1, idx
    MoveNamesToKeys = keys
End Function

Private Function GetMoveMeta(ByVal moveKey As String) As Variant
    If mMoveMetaByKey Is Nothing Then Exit Function
    If mMoveMetaByKey.Exists(moveKey) Then
        GetMoveMeta = mMoveMetaByKey(moveKey)
    End If
End Function

Private Sub SortParallelByNames(ByRef keys() As String, ByRef names() As String, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub

    Dim i As Long, j As Long
    i = lo
    j = hi

    Dim pivot As String
    pivot = names((lo + hi) \ 2)

    Do While i <= j
        Do While StrComp(names(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(names(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            SwapStrings keys(i), keys(j)
            SwapStrings names(i), names(j)
            i = i + 1
            j = j - 1
        End If
    Loop

    If lo < j Then SortParallelByNames keys, names, lo, j
    If i < hi Then SortParallelByNames keys, names, i, hi
End Sub

Private Sub SwapStrings(ByRef a As String, ByRef b As String)
    Dim tmp As String
    tmp = a
    a = b
    b = tmp
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
End Sub

Private Sub AddCellLabel(ByVal name As String, ByVal caption As String, _
                         ByVal leftX As Single, ByVal topY As Single, _
                         ByVal w As Single, ByVal h As Single, _
                         ByVal center As Boolean, Optional ByVal wrap As Boolean = False)
    Dim lbl As MSForms.label
    Set lbl = mFraGrid.Controls.Add("Forms.Label.1", name, True)

    With lbl
        .Left = leftX
        .Top = topY
        .Width = w
        .Height = h
        .caption = caption
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
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
    Dim lbl As MSForms.label
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
    Dim columnWidthPts As Single
    columnWidthPts = mColW(gcDescription)

    Dim lineCount As Long
    lineCount = EstimateLineCount(desc, columnWidthPts)
    If lineCount < 1 Then lineCount = 1

    CalcRowHeight = Application.WorksheetFunction.Max(ROW_MIN_H, lineCount * (ROW_MIN_H - 2))
End Function

Private Function EstimateLineCount(ByVal desc As String, ByVal columnWidthPts As Single) As Long
    Dim normalized As String
    normalized = Replace(desc, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Dim segments As Variant
    segments = Split(normalized, vbLf)

    Dim totalLines As Long
    Dim idx As Long
    For idx = LBound(segments) To UBound(segments)
        Dim segmentText As String
        segmentText = NormalizeDescriptionSegment(segments(idx))
        totalLines = totalLines + EstimateLinesForSegment(segmentText, columnWidthPts)
    Next idx

    If totalLines <= 0 Then totalLines = 1
    EstimateLineCount = totalLines
End Function

Private Function EstimateLinesForSegment(ByVal textValue As String, ByVal columnWidthPts As Single) As Long
    Dim widthLimit As Single
    widthLimit = columnWidthPts - 6
    If widthLimit <= 0 Then widthLimit = columnWidthPts
    If widthLimit <= 0 Then widthLimit = 1

    If Len(textValue) = 0 Then
        EstimateLinesForSegment = 1
        Exit Function
    End If

    Dim tokens As Variant
    tokens = Split(textValue, " ")

    Dim currentLine As String
    Dim lines As Long
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        Dim word As String
        word = Trim$(tokens(i))
        If Len(word) = 0 Then GoTo ContinueLoop

        Dim candidate As String
        If Len(currentLine) = 0 Then
            candidate = word
        Else
            candidate = currentLine & " " & word
        End If

        If MeasureTextWidth(candidate) > widthLimit Then
            If Len(currentLine) = 0 Then
                lines = lines + LinesNeededForLongWord(word, widthLimit)
                currentLine = vbNullString
            Else
                lines = lines + 1
                currentLine = word
            End If
        Else
            currentLine = candidate
        End If

ContinueLoop:
    Next i

    If Len(currentLine) > 0 Then lines = lines + 1
    If lines = 0 Then lines = 1
    EstimateLinesForSegment = lines
End Function

Private Function LinesNeededForLongWord(ByVal word As String, ByVal widthLimit As Single) As Long
    If Len(word) = 0 Then
        LinesNeededForLongWord = 1
        Exit Function
    End If
    If widthLimit <= 0 Then widthLimit = 1

    Dim ratio As Double
    ratio = MeasureTextWidth(word) / widthLimit
    LinesNeededForLongWord = CeilingPositive(ratio)
End Function

Private Function NormalizeDescriptionSegment(ByVal textValue As String) As String
    Dim cleaned As String
    cleaned = Replace(textValue, vbTab, " ")
    Do While InStr(cleaned, "  ") > 0
        cleaned = Replace(cleaned, "  ", " ")
    Loop
    NormalizeDescriptionSegment = Trim$(cleaned)
End Function

Private Function CeilingPositive(ByVal value As Double) As Long
    Dim floored As Long
    floored = Int(value)
    If value > floored Then
        CeilingPositive = floored + 1
    Else
        CeilingPositive = floored
    End If
    If CeilingPositive < 1 Then CeilingPositive = 1
End Function

Private Sub EnsureMeasureLabel()
    If Not mMeasureLabel Is Nothing Then Exit Sub
    Set mMeasureLabel = Me.Controls.Add("Forms.Label.1", "lblMeasureHidden", True)
    With mMeasureLabel
        .Visible = False
        .WordWrap = False
        .AutoSize = True
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
End Sub

Private Function MeasureTextWidth(ByVal textValue As String) As Single
    EnsureMeasureLabel
    mMeasureLabel.caption = textValue
    MeasureTextWidth = mMeasureLabel.Width
End Function

' =============================
' Filtering + sorting
' =============================
Private Function GetFilteredIndices() As Long()
    Dim pokeSel As String, typeSel As String, gameSel As String
    pokeSel = Trim$(CStr(mCboPokemon.value))
    typeSel = Trim$(CStr(mCboType.value))
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
' Helpers
' =============================
Private Function Nz(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        Nz = ""
    Else
        Nz = Trim$(CStr(v))
    End If
End Function

