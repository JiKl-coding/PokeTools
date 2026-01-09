VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pokelist 
   Caption         =   "Pokelist"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23760
   OleObjectBlob   =   "Pokelist.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Pokelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'===============================
' UserForm: Pokelist
' Custom Grid (Label + Frame)
'===============================
Option Explicit

Private Const FILTER_ALL As String = "All"
Private Const GAME_KEY_ALL As String = "__all__"
Private Const UI_FONT_NAME As String = "Aptos Narrow"
Private Const UI_FONT_SIZE As Integer = 12

' Visual layout (points)
Private Const PAD As Single = 6
Private Const HEADER_H As Single = 18
Private Const ROW_MIN_H As Single = 22

' Column indices
Private Enum GridCol
    gcDexId = 0
    gcPokemon = 1
    gcForm = 2
    gcType1 = 3
    gcType2 = 4
    gcHP = 5
    gcAtk = 6
    gcDef = 7
    gcSpA = 8
    gcSpD = 9
    gcSpe = 10
    gcTotal = 11
    gcAbilities = 12
End Enum

' Column widths
Private mColW(0 To 12) As Single

' Runtime UI
Private mLblInfo As MSForms.label
Private mCboType As MSForms.ComboBox
Private mCboAbility As MSForms.ComboBox
Private mCboGame As MSForms.ComboBox
Private WithEvents mBtnApply As MSForms.CommandButton
Attribute mBtnApply.VB_VarHelpID = -1
Private WithEvents mBtnClear As MSForms.CommandButton
Attribute mBtnClear.VB_VarHelpID = -1
Private WithEvents mBtnClose As MSForms.CommandButton
Attribute mBtnClose.VB_VarHelpID = -1
Private mFraHeader As MSForms.Frame
Private mFraGrid As MSForms.Frame

' Event handlers for dynamic controls
Private mHeaderEvents As Collection
Private mFilterEvents As Collection
Private mRowEvents As Collection

Private Type PokemonColumnMap
    DexId As Long
    Pokemon As Long
    Form As Long
    Type1 As Long
    Type2 As Long
    HP As Long
    Attack As Long
    Defense As Long
    SpAtt As Long
    SpDef As Long
    Speed As Long
    Total As Long
    Ability1 As Long
    Ability2 As Long
    AbilityHidden As Long
End Type

' Data
Private mPokemonCols As PokemonColumnMap

Private Type PokeRow
    DexId As Long
    Pokemon As String
    Form As String
    Type1 As String
    Type2 As String
    HP As Long
    Attack As Long
    Defense As Long
    SpAtt As Long
    SpDef As Long
    Speed As Long
    Total As Long
    Ability1 As String
    Ability2 As String
    AbilityHidden As String
    AbilitiesDisplay As String
End Type

Private mRows() As PokeRow
Private mRowCount As Long

 ' Sort state
Private mSortCol As Long
Private mSortAsc As Boolean

' Last selections to detect changes
Private mLastGameSel As String
Private mInFilterUpdate As Boolean

' Ability caches for fast dropdown refresh
Private mAbilitiesByType As Object         ' Dictionary: type -> Dictionary set
Private mAbilityLists As Object            ' GAMEVERSIONS ability lists keyed by version

' Cached tables + options
Private mMovesetColumns As Object
Private mTypeOptions As Variant
Private mGameOptions As Variant
Private mGameKeyToLabel As Object
Private mCachesReady As Boolean

' =============================
' Form init
' =============================
Private Sub UserForm_Initialize()
    On Error GoTo CleanFail

    mInFilterUpdate = True

    Me.BackColor = RGB(204, 0, 0)
    On Error Resume Next
    Me.Font.name = UI_FONT_NAME
    Me.Font.Size = UI_FONT_SIZE
    On Error GoTo 0

    InitColumnWidths
    BuildRuntimeUI

    Dim defaultGame As String
    defaultGame = DefaultGameValue()

    LoadData
    PopulateFilters defaultGame, True
    SetInfoLabel

    mSortCol = gcPokemon
    mSortAsc = True

    RenderGrid

    mInFilterUpdate = False
    Exit Sub

CleanFail:
    mInFilterUpdate = False
    MsgBox "Unable to initialize Pokelist: " & Err.Description, vbExclamation
End Sub

Private Sub InitColumnWidths()
    ' DexId | Pokemon | Form | Type1 | Type2 | HP | Atk | Def | SpA | SpD | Spe | Total | Abilities
    mColW(gcDexId) = 50
    mColW(gcPokemon) = 120
    mColW(gcForm) = 70
    mColW(gcType1) = 70
    mColW(gcType2) = 70
    mColW(gcHP) = 45
    mColW(gcAtk) = 50
    mColW(gcDef) = 55
    mColW(gcSpA) = 60
    mColW(gcSpD) = 60
    mColW(gcSpe) = 50
    mColW(gcTotal) = 55
    mColW(gcAbilities) = 260
End Sub

Private Sub BuildRuntimeUI()
    Dim x As Single, y As Single

    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        ctrl.Visible = False
    Next ctrl

    ' Info label
    Set mLblInfo = Me.Controls.Add("Forms.Label.1", "lblInfoPL", True)
    With mLblInfo
        .Left = PAD
        .Top = PAD
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = 25
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .caption = "Pokelist"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE + 5
        .Font.Bold = True
    End With

    ' Filters row
    y = mLblInfo.Top + mLblInfo.Height + PAD
    x = PAD

    ' Type filter
    Dim lblT As MSForms.label
    Set lblT = Me.Controls.Add("Forms.Label.1", "lblTypePL", True)
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
    Set mCboType = Me.Controls.Add("Forms.ComboBox.1", "cboTypePL", True)
    With mCboType
        .Left = x
        .Top = y
        .Width = 120
        .Height = 20
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        On Error Resume Next
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
        On Error GoTo 0
    End With

    x = mCboType.Left + mCboType.Width + 12

    ' Ability filter
    Dim lblA As MSForms.label
    Set lblA = Me.Controls.Add("Forms.Label.1", "lblAbilityPL", True)
    With lblA
        .Left = x
        .Top = y + 2
        .Width = 45
        .Height = 16
        .caption = "Ability"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + lblA.Width + 4
    Set mCboAbility = Me.Controls.Add("Forms.ComboBox.1", "cboAbilityPL", True)
    With mCboAbility
        .Left = x
        .Top = y
        .Width = 160
        .Height = 20
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        On Error Resume Next
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
        On Error GoTo 0
    End With

    x = mCboAbility.Left + mCboAbility.Width + 12

    ' Game filter
    Dim lblG As MSForms.label
    Set lblG = Me.Controls.Add("Forms.Label.1", "lblGamePL", True)
    With lblG
        .Left = x
        .Top = y + 2
        .Width = 45
        .Height = 16
        .caption = "Game"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + lblG.Width + 4
    Set mCboGame = Me.Controls.Add("Forms.ComboBox.1", "cboGamePL", True)
    With mCboGame
        .Left = x
        .Top = y
        .Width = 140
        .Height = 20
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        On Error Resume Next
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
        On Error GoTo 0
    End With

    x = mCboGame.Left + mCboGame.Width + 12

    Set mBtnApply = Me.Controls.Add("Forms.CommandButton.1", "btnApplyPL", True)
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

    Set mBtnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClearPL", True)
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
    Set mBtnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClosePL", True)
    With mBtnClose
        .Width = closeWidth
        .Height = 24
        .Top = y - 1
        .Left = Me.InsideWidth - PAD - closeWidth
        .caption = "Close"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With

    ' Events wiring
    Set mFilterEvents = New Collection
    Dim eT As CGridComboEvents, eA As CGridComboEvents, eG As CGridComboEvents
    Set eT = New CGridComboEvents: eT.Init Me, mCboType: mFilterEvents.Add eT
    Set eA = New CGridComboEvents: eA.Init Me, mCboAbility: mFilterEvents.Add eA
    Set eG = New CGridComboEvents: eG.Init Me, mCboGame: mFilterEvents.Add eG

    ' Header frame
    y = y + 22 + PAD
    Set mFraHeader = Me.Controls.Add("Forms.Frame.1", "fraHeaderPL", True)
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

    ' Grid frame
    y = mFraHeader.Top + mFraHeader.Height
    Set mFraGrid = Me.Controls.Add("Forms.Frame.1", "fraGridPL", True)
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

Private Sub mBtnClear_Click()
    On Error GoTo CleanFail
    mInFilterUpdate = True

    mCboGame.value = FILTER_ALL
    mCboType.value = FILTER_ALL
    mCboAbility.value = FILTER_ALL
    PopulateAbilityFilterFromRows

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

    Dim captions(0 To 12) As String
    captions(gcDexId) = "DexId"
    captions(gcPokemon) = "Pokemon"
    captions(gcForm) = "Form"
    captions(gcType1) = "Type1"
    captions(gcType2) = "Type2"
    captions(gcHP) = "HP"
    captions(gcAtk) = "Attack"
    captions(gcDef) = "Defense"
    captions(gcSpA) = "Sp. Att"
    captions(gcSpD) = "Sp. Def"
    captions(gcSpe) = "Speed"
    captions(gcTotal) = "Total"
    captions(gcAbilities) = "Abilities"

    Dim i As Long
    Dim x As Single: x = 2

    For i = 0 To 12
        Dim h As MSForms.label
        Set h = mFraHeader.Controls.Add("Forms.Label.1", "hPL" & CStr(i), True)
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

    Dim curGame As String
    curGame = CleanSelection(mCboGame.value, FILTER_ALL)

    Dim gameChanged As Boolean
    gameChanged = (StrComp(curGame, mLastGameSel, vbTextCompare) <> 0)

    If gameChanged Then
        UpdateContextGame curGame
        LoadData
        SetInfoLabel
    End If

    Dim prevAbility As String
    prevAbility = CleanSelection(mCboAbility.value, FILTER_ALL)
    PopulateAbilityFilterFromRows
    If ComboContains(mCboAbility, prevAbility) Then
        mCboAbility.value = prevAbility
    ElseIf mCboAbility.ListCount > 0 Then
        mCboAbility.ListIndex = 0
    End If

    RenderGrid

    mLastGameSel = curGame
    mInFilterUpdate = False
    Exit Sub

CleanFail:
    mInFilterUpdate = False
End Sub

' =============================
' Context + info
' =============================
Private Sub SetInfoLabel()
    Dim game As String
    game = Trim$(CStr(Pokedex.Range("GAME").value))
    mLblInfo.caption = "Pokelist (" & game & ")"
    Me.caption = mLblInfo.caption
End Sub

' =============================
' Load + filters
' =============================
Private Sub LoadData()
    On Error GoTo LoadDataFail
    EnsureDataCaches

    Dim tbl As Variant
    tbl = GlobalTables.PokemonTable
    If IsEmpty(tbl) Then
        ReDim mRows(1 To 1)
        mRowCount = 0
        Exit Sub
    End If

    ' Ability caches drive the Ability filter combo
    Set mAbilitiesByType = CreateObject("Scripting.Dictionary")
    mAbilitiesByType.CompareMode = vbTextCompare

    Dim curGameLabel As String
    curGameLabel = CleanSelection(Pokedex.Range("GAME").value, FILTER_ALL)
    Dim curGameKey As String
    curGameKey = GameVersionKey(curGameLabel)
    Dim movesetCol As Long
    movesetCol = MovesetColumnForKey(curGameKey)

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstRow As Long
    firstRow = headerRow + 1
    Dim lastRow As Long
    lastRow = UBound(tbl, 1)
    If lastRow < firstRow Then
        ReDim mRows(1 To 1)
        mRowCount = 0
        Exit Sub
    End If

    Dim capacity As Long
    capacity = lastRow - headerRow
    ReDim mRows(1 To capacity)
    mRowCount = 0

    Dim r As Long
    For r = firstRow To lastRow
        Dim displayName As String
        displayName = Nz(tbl(r, mPokemonCols.Pokemon))
        If Len(displayName) = 0 Then GoTo ContinueRow

        Dim includeRow As Boolean
        includeRow = True
        If curGameKey <> GAME_KEY_ALL And movesetCol > 0 Then
            includeRow = (Len(Nz(tbl(r, movesetCol))) > 0)
        End If

        If includeRow Then
            mRowCount = mRowCount + 1
            Dim pr As PokeRow
            pr.DexId = SafeToLong(tbl(r, mPokemonCols.DexId))
            pr.Pokemon = displayName
            pr.Form = Nz(tbl(r, mPokemonCols.Form))
            Dim type1Text As String
            Dim type2Text As String
            type1Text = FormatTypeName(tbl(r, mPokemonCols.Type1))
            type2Text = FormatTypeName(tbl(r, mPokemonCols.Type2))
            If Len(type1Text) = 0 Then type1Text = Nz(tbl(r, mPokemonCols.Type1))
            If Len(type2Text) = 0 Then type2Text = Nz(tbl(r, mPokemonCols.Type2))
            pr.Type1 = type1Text
            pr.Type2 = type2Text
            pr.HP = SafeToLong(tbl(r, mPokemonCols.HP))
            pr.Attack = SafeToLong(tbl(r, mPokemonCols.Attack))
            pr.Defense = SafeToLong(tbl(r, mPokemonCols.Defense))
            pr.SpAtt = SafeToLong(tbl(r, mPokemonCols.SpAtt))
            pr.SpDef = SafeToLong(tbl(r, mPokemonCols.SpDef))
            pr.Speed = SafeToLong(tbl(r, mPokemonCols.Speed))
            pr.Total = SafeToLong(tbl(r, mPokemonCols.Total))
            pr.Ability1 = Nz(tbl(r, mPokemonCols.Ability1))
            pr.Ability2 = Nz(tbl(r, mPokemonCols.Ability2))
            pr.AbilityHidden = Nz(tbl(r, mPokemonCols.AbilityHidden))
            pr.AbilitiesDisplay = BuildAbilitiesText(pr.Ability1, pr.Ability2, pr.AbilityHidden)

            mRows(mRowCount) = pr

            If Len(pr.Type1) > 0 Then AbilCacheAddToType pr.Type1, pr.Ability1, pr.Ability2, pr.AbilityHidden
            If Len(pr.Type2) > 0 Then AbilCacheAddToType pr.Type2, pr.Ability1, pr.Ability2, pr.AbilityHidden
        End If

ContinueRow:
    Next r

    If mRowCount = 0 Then
        ReDim mRows(1 To 1)
    ElseIf mRowCount < capacity Then
        ReDim Preserve mRows(1 To mRowCount)
    End If
    Exit Sub

LoadDataFail:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Removed: GetPokemonNamesFromListsO (dataset now uses MOVESET rule directly)


Private Function BuildAbilitiesText(ByVal a1 As String, ByVal a2 As String, ByVal ah As String) As String
    Dim parts As String
    If Len(a1) > 0 And a1 <> "-" Then parts = parts & "1: " & a1 & vbNewLine
    If Len(a2) > 0 And a2 <> "-" Then parts = parts & "2: " & a2 & vbNewLine
    If Len(ah) > 0 And ah <> "-" Then parts = parts & "Hidden: " & ah
    If Right$(parts, 2) = vbNewLine Then parts = Left$(parts, Len(parts) - 2)
    BuildAbilitiesText = parts
End Function

Private Sub PopulateFilters(ByVal defaultGame As String, ByVal initial As Boolean)
    EnsureDataCaches

    PopulateGameCombo defaultGame
    PopulateTypeCombo
    PopulateAbilityFilterFromRows

    If initial Then
        EnsureComboSelection mCboType, FILTER_ALL
        If mCboAbility.ListCount > 0 Then mCboAbility.ListIndex = 0
    End If

    mLastGameSel = CleanSelection(mCboGame.value, FILTER_ALL)
End Sub

Private Sub PopulateAbilityFilterFromRows()
    Dim gameKey As String
    Dim typeSel As String
    Dim baseList As Variant
    Dim filteredList As Variant
    Dim i As Long

    gameKey = GameVersionKey(CleanSelection(mCboGame.value, FILTER_ALL))
    baseList = AbilityListForGame(gameKey)
    typeSel = CleanSelection(mCboType.value, FILTER_ALL)
    filteredList = FilterAbilityListForType(baseList, typeSel)

    mCboAbility.Clear
    mCboAbility.AddItem FILTER_ALL

    If HasArrayValues(filteredList) Then
        For i = LBound(filteredList) To UBound(filteredList)
            mCboAbility.AddItem CStr(filteredList(i))
        Next i
    End If

End Sub

Private Function AbilityListForGame(ByVal gameKey As String) As Variant
    Dim lookupKey As String
    lookupKey = gameKey
    If Len(lookupKey) = 0 Then lookupKey = GAME_KEY_ALL

    If mAbilityLists Is Nothing Then
        AbilityListForGame = Empty
        Exit Function
    End If

    If mAbilityLists.Exists(lookupKey) Then
        AbilityListForGame = mAbilityLists(lookupKey)
    ElseIf lookupKey <> GAME_KEY_ALL And mAbilityLists.Exists(GAME_KEY_ALL) Then
        AbilityListForGame = mAbilityLists(GAME_KEY_ALL)
    Else
        AbilityListForGame = Empty
    End If
End Function

Private Function FilterAbilityListForType(ByVal baseList As Variant, ByVal typeSel As String) As Variant
    If Not HasArrayValues(baseList) Then Exit Function
    If Len(typeSel) = 0 Or StrComp(typeSel, FILTER_ALL, vbTextCompare) = 0 Then
        FilterAbilityListForType = baseList
        Exit Function
    End If

    If mAbilitiesByType Is Nothing Then
        FilterAbilityListForType = baseList
        Exit Function
    End If

    If Not mAbilitiesByType.Exists(typeSel) Then
        FilterAbilityListForType = baseList
        Exit Function
    End If

    Dim typeDict As Object
    Set typeDict = mAbilitiesByType(typeSel)

    Dim results() As String
    Dim count As Long
    Dim i As Long
    For i = LBound(baseList) To UBound(baseList)
        Dim abilityName As String
        abilityName = CStr(baseList(i))
        If typeDict.Exists(abilityName) Then
            count = count + 1
            If count = 1 Then
                ReDim results(1 To 1)
            Else
                ReDim Preserve results(1 To count)
            End If
            results(count) = abilityName
        End If
    Next i

    If count = 0 Then
        FilterAbilityListForType = baseList
    Else
        FilterAbilityListForType = results
    End If
End Function

' Simple in-place ascending sort for string array (1-D Variant)
Private Sub AbilCacheAdd(ByVal dict As Object, ByVal v As String)
    Dim t As String
    t = Trim$(CStr(v))
    If Len(t) = 0 Or t = "-" Or t = "0" Then Exit Sub
    If Not dict.Exists(t) Then dict.Add t, True
End Sub

Private Sub AbilCacheAddToType(ByVal typeName As String, ByVal a1 As String, ByVal a2 As String, ByVal ah As String)
    Dim key As String
    key = Trim$(CStr(typeName))
    If Len(key) = 0 Then Exit Sub
    Dim setRef As Object
    If mAbilitiesByType.Exists(key) Then
        Set setRef = mAbilitiesByType(key)
    Else
        Set setRef = CreateObject("Scripting.Dictionary")
        setRef.CompareMode = vbTextCompare
        mAbilitiesByType.Add key, setRef
    End If
    AbilCacheAdd setRef, a1
    AbilCacheAdd setRef, a2
    AbilCacheAdd setRef, ah
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

Private Sub SetGameSelectionFromContext(ByVal desiredValue As String)
    Dim target As String
    target = CleanSelection(desiredValue, FILTER_ALL)

    If Len(target) = 0 And mCboGame.ListCount > 0 Then
        mCboGame.ListIndex = 0
        Exit Sub
    End If

    If Len(target) = 0 Then target = FILTER_ALL

    EnsureComboHasValue mCboGame, target, False

    If ComboContains(mCboGame, target) Then
        mCboGame.value = target
    ElseIf mCboGame.ListCount > 0 Then
        mCboGame.ListIndex = 0
    End If
End Sub

' =============================
' Typed filtering (UI-only)
' =============================
Public Sub ComboTyped(ByVal ctrlName As String, ByVal typedText As String)
    ' Typed filtering disabled for Pokelist.
End Sub

Public Sub FilterControlChanged(ByVal ctrlName As String)
    ' Changes are committed explicitly via Apply.
End Sub

Public Sub ComboClicked(ByVal ctrlName As String)
    Dim target As MSForms.ComboBox
    Select Case True
        Case StrComp(ctrlName, mCboType.name, vbTextCompare) = 0
            Set target = mCboType
        Case StrComp(ctrlName, mCboAbility.name, vbTextCompare) = 0
            Set target = mCboAbility
        Case StrComp(ctrlName, mCboGame.name, vbTextCompare) = 0
            Set target = mCboGame
        Case Else
            Exit Sub
    End Select
    HighlightComboText target
End Sub

Private Function HasArrayValues(ByVal arr As Variant) As Boolean
    On Error Resume Next
    If Not IsArray(arr) Then
        HasArrayValues = False
    Else
        Dim ub As Long
        ub = UBound(arr)
        HasArrayValues = (Err.Number = 0 And ub >= LBound(arr))
        Err.Clear
    End If
    On Error GoTo 0
End Function

' =============================
' Rendering
' =============================
Private Sub RenderGrid()
    Dim prevSU As Boolean
    Dim filteredIdx() As Long
    Dim y As Single
    Dim i As Long
    Dim idx As Long
    Dim abilityText As String
    Dim rh As Single

    On Error GoTo RenderGridFail

    prevSU = Application.ScreenUpdating

    On Error Resume Next
    Application.ScreenUpdating = False
    On Error GoTo RenderGridFail

    mFraGrid.Visible = False
    ClearGridRows
    Set mRowEvents = New Collection
    If mRowCount <= 0 Then GoTo CleanExit

    filteredIdx = GetFilteredIndices()
    If UBound(filteredIdx) = 0 Then GoTo CleanExit

    SortIndices filteredIdx, 1, UBound(filteredIdx)

    y = 2
    For i = 1 To UBound(filteredIdx)
        idx = filteredIdx(i)
        abilityText = mRows(idx).AbilitiesDisplay
        rh = CalcRowHeight(abilityText)
        AddGridRow idx, y, rh
        y = y + rh
    Next i

    mFraGrid.ScrollHeight = y + 4
    mFraGrid.Visible = True

CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = prevSU
    On Error GoTo 0
    Exit Sub

RenderGridFail:
    Resume CleanExit
End Sub

Private Sub ClearGridRows()
    On Error Resume Next
    Dim c As MSForms.Control
    Dim toRemove As Collection: Set toRemove = New Collection
    For Each c In mFraGrid.Controls
        If Left$(c.name, 3) = "r__" Then toRemove.Add c.name
    Next c
    Dim n As Variant
    For Each n In toRemove
        mFraGrid.Controls.Remove CStr(n)
    Next n
    On Error GoTo 0
End Sub

Private Sub AddGridRow(ByVal rowIndex As Long, ByVal topY As Single, ByVal rowH As Single)
    Dim x As Single: x = 2
    Dim row As PokeRow: row = mRows(rowIndex)

    AddCellLabel "r__dx" & rowIndex, CStr(row.DexId), x, topY, mColW(gcDexId), rowH, True: x = x + mColW(gcDexId)
    AddCellLabel "r__nm" & rowIndex, row.Pokemon, x, topY, mColW(gcPokemon), rowH, False: x = x + mColW(gcPokemon)
    AttachRowEvent "r__nm" & rowIndex, rowIndex
    AddCellLabel "r__fm" & rowIndex, row.Form, x, topY, mColW(gcForm), rowH, False: x = x + mColW(gcForm)
    AddCellLabel "r__t1" & rowIndex, row.Type1, x, topY, mColW(gcType1), rowH, False: x = x + mColW(gcType1)
    AddCellLabel "r__t2" & rowIndex, row.Type2, x, topY, mColW(gcType2), rowH, False: x = x + mColW(gcType2)
    AddCellLabel "r__hp" & rowIndex, CStr(row.HP), x, topY, mColW(gcHP), rowH, True: x = x + mColW(gcHP)
    AddCellLabel "r__at" & rowIndex, CStr(row.Attack), x, topY, mColW(gcAtk), rowH, True: x = x + mColW(gcAtk)
    AddCellLabel "r__df" & rowIndex, CStr(row.Defense), x, topY, mColW(gcDef), rowH, True: x = x + mColW(gcDef)
    AddCellLabel "r__sa" & rowIndex, CStr(row.SpAtt), x, topY, mColW(gcSpA), rowH, True: x = x + mColW(gcSpA)
    AddCellLabel "r__sd" & rowIndex, CStr(row.SpDef), x, topY, mColW(gcSpD), rowH, True: x = x + mColW(gcSpD)
    AddCellLabel "r__sp" & rowIndex, CStr(row.Speed), x, topY, mColW(gcSpe), rowH, True: x = x + mColW(gcSpe)
    AddCellLabel "r__tt" & rowIndex, CStr(row.Total), x, topY, mColW(gcTotal), rowH, True: x = x + mColW(gcTotal)
    AddCellLabel "r__ab" & rowIndex, row.AbilitiesDisplay, x, topY, mColW(gcAbilities), rowH, False, True
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
    Dim pname As String
    pname = mRows(rowIndex).Pokemon
    On Error Resume Next
    Pokedex.Range("PKMN_DEX").value = pname
    On Error GoTo 0
    Unload Me
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
        If center Then .TextAlign = fmTextAlignCenter Else .TextAlign = fmTextAlignLeft
    End With
End Sub

Private Function CalcRowHeight(ByVal abilitiesText As String) As Single
    Const CHARS_PER_LINE As Long = 42
    Dim lines As Long
    If Len(abilitiesText) = 0 Then
        lines = 1
    Else
        Dim normalized As String
        normalized = Replace$(abilitiesText, vbNewLine, " ")
        lines = (Len(normalized) + CHARS_PER_LINE - 1) \ CHARS_PER_LINE
        lines = Application.WorksheetFunction.Max(lines, AbilityLineCount(abilitiesText))
        If lines < 1 Then lines = 1
        If lines > 6 Then lines = 6
    End If
    CalcRowHeight = Application.WorksheetFunction.Max(ROW_MIN_H, (ROW_MIN_H - 2) * lines)
End Function

Private Function AbilityLineCount(ByVal abilitiesText As String) As Long
    If Len(abilitiesText) = 0 Then
        AbilityLineCount = 1
        Exit Function
    End If

    Dim parts() As String
    parts = Split(abilitiesText, vbNewLine)
    AbilityLineCount = (UBound(parts) - LBound(parts)) + 1
End Function

' =============================
' Filtering + sorting
' =============================
Private Function GetFilteredIndices() As Long()
    Dim typeSel As String, abilitySel As String
    typeSel = Trim$(CStr(mCboType.value))
    abilitySel = Trim$(CStr(mCboAbility.value))

    Dim tmp() As Long: ReDim tmp(1 To mRowCount)
    Dim n As Long: n = 0

    Dim i As Long
    For i = 1 To mRowCount
        Dim ok As Boolean: ok = True
        If ok And Len(typeSel) > 0 And StrComp(typeSel, FILTER_ALL, vbTextCompare) <> 0 Then
            If Not (StrComp(mRows(i).Type1, typeSel, vbTextCompare) = 0 _
                 Or StrComp(mRows(i).Type2, typeSel, vbTextCompare) = 0) Then ok = False
        End If
        If ok And Len(abilitySel) > 0 And StrComp(abilitySel, FILTER_ALL, vbTextCompare) <> 0 Then
            If Not (StrComp(mRows(i).Ability1, abilitySel, vbTextCompare) = 0 _
                 Or StrComp(mRows(i).Ability2, abilitySel, vbTextCompare) = 0 _
                 Or StrComp(mRows(i).AbilityHidden, abilitySel, vbTextCompare) = 0) Then ok = False
        End If
        If ok Then n = n + 1: tmp(n) = i
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
    Dim i As Long: i = lo
    Dim j As Long: j = hi
    Dim pivot As Variant: pivot = SortKeyForIndex(idx((lo + hi) \ 2))
    Do While i <= j
        Do While CompareKeys(SortKeyForIndex(idx(i)), pivot) < 0: i = i + 1: Loop
        Do While CompareKeys(SortKeyForIndex(idx(j)), pivot) > 0: j = j - 1: Loop
        If i <= j Then
            Dim t As Long: t = idx(i): idx(i) = idx(j): idx(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortIndices idx, lo, j
    If i < hi Then SortIndices idx, i, hi
End Sub

Private Function SortKeyForIndex(ByVal i As Long) As Variant
    Select Case mSortCol
        Case gcDexId: SortKeyForIndex = SafeToLong(mRows(i).DexId)
        Case gcPokemon: SortKeyForIndex = LCase$(mRows(i).Pokemon)
        Case gcForm: SortKeyForIndex = LCase$(mRows(i).Form)
        Case gcType1: SortKeyForIndex = LCase$(mRows(i).Type1)
        Case gcType2: SortKeyForIndex = LCase$(mRows(i).Type2)
        Case gcHP: SortKeyForIndex = SafeToLong(mRows(i).HP)
        Case gcAtk: SortKeyForIndex = SafeToLong(mRows(i).Attack)
        Case gcDef: SortKeyForIndex = SafeToLong(mRows(i).Defense)
        Case gcSpA: SortKeyForIndex = SafeToLong(mRows(i).SpAtt)
        Case gcSpD: SortKeyForIndex = SafeToLong(mRows(i).SpDef)
        Case gcSpe: SortKeyForIndex = SafeToLong(mRows(i).Speed)
        Case gcTotal: SortKeyForIndex = SafeToLong(mRows(i).Total)
        Case gcAbilities: SortKeyForIndex = LCase$(mRows(i).AbilitiesDisplay)
        Case Else: SortKeyForIndex = LCase$(mRows(i).Pokemon)
    End Select
End Function

Private Function CompareKeys(ByVal a As Variant, ByVal b As Variant) As Long
    Dim r As Long
    If IsNumeric(a) And IsNumeric(b) Then
        Dim aLong As Long, bLong As Long
        aLong = SafeToLong(a)
        bLong = SafeToLong(b)
        If aLong < bLong Then
            r = -1
        ElseIf aLong > bLong Then
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
' Context helpers
' =============================
Private Function DefaultGameValue() As String
    On Error GoTo CleanFail
    DefaultGameValue = CleanSelection(Pokedex.Range("GAME").value, FILTER_ALL)
    Exit Function
CleanFail:
    DefaultGameValue = FILTER_ALL
End Function

Private Sub UpdateContextGame(ByVal gameSelection As String)
    On Error Resume Next
    Pokedex.Range("GAME").value = CleanSelection(gameSelection, FILTER_ALL)
    On Error GoTo 0
End Sub

Private Sub EnsureDataCaches()
    If mCachesReady Then Exit Sub

    GlobalTables.LoadPokemonTable
    GlobalTables.LoadGameversionsTable
    GlobalTables.LoadAssetsTable

    BuildPokemonColumnMap
    BuildMovesetColumnMap
    BuildTypeOptions
    BuildGameOptions
    BuildAbilityLists

    mCachesReady = True
End Sub

Private Sub BuildPokemonColumnMap()
    Dim tbl As Variant
    tbl = GlobalTables.PokemonTable
    If IsEmpty(tbl) Then Exit Sub

    With mPokemonCols
        .DexId = GlobalTables.FindHeaderColumn(tbl, "DEX_ID")
        .Pokemon = GlobalTables.FindHeaderColumn(tbl, "DISPLAY_NAME")
        .Form = GlobalTables.FindHeaderColumn(tbl, "FORM_GROUP")
        .Type1 = GlobalTables.FindHeaderColumn(tbl, "TYPE1")
        .Type2 = GlobalTables.FindHeaderColumn(tbl, "TYPE2")
        .HP = GlobalTables.FindHeaderColumn(tbl, "HP")
        .Attack = GlobalTables.FindHeaderColumn(tbl, "ATK")
        .Defense = GlobalTables.FindHeaderColumn(tbl, "DEF")
        .SpAtt = GlobalTables.FindHeaderColumn(tbl, "SPA")
        .SpDef = GlobalTables.FindHeaderColumn(tbl, "SPD")
        .Speed = GlobalTables.FindHeaderColumn(tbl, "SPE")
        .Total = GlobalTables.FindHeaderColumn(tbl, "TOTAL")
        .Ability1 = GlobalTables.FindHeaderColumn(tbl, "ABILITY1")
        .Ability2 = GlobalTables.FindHeaderColumn(tbl, "ABILITY2")
        .AbilityHidden = GlobalTables.FindHeaderColumn(tbl, "HIDDEN_ABILITY")
    End With
End Sub

Private Sub BuildMovesetColumnMap()
    Dim tbl As Variant
    tbl = GlobalTables.PokemonTable
    If IsEmpty(tbl) Then Exit Sub

    Set mMovesetColumns = CreateObject("Scripting.Dictionary")
    mMovesetColumns.CompareMode = vbTextCompare

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstCol As Long
    Dim lastCol As Long
    firstCol = LBound(tbl, 2)
    lastCol = UBound(tbl, 2)

    Dim c As Long
    For c = firstCol To lastCol
        Dim header As String
        header = Nz(tbl(headerRow, c))
        If StrComp(Left$(header, 8), "MOVESET_", vbTextCompare) = 0 Then
            Dim suffix As String
            suffix = Mid$(header, 9)
            Dim key As String
            key = GameVersionKey(suffix)
            If Len(key) > 0 And StrComp(key, GAME_KEY_ALL, vbTextCompare) <> 0 Then
                If Not mMovesetColumns.Exists(key) Then
                    mMovesetColumns.Add key, c
                End If
            End If
        End If
    Next c
End Sub

Private Sub BuildTypeOptions()
    mTypeOptions = CollectMoveTypeOptions()
End Sub

Private Sub BuildAbilityLists()
    Set mAbilityLists = CreateObject("Scripting.Dictionary")
    mAbilityLists.CompareMode = vbTextCompare

    Dim tbl As Variant
    tbl = GlobalTables.GameversionsTable
    If IsEmpty(tbl) Then Exit Sub

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstCol As Long
    firstCol = LBound(tbl, 2)
    Dim lastCol As Long
    lastCol = UBound(tbl, 2)

    Dim c As Long
    For c = firstCol To lastCol
        Dim header As String
        header = Nz(tbl(headerRow, c))
        Dim listKey As String
        listKey = AbilityListKeyFromHeader(header)
        If Len(listKey) = 0 Then GoTo ContinueCol

        Dim values As Variant
        values = ExtractAbilityColumnValues(tbl, c, headerRow)
        If HasArrayValues(values) Then
            mAbilityLists(listKey) = values
        End If
ContinueCol:
    Next c
End Sub

Private Function AbilityListKeyFromHeader(ByVal header As String) As String
    If StrComp(header, "ABILITIES_ALL", vbTextCompare) = 0 Then
        AbilityListKeyFromHeader = GAME_KEY_ALL
    ElseIf Left$(header, 10) = "ABILITIES_" Then
        Dim suffix As String
        suffix = Mid$(header, 11)
        Dim key As String
        key = GameVersionKey(suffix)
        If Len(key) > 0 And StrComp(key, GAME_KEY_ALL, vbTextCompare) <> 0 Then
            AbilityListKeyFromHeader = key
        End If
    End If
End Function

Private Function ExtractAbilityColumnValues(ByRef tbl As Variant, ByVal columnIndex As Long, ByVal headerRow As Long) As Variant
    Dim firstRow As Long
    firstRow = headerRow + 1
    Dim lastRow As Long
    lastRow = UBound(tbl, 1)
    If firstRow > lastRow Then Exit Function

    Dim arr() As String
    Dim count As Long
    Dim r As Long
    For r = firstRow To lastRow
        Dim valueText As String
        valueText = Nz(tbl(r, columnIndex))
        If Len(valueText) > 0 Then
            count = count + 1
            If count = 1 Then
                ReDim arr(1 To 1)
            Else
                ReDim Preserve arr(1 To count)
            End If
            arr(count) = valueText
        End If
    Next r

    If count > 0 Then
        ExtractAbilityColumnValues = arr
    End If
End Function

Private Sub PopulateGameCombo(ByVal desiredSelection As String)
    Dim target As String
    target = ResolveGameLabel(desiredSelection)

    mCboGame.Clear
    mCboGame.AddItem FILTER_ALL

    If Not IsEmpty(mGameOptions) Then
        Dim i As Long
        For i = LBound(mGameOptions) To UBound(mGameOptions)
            Dim optionLabel As String
            optionLabel = CStr(mGameOptions(i))
            If Len(optionLabel) > 0 And StrComp(optionLabel, FILTER_ALL, vbTextCompare) <> 0 Then
                mCboGame.AddItem optionLabel
            End If
        Next i
    End If

    SetGameSelectionFromContext target
End Sub

Private Sub PopulateTypeCombo()
    mCboType.Clear
    mCboType.AddItem FILTER_ALL

    If IsArray(mTypeOptions) Then
        Dim i As Long
        For i = LBound(mTypeOptions) To UBound(mTypeOptions)
            If Len(CStr(mTypeOptions(i))) > 0 Then
                mCboType.AddItem CStr(mTypeOptions(i))
            End If
        Next i
    End If
End Sub

Private Function MovesetColumnForKey(ByVal versionKey As String) As Long
    If versionKey = GAME_KEY_ALL Then Exit Function
    If mMovesetColumns Is Nothing Then Exit Function
    If mMovesetColumns.Exists(versionKey) Then
        MovesetColumnForKey = CLng(mMovesetColumns(versionKey))
    End If
End Function

Private Function ResolveGameLabel(ByVal rawValue As String) As String
    Dim cleaned As String
    cleaned = CleanSelection(rawValue, FILTER_ALL)
    If StrComp(cleaned, FILTER_ALL, vbTextCompare) = 0 Then
        ResolveGameLabel = FILTER_ALL
        Exit Function
    End If

    If Not mGameKeyToLabel Is Nothing Then
        Dim key As String
        key = GameVersionKey(cleaned)
        If mGameKeyToLabel.Exists(key) Then
            ResolveGameLabel = CStr(mGameKeyToLabel(key))
            Exit Function
        End If
    End If

    ResolveGameLabel = cleaned
End Function

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

    If dict.count = 0 Then Exit Sub
    mGameOptions = DictionaryToSortedArray(dict)
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
        labelText = Nz(values(i))
        If Len(labelText) = 0 Then GoTo ContinueRow
        If Not dict.Exists(labelText) Then
            dict.Add labelText, True
            Dim key As String
            key = GameVersionKey(labelText)
            If Len(key) > 0 And StrComp(key, GAME_KEY_ALL, vbTextCompare) <> 0 Then
                If Not mGameKeyToLabel.Exists(key) Then
                    mGameKeyToLabel.Add key, labelText
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

Private Function Nz(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        Nz = ""
    Else
        Nz = Trim$(CStr(v))
    End If
End Function

Private Function SafeToLong(ByVal value As Variant) As Long
    On Error GoTo CleanZero
    Dim text As String
    text = Nz(value)
    If Len(text) = 0 Then
        SafeToLong = 0
    ElseIf IsNumeric(text) Then
        SafeToLong = CLng(CDbl(text))
    Else
        SafeToLong = CLng(val(text))
    End If
    Exit Function
CleanZero:
    SafeToLong = 0
End Function
