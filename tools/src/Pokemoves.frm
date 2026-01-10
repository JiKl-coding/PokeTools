VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pokemoves 
   Caption         =   "Pokemoves"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15390
   OleObjectBlob   =   "Pokemoves.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Pokemoves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Const FILTER_ALL As String = "All"
Private Const GAME_KEY_ALL As String = "__all__"
Private Const UI_FONT_NAME As String = "Aptos Narrow"
Private Const UI_FONT_SIZE As Integer = 12
Private Const PAD As Single = 6
Private Const HEADER_H As Single = 18
Private Const ROW_MIN_H As Single = 22

Private Enum GridCol
    gcPokemon = 0
    gcMethods = 1
End Enum

Private Enum PokemovesContext
    MoveCtxMovedex = 0
    MoveCtxOther = 1
End Enum

Private Type LearnerRow
    pokemonName As String
    formKey As String
    Methods As String
End Type

Private mColW(0 To 1) As Single
Private mRows() As LearnerRow
Private mRowCount As Long
Private mSortCol As Long
Private mSortAsc As Boolean

Private mLblInfo As MSForms.label
Private mLblHint As MSForms.label
Private WithEvents mCboGame As MSForms.ComboBox
Attribute mCboGame.VB_VarHelpID = -1
Private WithEvents mCboMove As MSForms.ComboBox
Attribute mCboMove.VB_VarHelpID = -1
Private WithEvents mBtnApply As MSForms.CommandButton
Attribute mBtnApply.VB_VarHelpID = -1
Private WithEvents mBtnClear As MSForms.CommandButton
Attribute mBtnClear.VB_VarHelpID = -1
Private WithEvents mBtnClose As MSForms.CommandButton
Attribute mBtnClose.VB_VarHelpID = -1
Private mFraHeader As MSForms.Frame
Private mFraGrid As MSForms.Frame
Private mMeasureLabel As MSForms.label

Private mHeaderEvents As Collection
Private mRowEvents As Collection
Private mFilterEvents As Collection

Private mInFilterUpdate As Boolean
Private mContextTarget As PokemovesContext
Private mLastGameSel As String
Private mLastMoveSel As String

Private mGameOptions As Variant
Private mGameLabelToSlug As Object
Private mGameSlugToLabel As Object
Private mMoveDisplayToKey As Object

' =============================
' Form lifecycle
' =============================
Private Sub UserForm_Initialize()
    On Error GoTo CleanFail

    mContextTarget = DetectFormContext()
    mInFilterUpdate = True
    Me.BackColor = RGB(204, 0, 0)
    On Error Resume Next
    Me.Font.name = UI_FONT_NAME
    Me.Font.Size = UI_FONT_SIZE
    On Error GoTo 0

    InitColumnWidths
    BuildRuntimeUI
    EnsureDataCaches

    PopulateGameCombo DefaultGameValue()
    PopulateMoveCombo CleanSelection(mCboGame.value, FILTER_ALL), DefaultMoveValue()

    mSortCol = gcPokemon
    mSortAsc = True

    SetInfoLabel

    mInFilterUpdate = False
    
    FiltersChanged
    Exit Sub

CleanFail:
    mInFilterUpdate = False
    MsgBox "Unable to initialize Pokemoves: " & Err.Description, vbExclamation
End Sub

Private Sub InitColumnWidths()
    mColW(gcPokemon) = 220
    mColW(gcMethods) = 500
End Sub

Private Sub BuildRuntimeUI()
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        ctrl.Visible = False
    Next ctrl

    Set mLblInfo = Me.Controls.Add("Forms.Label.1", "lblInfoPM", True)
    With mLblInfo
        .Left = PAD
        .Top = PAD
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = 26
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE + 5
        .Font.Bold = True
        .caption = "Pokemon Move Learners"
    End With

    Dim x As Single, y As Single
    y = mLblInfo.Top + mLblInfo.Height + PAD
    x = PAD

    Dim lblGame As MSForms.label
    Set lblGame = Me.Controls.Add("Forms.Label.1", "lblGamePM", True)
    With lblGame
        .Left = x
        .Top = y + 2
        .Width = 60
        .Height = 16
        .caption = "Game"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + lblGame.Width + 4

    Set mCboGame = Me.Controls.Add("Forms.ComboBox.1", "cboGamePM", True)
    With mCboGame
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
    x = mCboGame.Left + mCboGame.Width + 18

    Dim lblMove As MSForms.label
    Set lblMove = Me.Controls.Add("Forms.Label.1", "lblMovePM", True)
    With lblMove
        .Left = x
        .Top = y + 2
        .Width = 45
        .Height = 16
        .caption = "Move"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + lblMove.Width + 4

    Set mCboMove = Me.Controls.Add("Forms.ComboBox.1", "cboMovePM", True)
    With mCboMove
        .Left = x
        .Top = y
        .Width = 220
        .Height = 20
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .Font.name = UI_FONT_NAME
        On Error Resume Next
        .Font.Size = UI_FONT_SIZE
        On Error GoTo 0
    End With

    Set mFilterEvents = New Collection
    Dim evGame As CGridComboEvents
    Dim evMove As CGridComboEvents
    Set evGame = New CGridComboEvents
    evGame.Init Me, mCboGame
    mFilterEvents.Add evGame
    Set evMove = New CGridComboEvents
    evMove.Init Me, mCboMove
    mFilterEvents.Add evMove

    Dim btnY As Single
    btnY = y + 28
    x = PAD

    Set mBtnApply = Me.Controls.Add("Forms.CommandButton.1", "btnApplyPM", True)
    With mBtnApply
        .Left = x
        .Top = btnY
        .Width = 80
        .Height = 24
        .caption = "Apply"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + mBtnApply.Width + 8

    Set mBtnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClearPM", True)
    With mBtnClear
        .Left = x
        .Top = btnY
        .Width = 80
        .Height = 24
        .caption = "Clear"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    x = x + mBtnClear.Width + 8

    Set mBtnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClosePM", True)
    With mBtnClose
        .Left = x
        .Top = btnY
        .Width = 80
        .Height = 24
        .caption = "Close"
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With

    y = btnY + mBtnClose.Height + PAD
    Set mFraHeader = Me.Controls.Add("Forms.Frame.1", "fraHeaderPM", True)
    With mFraHeader
        .Left = PAD
        .Top = y
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = HEADER_H + 4
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BackColor = vbWhite
    End With

    y = mFraHeader.Top + mFraHeader.Height
    Set mFraGrid = Me.Controls.Add("Forms.Frame.1", "fraGridPM", True)
    With mFraGrid
        .Left = PAD
        .Top = y
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = Me.InsideHeight - y - 60
        .SpecialEffect = fmSpecialEffectSunken
        .BorderStyle = fmBorderStyleSingle
        .BackColor = vbWhite
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = .InsideHeight
    End With

    Set mLblHint = Me.Controls.Add("Forms.Label.1", "lblHintPM", True)
    With mLblHint
        .Left = PAD
        .Top = mFraGrid.Top + mFraGrid.Height + 6
        .Width = Me.InsideWidth - (PAD * 2)
        .Height = 18
        .caption = "Double-click to see the Pokemon in Pokedex"
        .ForeColor = vbWhite
        .BackStyle = fmBackStyleTransparent
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE - 1
        .Font.Italic = True
        .Visible = (mContextTarget = MoveCtxMovedex)
    End With

    Set mHeaderEvents = New Collection
    Set mRowEvents = New Collection
    BuildHeaderLabels
    EnsureMeasureLabel
End Sub

Private Sub BuildHeaderLabels()
    Set mHeaderEvents = New Collection
    Dim captions() As String
    ReDim captions(LBound(mColW) To UBound(mColW))
    captions(gcPokemon) = "Pokemon"
    captions(gcMethods) = "Methods"

    Dim x As Single
    x = 2

    Dim i As Long
    For i = LBound(mColW) To UBound(mColW)
        Dim lbl As MSForms.label
        Set lbl = mFraHeader.Controls.Add("Forms.Label.1", "hdrPM" & i, True)
        With lbl
            .Left = x
            .Top = 2
            .Width = mColW(i)
            .Height = HEADER_H
            .caption = captions(i)
            .Font.name = UI_FONT_NAME
            .Font.Size = UI_FONT_SIZE - 1
            .Font.Bold = True
            .BackStyle = fmBackStyleTransparent
        End With
        Dim ev As CGridHeaderLabel
        Set ev = New CGridHeaderLabel
        ev.Init Me, lbl, i
        mHeaderEvents.Add ev
        x = x + mColW(i)
    Next i
End Sub

' =============================
' Control events
' =============================
Private Sub mBtnApply_Click()
    FiltersChanged
End Sub

Private Sub mBtnClear_Click()
    On Error GoTo CleanFail
    mInFilterUpdate = True
    EnsureComboSelection mCboGame, FILTER_ALL
    PopulateMoveCombo FILTER_ALL, vbNullString
    mInFilterUpdate = False
    FiltersChanged
    Exit Sub
CleanFail:
    mInFilterUpdate = False
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub


' =============================
' Sorting hooks
' =============================
Public Sub HeaderClicked(ByVal colIndex As Long)
    If colIndex < LBound(mColW) Or colIndex > UBound(mColW) Then Exit Sub
    If mSortCol = colIndex Then
        mSortAsc = Not mSortAsc
    Else
        mSortCol = colIndex
        mSortAsc = True
    End If
    SortLearnerRows
    RenderGrid
End Sub

' =============================
' Filtering + data
' =============================
Private Sub FiltersChanged()
    If mInFilterUpdate Then Exit Sub
    mInFilterUpdate = True

    Dim gameLabel As String
    gameLabel = CleanSelection(mCboGame.value, FILTER_ALL)
    If Len(gameLabel) = 0 Then
        EnsureComboSelection mCboGame, FILTER_ALL
        gameLabel = FILTER_ALL
    End If

    Dim moveLabel As String
    moveLabel = CleanSelection(mCboMove.value, vbNullString)

    Dim gameChanged As Boolean
    gameChanged = (StrComp(gameLabel, mLastGameSel, vbTextCompare) <> 0)

    If gameChanged Then
        PopulateMoveCombo gameLabel, moveLabel
        moveLabel = CleanSelection(mCboMove.value, vbNullString)
    End If

    If Len(moveLabel) = 0 And mCboMove.ListCount > 0 Then
        mCboMove.ListIndex = 0
        moveLabel = CleanSelection(mCboMove.value, vbNullString)
    End If

    mLastGameSel = gameLabel
    mLastMoveSel = moveLabel

    SetInfoLabel
    LoadLearnerRows moveLabel, gameLabel
    RenderGrid

    mInFilterUpdate = False
End Sub

Public Sub ComboClicked(ByVal ctrlName As String)
    Dim target As MSForms.ComboBox
    Select Case LCase$(ctrlName)
        Case "cbogamepm": Set target = mCboGame
        Case "cbomovepm": Set target = mCboMove
        Case Else: Exit Sub
    End Select
    HighlightComboText target
End Sub

Public Sub ComboTyped(ByVal ctrlName As String, ByVal typed As String)
    ' Typed filtering not implemented for Pokemoves combos.
End Sub

Public Sub FilterControlChanged(ByVal ctrlName As String)
    ' Filters apply only when the user clicks the Apply button.
End Sub

Private Sub SetInfoLabel()
    Dim gameLabel As String
    gameLabel = CleanSelection(mCboGame.value, FILTER_ALL)
    Dim moveLabel As String
    moveLabel = CleanSelection(mCboMove.value, "(select move)")
    If Len(moveLabel) = 0 Then moveLabel = "(select move)"
    mLblInfo.caption = "Pokemon that learn " & moveLabel & " (" & gameLabel & ")"
End Sub

Private Sub LoadLearnerRows(ByVal moveDisplay As String, ByVal gameLabel As String)
    ReDim mRows(1 To 1)
    mRowCount = 0

    Dim moveKey As String
    moveKey = FindMoveKeyByDisplay(moveDisplay)
    If Len(moveKey) = 0 Then Exit Sub

    Dim gameSlug As String
    gameSlug = GameSlugForSelection(gameLabel)

    GlobalTables.LoadLearnsetsTable
    Dim tbl As Variant
    tbl = GlobalTables.LearnsetsTable
    If IsEmpty(tbl) Then Exit Sub

    Dim colMove As Long
    Dim colName As Long
    Dim colForm As Long
    Dim colMethod As Long
    Dim colLevel As Long
    Dim colVersion As Long

    colMove = GlobalTables.FindHeaderColumn(tbl, "MOVE_KEY")
    colName = GlobalTables.FindHeaderColumn(tbl, "DISPLAY_NAME")
    colForm = GlobalTables.FindHeaderColumn(tbl, "FORM_KEY")
    colMethod = GlobalTables.FindHeaderColumn(tbl, "METHOD")
    colLevel = GlobalTables.FindHeaderColumn(tbl, "LEVEL")
    colVersion = GlobalTables.FindHeaderColumn(tbl, "VERSION_GROUP")

    If colMove = 0 Or colName = 0 Or colVersion = 0 Then Exit Sub

    Dim learners As Object
    Set learners = CreateObject("Scripting.Dictionary")
    learners.CompareMode = vbTextCompare

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstRow As Long
    firstRow = headerRow + 1
    Dim lastRow As Long
    lastRow = UBound(tbl, 1)

    Dim r As Long
    For r = firstRow To lastRow
        If StrComp(Nz(tbl(r, colMove)), moveKey, vbTextCompare) <> 0 Then GoTo ContinueLoop

        Dim versionKey As String
        versionKey = CleanSelection(tbl(r, colVersion), vbNullString)
        If gameSlug <> GAME_KEY_ALL Then
            If StrComp(versionKey, gameSlug, vbTextCompare) <> 0 Then GoTo ContinueLoop
        End If

        Dim pokemonName As String
        pokemonName = CleanSelection(tbl(r, colName), vbNullString)
        Dim formKey As String
        If colForm > 0 Then formKey = CleanSelection(tbl(r, colForm), vbNullString)
        If Len(pokemonName) = 0 And Len(formKey) = 0 Then GoTo ContinueLoop

        Dim rowKey As String
        rowKey = LCase$(pokemonName & "|" & formKey)
        If Len(rowKey) = 0 Then rowKey = LCase$(formKey)
        If Len(rowKey) = 0 Then rowKey = CStr(r)

        Dim bucket As Object
        If learners.Exists(rowKey) Then
            Set bucket = learners(rowKey)
        Else
            Set bucket = CreateObject("Scripting.Dictionary")
            bucket.CompareMode = vbTextCompare
            bucket.Add "pokemon", pokemonName
            bucket.Add "form", formKey
            bucket.Add "methods", CreateObject("Scripting.Dictionary")
            bucket("methods").CompareMode = vbTextCompare
            learners.Add rowKey, bucket
        End If

        Dim methodText As String
        methodText = MethodDescriptor(tbl(r, colMethod), tbl(r, colLevel))
        If Len(methodText) = 0 Then methodText = "Unknown"

        Dim methodsDict As Object
        Set methodsDict = bucket("methods")
        If Not methodsDict.Exists(methodText) Then methodsDict.Add methodText, True
ContinueLoop:
    Next r

    If learners.count = 0 Then Exit Sub

    ReDim mRows(1 To learners.count)
    Dim idx As Long
    Dim key As Variant
    For Each key In learners.keys
        idx = idx + 1
        Dim rec As Object
        Set rec = learners(key)
        mRows(idx).pokemonName = Nz(rec("pokemon"))
        mRows(idx).formKey = Nz(rec("form"))
        mRows(idx).Methods = JoinDictionaryKeys(rec("methods"))
    Next key
    mRowCount = idx
    SortLearnerRows
End Sub

' =============================
' Combo population
' =============================
Private Sub PopulateGameCombo(ByVal desiredSelection As String)
    BuildGameOptions

    Dim target As String
    target = ResolveGameLabel(desiredSelection)

    mCboGame.Clear
    If Not IsEmpty(mGameOptions) Then
        Dim i As Long
        For i = LBound(mGameOptions) To UBound(mGameOptions)
            If Len(mGameOptions(i)) > 0 Then mCboGame.AddItem mGameOptions(i)
        Next i
    End If

    If mCboGame.ListCount = 0 Then
        mCboGame.AddItem FILTER_ALL
    End If

    EnsureComboSelection mCboGame, target
End Sub

Private Sub PopulateMoveCombo(ByVal gameLabel As String, ByVal desiredMove As String)
    Dim slug As String
    slug = GameSlugForSelection(gameLabel)

    Dim options As Variant
    options = GetMoveOptionsForGame(slug)

    mCboMove.Clear
    If Not IsEmpty(options) Then
        Dim i As Long
        For i = LBound(options) To UBound(options)
            If Len(options(i)) > 0 Then mCboMove.AddItem options(i)
        Next i
    End If

    Dim fallback As String
    fallback = CleanSelection(desiredMove, vbNullString)
    If Len(fallback) = 0 And mCboMove.ListCount > 0 Then
        mCboMove.ListIndex = 0
    Else
        EnsureComboSelection mCboMove, fallback, vbNullString
    End If
End Sub

Private Function GetMoveOptionsForGame(ByVal gameSlug As String) As Variant
    GlobalTables.LoadGameversionsTable
    Dim tbl As Variant
    tbl = GlobalTables.GameversionsTable
    If IsEmpty(tbl) Then Exit Function

    Dim headerName As String
    If StrComp(gameSlug, GAME_KEY_ALL, vbTextCompare) = 0 Then
        headerName = "MOVES_ALL"
    Else
        headerName = "MOVES_" & gameSlug
    End If

    Dim colIndex As Long
    colIndex = GlobalTables.FindHeaderColumn(tbl, headerName)
    If colIndex = 0 And StrComp(headerName, "MOVES_ALL", vbTextCompare) <> 0 Then
        colIndex = GlobalTables.FindHeaderColumn(tbl, "MOVES_ALL")
    End If
    If colIndex = 0 Then Exit Function

    Dim columnValues As Variant
    columnValues = GlobalTables.ExtractColumnValues(tbl, colIndex, True)
    If IsEmpty(columnValues) Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(columnValues) To UBound(columnValues)
        Dim textValue As String
        textValue = CleanSelection(columnValues(i), vbNullString)
        If Len(textValue) > 0 And textValue <> "0" Then
            If Not dict.Exists(textValue) Then dict.Add textValue, True
        End If
    Next i

    GetMoveOptionsForGame = DictionaryToSortedArray(dict)
End Function

' =============================
' Context helpers
' =============================
Private Function DefaultGameValue() As String
    Select Case mContextTarget
        Case MoveCtxMovedex
            On Error GoTo CleanFallback
            DefaultGameValue = CleanSelection(Movedex.Range("GAMEVERSION").value, FILTER_ALL)
            Exit Function
    End Select
CleanFallback:
    DefaultGameValue = FILTER_ALL
End Function

Private Function DefaultMoveValue() As String
    If mContextTarget = MoveCtxMovedex Then
        On Error GoTo CleanFallback
        DefaultMoveValue = CleanSelection(Movedex.Range("MVLIST").value, vbNullString)
        Exit Function
    End If
CleanFallback:
    DefaultMoveValue = vbNullString
End Function

Private Function DetectFormContext() As PokemovesContext
    On Error GoTo CleanFallback
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet
    If ws Is Nothing Then GoTo CleanFallback

    If ws Is Movedex Then
        DetectFormContext = MoveCtxMovedex
    Else
        DetectFormContext = MoveCtxOther
    End If
    Exit Function

CleanFallback:
    DetectFormContext = MoveCtxOther
End Function

Private Function GameSlugForSelection(ByVal selectionText As String) As String
    BuildGameOptions
    Dim cleaned As String
    cleaned = CleanSelection(selectionText, FILTER_ALL)
    If StrComp(cleaned, FILTER_ALL, vbTextCompare) = 0 Then
        GameSlugForSelection = GAME_KEY_ALL
        Exit Function
    End If

    If Not mGameLabelToSlug Is Nothing Then
        If mGameLabelToSlug.Exists(cleaned) Then
            GameSlugForSelection = mGameLabelToSlug(cleaned)
            Exit Function
        End If
    End If

    GameSlugForSelection = GameVersionKey(cleaned, FILTER_ALL, GAME_KEY_ALL)
End Function

Private Sub BuildGameOptions()
    If Not IsEmpty(mGameOptions) Then Exit Sub

    GlobalTables.LoadAssetsTable
    GlobalTables.LoadGameversionsTable

    mGameOptions = Empty
    Set mGameLabelToSlug = Nothing
    Set mGameSlugToLabel = Nothing

    If BuildGameOptionsFromAssetsTable() Then Exit Sub
    If BuildGameOptionsFromGameversions() Then Exit Sub
End Sub

Private Function BuildGameOptionsFromAssetsTable() As Boolean
    Dim tbl As Variant
    tbl = GlobalTables.AssetsTable
    If IsEmpty(tbl) Then Exit Function

    Dim colIndex As Long
    colIndex = GlobalTables.FindHeaderColumn(tbl, "GAMES")
    If colIndex = 0 Then Exit Function

    Dim values As Variant
    values = GlobalTables.ExtractColumnValues(tbl, colIndex, True)
    If IsEmpty(values) Then Exit Function

    Dim labelsDict As Object
    Set labelsDict = CreateObject("Scripting.Dictionary")
    labelsDict.CompareMode = vbTextCompare

    Set mGameLabelToSlug = CreateObject("Scripting.Dictionary")
    mGameLabelToSlug.CompareMode = vbTextCompare
    Set mGameSlugToLabel = CreateObject("Scripting.Dictionary")
    mGameSlugToLabel.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(values) To UBound(values)
        Dim labelText As String
        labelText = CleanSelection(values(i), vbNullString)
        If Len(labelText) = 0 Then GoTo ContinueLoop

        Dim slug As String
        slug = GameVersionKey(labelText, FILTER_ALL, GAME_KEY_ALL)

        If Not labelsDict.Exists(labelText) Then labelsDict.Add labelText, True
        If Not mGameLabelToSlug.Exists(labelText) Then mGameLabelToSlug.Add labelText, slug

        If Len(slug) > 0 Then
            If Not mGameLabelToSlug.Exists(slug) Then mGameLabelToSlug.Add slug, slug
            If StrComp(slug, GAME_KEY_ALL, vbTextCompare) <> 0 Then
                If Not mGameSlugToLabel.Exists(slug) Then mGameSlugToLabel.Add slug, labelText
            End If
        End If
ContinueLoop:
    Next i

    If labelsDict.count = 0 Then Exit Function

    mGameOptions = DictionaryToSortedArray(labelsDict)
    BuildGameOptionsFromAssetsTable = True
End Function

Private Function BuildGameOptionsFromGameversions() As Boolean
    Dim tbl As Variant
    tbl = GlobalTables.GameversionsTable
    If IsEmpty(tbl) Then Exit Function

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstCol As Long
    firstCol = LBound(tbl, 2)
    Dim lastCol As Long
    lastCol = UBound(tbl, 2)

    Dim labelsDict As Object
    Set labelsDict = CreateObject("Scripting.Dictionary")
    labelsDict.CompareMode = vbTextCompare

    Set mGameLabelToSlug = CreateObject("Scripting.Dictionary")
    mGameLabelToSlug.CompareMode = vbTextCompare
    Set mGameSlugToLabel = CreateObject("Scripting.Dictionary")
    mGameSlugToLabel.CompareMode = vbTextCompare

    Dim c As Long
    For c = firstCol To lastCol
        Dim headerText As String
        headerText = CleanSelection(tbl(headerRow, c), vbNullString)
        If Left$(headerText, Len("MOVES_")) = "MOVES_" Then
            Dim rawSlug As String
            rawSlug = Mid$(headerText, Len("MOVES_") + 1)
            Dim slug As String
            slug = GameVersionKey(rawSlug, FILTER_ALL, GAME_KEY_ALL)
            If StrComp(slug, GAME_KEY_ALL, vbTextCompare) = 0 Then GoTo ContinueCol

            Dim labelText As String
            labelText = CleanSelection(DexLogic.NormalizeGameVersion(rawSlug), slug)
            If Len(labelText) = 0 Then labelText = slug

            If Not labelsDict.Exists(labelText) Then labelsDict.Add labelText, True
            If Not mGameLabelToSlug.Exists(labelText) Then mGameLabelToSlug.Add labelText, slug
            If Not mGameLabelToSlug.Exists(slug) Then mGameLabelToSlug.Add slug, slug
            If Not mGameSlugToLabel.Exists(slug) Then mGameSlugToLabel.Add slug, labelText
        End If
ContinueCol:
    Next c

    If labelsDict.count = 0 Then Exit Function

    mGameOptions = DictionaryToSortedArray(labelsDict)
    BuildGameOptionsFromGameversions = True
End Function

Private Function ResolveGameLabel(ByVal rawValue As String) As String
    BuildGameOptions

    Dim cleaned As String
    cleaned = CleanSelection(rawValue, FILTER_ALL)
    If Len(cleaned) = 0 Or StrComp(cleaned, FILTER_ALL, vbTextCompare) = 0 Then
        ResolveGameLabel = FILTER_ALL
        Exit Function
    End If

    Dim slug As String
    slug = GameVersionKey(cleaned, FILTER_ALL, GAME_KEY_ALL)
    If StrComp(slug, GAME_KEY_ALL, vbTextCompare) = 0 Then
        ResolveGameLabel = FILTER_ALL
        Exit Function
    End If

    Dim label As String
    label = GameLabelForSlug(slug)
    If Len(label) > 0 Then
        ResolveGameLabel = label
    Else
        ResolveGameLabel = cleaned
    End If
End Function

Private Function GameLabelForSlug(ByVal slugValue As String) As String
    If Len(slugValue) = 0 Then Exit Function

    Dim normalized As String
    normalized = GameVersionKey(slugValue, FILTER_ALL, GAME_KEY_ALL)
    If StrComp(normalized, GAME_KEY_ALL, vbTextCompare) = 0 Then
        GameLabelForSlug = FILTER_ALL
        Exit Function
    End If

    If mGameSlugToLabel Is Nothing Then Exit Function
    If mGameSlugToLabel.Exists(normalized) Then
        GameLabelForSlug = mGameSlugToLabel(normalized)
    End If
End Function

' =============================
' Data caches
' =============================
Private Sub EnsureDataCaches()
    If Not mMoveDisplayToKey Is Nothing Then Exit Sub

    GlobalTables.LoadMovesTable
    Dim tbl As Variant
    tbl = GlobalTables.movesTable
    If IsEmpty(tbl) Then
        Set mMoveDisplayToKey = CreateObject("Scripting.Dictionary")
        mMoveDisplayToKey.CompareMode = vbTextCompare
        Exit Sub
    End If

    Dim colDisplay As Long
    Dim colKey As Long
    colDisplay = GlobalTables.FindHeaderColumn(tbl, "DISPLAY_NAME")
    colKey = GlobalTables.FindHeaderColumn(tbl, "MOVE_KEY")
    Set mMoveDisplayToKey = CreateObject("Scripting.Dictionary")
    mMoveDisplayToKey.CompareMode = vbTextCompare

    If colDisplay = 0 Or colKey = 0 Then Exit Sub

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstRow As Long
    firstRow = headerRow + 1
    Dim lastRow As Long
    lastRow = UBound(tbl, 1)

    Dim r As Long
    For r = firstRow To lastRow
        Dim displayName As String
        displayName = CleanSelection(tbl(r, colDisplay), vbNullString)
        Dim moveKey As String
        moveKey = CleanSelection(tbl(r, colKey), vbNullString)
        If Len(displayName) = 0 Or Len(moveKey) = 0 Then GoTo ContinueLoop
        AddMoveKeyMapping displayName, moveKey
        Dim normalized As String
        normalized = NormalizeMoveNameKey(displayName)
        If Len(normalized) > 0 Then AddMoveKeyMapping normalized, moveKey
ContinueLoop:
    Next r
End Sub

Private Sub AddMoveKeyMapping(ByVal keyName As String, ByVal moveKey As String)
    If Len(keyName) = 0 Or Len(moveKey) = 0 Then Exit Sub
    If mMoveDisplayToKey.Exists(keyName) Then Exit Sub
    mMoveDisplayToKey.Add keyName, moveKey
End Sub

Private Function FindMoveKeyByDisplay(ByVal moveDisplay As String) As String
    EnsureDataCaches
    Dim cleaned As String
    cleaned = CleanSelection(moveDisplay, vbNullString)
    If Len(cleaned) = 0 Then Exit Function

    If mMoveDisplayToKey.Exists(cleaned) Then
        FindMoveKeyByDisplay = mMoveDisplayToKey(cleaned)
        Exit Function
    End If

    Dim normalized As String
    normalized = NormalizeMoveNameKey(cleaned)
    If Len(normalized) = 0 Then Exit Function
    If mMoveDisplayToKey.Exists(normalized) Then
        FindMoveKeyByDisplay = mMoveDisplayToKey(normalized)
    End If
End Function

' =============================
' Rendering
' =============================
Private Sub RenderGrid()
    If mFraGrid Is Nothing Then Exit Sub

    ClearGridRows

    Dim contentHeight As Single
    contentHeight = 0

    If mRowCount = 0 Then
        Dim lblEmpty As MSForms.label
        Set lblEmpty = mFraGrid.Controls.Add("Forms.Label.1", "lblEmptyPM", True)
        With lblEmpty
            .Left = 4
            .Top = 4
            .Width = mFraGrid.InsideWidth - 8
            .Height = 20
            .caption = "No Pokemon found for the current filters."
            .Font.name = UI_FONT_NAME
            .Font.Size = UI_FONT_SIZE
        End With
        mFraGrid.ScrollHeight = mFraGrid.InsideHeight
        Exit Sub
    End If

    Dim rowIndex As Long
    Dim topY As Single
    topY = 4

    For rowIndex = 1 To mRowCount
        Dim rowH As Single
        rowH = CalcRowHeight(mRows(rowIndex).Methods)
        AddGridRow rowIndex, topY, rowH
        topY = topY + rowH + 2
    Next rowIndex

    contentHeight = topY + PAD
    mFraGrid.ScrollHeight = Application.WorksheetFunction.Max(contentHeight, mFraGrid.InsideHeight)
End Sub

Private Sub ClearGridRows()
    On Error Resume Next
    Dim ctrl As MSForms.Control
    Dim namesToRemove As Collection
    Set namesToRemove = New Collection

    For Each ctrl In mFraGrid.Controls
        If Left$(ctrl.name, 3) = "r__" Or ctrl.name = "lblEmptyPM" Then
            namesToRemove.Add ctrl.name
        End If
    Next ctrl

    Dim entry As Variant
    For Each entry In namesToRemove
        mFraGrid.Controls.Remove CStr(entry)
    Next entry
    On Error GoTo 0

    Set mRowEvents = New Collection
End Sub

Private Sub AddGridRow(ByVal rowIndex As Long, ByVal topY As Single, ByVal rowH As Single)
    Dim row As LearnerRow
    row = mRows(rowIndex)

    Dim x As Single
    x = 2

    AddCellLabel "r__nm" & rowIndex, row.pokemonName, x, topY, mColW(gcPokemon), rowH, False, False
    AttachRowEvent "r__nm" & rowIndex, rowIndex
    x = x + mColW(gcPokemon)

    AddCellLabel "r__mt" & rowIndex, row.Methods, x, topY, mColW(gcMethods), rowH, False, True
    AttachRowEvent "r__mt" & rowIndex, rowIndex
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
    If mContextTarget <> MoveCtxMovedex Then Exit Sub

    Dim pokemonName As String
    pokemonName = mRows(rowIndex).pokemonName
    If Len(pokemonName) = 0 Then Exit Sub

    Dim gameLabel As String
    gameLabel = CleanSelection(mCboGame.value, FILTER_ALL)
    Dim moveLabel As String
    moveLabel = CleanSelection(mCboMove.value, vbNullString)
    If Len(moveLabel) = 0 Then Exit Sub

    On Error Resume Next
    Pokedex.Activate
    Pokedex.Range("GAME").value = gameLabel
    Pokedex.Range("PKMN_DEX").value = pokemonName
    Pokedex.Range("PKMN_MOVELIST").value = moveLabel
    On Error GoTo 0

    Unload Me
End Sub

Private Function CalcRowHeight(ByVal methodsText As String) As Single
    Dim columnWidthPts As Single
    columnWidthPts = mColW(gcMethods)

    Dim lineCount As Long
    lineCount = EstimateLineCount(methodsText, columnWidthPts)
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
        segmentText = NormalizeSegment(segments(idx))
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

    Dim line As String
    Dim lines As Long
    lines = 1

    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        Dim token As String
        token = tokens(i)
        If (Len(line) = 0 And Len(token) > 0) Then
            line = token
        Else
            Dim candidate As String
            candidate = line & " " & token
            If MeasureTextWidth(candidate) > widthLimit Then
                lines = lines + LinesNeededForLongWord(token, widthLimit)
                line = token
            Else
                line = candidate
            End If
        End If
    Next i

    EstimateLinesForSegment = lines
End Function

Private Function LinesNeededForLongWord(ByVal word As String, ByVal widthLimit As Single) As Long
    If Len(word) = 0 Then
        LinesNeededForLongWord = 1
        Exit Function
    End If

    Dim wordWidth As Single
    wordWidth = MeasureTextWidth(word)
    If wordWidth <= widthLimit Then
        LinesNeededForLongWord = 1
    Else
        Dim ratio As Double
        ratio = wordWidth / widthLimit
        LinesNeededForLongWord = CeilingPositive(ratio)
    End If
End Function

Private Function NormalizeSegment(ByVal textValue As String) As String
    Dim cleaned As String
    cleaned = Trim$(textValue)
    Do While InStr(cleaned, "  ") > 0
        cleaned = Replace(cleaned, "  ", " ")
    Loop
    NormalizeSegment = cleaned
End Function

Private Function CeilingPositive(ByVal value As Double) As Long
    CeilingPositive = CLng(Int(value))
    If value > Int(value) Then CeilingPositive = CeilingPositive + 1
    If CeilingPositive < 1 Then CeilingPositive = 1
End Function

Private Sub EnsureMeasureLabel()
    If Not mMeasureLabel Is Nothing Then Exit Sub
    Set mMeasureLabel = Me.Controls.Add("Forms.Label.1", "lblMeasurePM", True)
    With mMeasureLabel
        .Visible = False
        .AutoSize = True
        .Font.name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
End Sub

Private Function MeasureTextWidth(ByVal textValue As String) As Single
    If mMeasureLabel Is Nothing Then EnsureMeasureLabel
    mMeasureLabel.caption = textValue
    MeasureTextWidth = mMeasureLabel.Width
End Function

' =============================
' Sorting helpers
' =============================
Private Sub SortLearnerRows()
    If mRowCount <= 1 Then Exit Sub
    QuickSortRows 1, mRowCount
End Sub

Private Sub QuickSortRows(ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long
    i = lo
    j = hi
    Dim pivot As LearnerRow
    pivot = mRows((lo + hi) \ 2)
    Do While i <= j
        Do While CompareRows(mRows(i), pivot) < 0
            i = i + 1
        Loop
        Do While CompareRows(mRows(j), pivot) > 0
            j = j - 1
        Loop
        If i <= j Then
            Dim tmp As LearnerRow
            tmp = mRows(i)
            mRows(i) = mRows(j)
            mRows(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then QuickSortRows lo, j
    If i < hi Then QuickSortRows i, hi
End Sub

Private Function CompareRows(ByRef a As LearnerRow, ByRef b As LearnerRow) As Long
    Dim result As Long
    Select Case mSortCol
        Case gcMethods
            result = StrComp(a.Methods, b.Methods, vbTextCompare)
        Case Else
            result = StrComp(a.pokemonName, b.pokemonName, vbTextCompare)
    End Select

    If mSortAsc Then
        CompareRows = result
    Else
        CompareRows = -result
    End If
End Function

' =============================
' Misc helpers
' =============================
Private Function MethodDescriptor(ByVal methodValue As Variant, ByVal levelValue As Variant) As String
    Dim methodText As String
    methodText = CleanSelection(methodValue, vbNullString)
    Dim levelText As String
    levelText = CleanSelection(levelValue, vbNullString)

    If Len(methodText) = 0 And Len(levelText) = 0 Then
        MethodDescriptor = vbNullString
        Exit Function
    End If

    If Len(methodText) = 0 Then
        methodText = "Level"
    End If

    If Len(levelText) > 0 Then
        MethodDescriptor = methodText & " (Lv " & levelText & ")"
    Else
        MethodDescriptor = methodText
    End If
End Function

Private Function JoinDictionaryKeys(ByVal dict As Object) As String
    If dict Is Nothing Then Exit Function
    If dict.count = 0 Then Exit Function

    Dim arr() As String
    ReDim arr(1 To dict.count)

    Dim idx As Long
    Dim key As Variant
    For Each key In dict.keys
        idx = idx + 1
        arr(idx) = CStr(key)
    Next key

    SortStringArray arr
    JoinDictionaryKeys = Join(arr, ", ")
End Function

Private Sub SortStringArray(ByRef arr() As String)
    On Error GoTo CleanExit
    If Not IsArray(arr) Then Exit Sub
    QuickSortStrings arr, LBound(arr), UBound(arr)
CleanExit:
End Sub

Private Sub QuickSortStrings(ByRef arr() As String, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long
    Dim pivot As String
    pivot = arr((lo + hi) \ 2)
    i = lo
    j = hi
    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            Dim tmp As String
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then QuickSortStrings arr, lo, j
    If i < hi Then QuickSortStrings arr, i, hi
End Sub

Private Function SafeText(ByVal valueVariant As Variant) As String
    If IsError(valueVariant) Or IsNull(valueVariant) Or IsEmpty(valueVariant) Then Exit Function
    SafeText = Trim$(CStr(valueVariant))
End Function

Private Function Nz(ByVal valueVariant As Variant, Optional ByVal fallback As String = vbNullString) As String
    If IsError(valueVariant) Or IsNull(valueVariant) Or IsEmpty(valueVariant) Then
        Nz = fallback
    Else
        Nz = CStr(valueVariant)
    End If
End Function

Private Function NormalizeMoveNameKey(ByVal textValue As Variant) As String
    Dim t As String
    t = LCase$(Trim$(CStr(textValue)))
    If Len(t) = 0 Then Exit Function

    t = Replace(t, ChrW(&H2019), "'")
    t = Replace(t, "'", vbNullString)
    t = Replace(t, "-", " ")

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    NormalizeMoveNameKey = Trim$(t)
End Function

Private Function UBoundSafe(ByVal arr As Variant) As Long
    On Error GoTo CleanFail
    If IsEmpty(arr) Then Exit Function
    If Not IsArray(arr) Then Exit Function
    UBoundSafe = UBound(arr)
    Exit Function
CleanFail:
    UBoundSafe = 0
End Function

Private Function SafeArrayValue(ByVal arr As Variant, ByVal index As Long) As String
    On Error GoTo CleanFail
    If IsEmpty(arr) Then Exit Function
    If Not IsArray(arr) Then Exit Function
    SafeArrayValue = CleanSelection(arr(index), vbNullString)
    Exit Function
CleanFail:
    SafeArrayValue = vbNullString
End Function

