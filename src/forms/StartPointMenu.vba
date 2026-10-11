Option Explicit

' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String

Private Const ERR_START_FAILED As Long = vbObjectError + 5121
Private Const ERR_START_VALIDATION_FAILED As Long = vbObjectError + 5122
Private Const ERR_DIRECTION_FAILED As Long = vbObjectError + 5123

Private mFirstErrorOperation As String
Private mFirstErrorNumber As Long
Private mFirstErrorSource As String
Private mFirstErrorDescription As String
Private mFirstCandidateError As String

Private Type OrganizerStats
    ShapesSeen As Long
    ShapesConverted As Long
    NonCurveSkipped As Long
    SubpathsSeen As Long
    OpenSubpathsSkipped As Long
    DirectionReversed As Long
    StartAlreadyCorrect As Long
    StartChanged As Long
    RelativeHandleStarts As Long
    LineMidpointStarts As Long
    ReverseFallbackStarts As Long
    NoCandidateFound As Long
    CandidateReadErrors As Long
    IncompleteSearches As Long
    DirectionRollbacks As Long
    StartRollbacks As Long
    ReverseFallbackRollbacks As Long
    MidpointRollbacks As Long
    ValidationWarnings As Long
    Failures As Long
End Type

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOrganize_Click()
    Dim doc As Document
    Dim selectedShapes As ShapeRange
    Dim shp As Shape
    Dim stats As OrganizerStats
    Dim targetClockwise As Boolean
    Dim commandGroupOpen As Boolean
    Dim sourceUnit As cdrUnit
    Dim fatalNumber As Long
    Dim fatalSource As String
    Dim fatalDescription As String
    Dim closeNumber As Long
    Dim closeSource As String
    Dim closeDescription As String
    Dim messageText As String
    Dim messageStyle As VbMsgBoxStyle

    ClearFirstError

    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo 0

    If doc Is Nothing Then
        ShowMessageAfterClose "Tidak ada dokumen aktif.", vbExclamation
        Exit Sub
    End If

    If Not optClockwise.Value And Not optCounterClockwise.Value Then
        ShowMessageAfterClose "Pilih arah path terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set selectedShapes = ActiveSelectionRange
    On Error GoTo 0

    If selectedShapes Is Nothing Then
        ShowMessageAfterClose "Pilih objek cutting terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    If selectedShapes.Count = 0 Then
        ShowMessageAfterClose "Pilih objek cutting terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    targetClockwise = optClockwise.Value

    On Error GoTo FatalFail
    sourceUnit = doc.Unit
    doc.BeginCommandGroup "Start Point Organizer"
    commandGroupOpen = True

    For Each shp In selectedShapes
        ProcessShape shp, targetClockwise, sourceUnit, stats
    Next shp
    GoTo CleanExit

FatalFail:
    fatalNumber = Err.Number
    fatalSource = Err.Source
    fatalDescription = Err.Description
    SaveFirstError "cmdOrganize / eksekusi"
    stats.Failures = stats.Failures + 1
    Resume CleanExit

CleanExit:
    On Error GoTo 0
    If Not TryEndOrganizerCommandGroup(doc, commandGroupOpen, closeNumber, closeSource, closeDescription) Then
        messageText = "Penutupan CommandGroup gagal." & vbCrLf & _
                      "Err.Number: " & closeNumber & vbCrLf & _
                      "Err.Source: " & closeSource & vbCrLf & _
                      "Err.Description: " & closeDescription
    End If

    If fatalNumber <> 0 Then
        If Len(messageText) > 0 Then messageText = vbCrLf & vbCrLf & messageText
        messageText = "Macro berhenti karena error." & vbCrLf & _
                      "Err.Number: " & fatalNumber & vbCrLf & _
                      "Err.Source: " & fatalSource & vbCrLf & _
                      "Err.Description: " & fatalDescription & messageText
    End If

    messageStyle = vbInformation
    If fatalNumber <> 0 Or closeNumber <> 0 Then
        messageStyle = vbCritical
    ElseIf stats.Failures > 0 Or stats.ValidationWarnings > 0 Or _
           stats.CandidateReadErrors > 0 Or stats.NoCandidateFound > 0 Then
        messageText = "Selesai dengan catatan."
        messageStyle = vbExclamation
    Else
        messageText = "Selesai."
    End If
    ShowMessageAfterClose messageText & vbCrLf & vbCrLf & BuildSummary(stats), messageStyle
End Sub

Private Function TryEndOrganizerCommandGroup(ByVal doc As Document, ByRef groupOpen As Boolean, _
                                             ByRef errorNumber As Long, ByRef errorSource As String, _
                                             ByRef errorDescription As String) As Boolean
    On Error GoTo CloseFailed
    If groupOpen Then
        doc.EndCommandGroup
        groupOpen = False
    End If
    TryEndOrganizerCommandGroup = True
    Exit Function
CloseFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
End Function

' Kontrol ini sengaja menumpang sementara di SPO
Private Sub cmdZOrder_Click()
    Dim doc As Document
    Dim selectedShapes As ShapeRange
    Dim resultText As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ZOrderFailed

    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ZOrderFailed
    If doc Is Nothing Then
        ShowMessageAfterClose "Tidak ada dokumen aktif.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set selectedShapes = ActiveSelectionRange
    On Error GoTo ZOrderFailed
    If selectedShapes Is Nothing Then
        ShowMessageAfterClose "Pilih group atau curve terlebih dahulu.", vbExclamation
        Exit Sub
    End If
    If selectedShapes.Count = 0 Then
        ShowMessageAfterClose "Pilih group atau curve terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    resultText = SPOZOrderSelected(doc, selectedShapes)
    ShowMessageAfterClose resultText, vbInformation
    Exit Sub

ZOrderFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    ShowMessageAfterClose "ZOrder gagal." & vbCrLf & _
                          "Err.Number: " & errorNumber & vbCrLf & _
                          "Err.Description: " & errorDescription, vbCritical

End Sub

Private Sub optClockwise_Click()
    If optClockwise.Value Then optCounterClockwise.Value = False
End Sub

Private Sub optCounterClockwise_Click()
    If optCounterClockwise.Value Then optClockwise.Value = False
End Sub

Private Sub UserForm_Initialize()
    If Not optClockwise.Value And Not optCounterClockwise.Value Then
        optClockwise.Value = True
    End If
End Sub

Private Sub ProcessShape(ByVal shp As Shape, ByVal targetClockwise As Boolean, _
                         ByVal sourceUnit As cdrUnit, _
                         ByRef stats As OrganizerStats)
    Dim childShape As Shape
    Dim crv As Curve
    Dim subpathCount As Long
    Dim subpathIndex As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ShapeFail

    stats.ShapesSeen = stats.ShapesSeen + 1

    If shp.Type = cdrGroupShape Then
        For Each childShape In shp.Shapes
            ProcessShape childShape, targetClockwise, sourceUnit, stats
        Next childShape
        Exit Sub
    End If

    If shp.Type <> cdrCurveShape Then
        shp.ConvertToCurves
        stats.ShapesConverted = stats.ShapesConverted + 1
    End If

    If shp.Type <> cdrCurveShape Then
        stats.NonCurveSkipped = stats.NonCurveSkipped + 1
        Exit Sub
    End If

    Set crv = shp.Curve
    If crv Is Nothing Then
        stats.NonCurveSkipped = stats.NonCurveSkipped + 1
        Exit Sub
    End If

    subpathCount = crv.SubPaths.Count
    For subpathIndex = 1 To subpathCount
        Set crv = shp.Curve
        ProcessSubPath crv, subpathIndex, targetClockwise, sourceUnit, stats
    Next subpathIndex

    Exit Sub

ShapeFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    SaveFirstError "ProcessShape / ConvertToCurves / Curve access"
    If SPOIsFatalError(errorNumber) Then Err.Raise errorNumber, errorSource, errorDescription
    stats.Failures = stats.Failures + 1
    Err.Clear
End Sub

Private Sub ProcessSubPath(ByVal crv As Curve, ByVal subpathIndex As Long, _
                           ByVal targetClockwise As Boolean, _
                           ByVal sourceUnit As cdrUnit, _
                           ByRef stats As OrganizerStats)
    Dim sp As SubPath
    Dim candidate As SPOStartCandidate
    Dim candidateIndex As Long
    Dim diagnostics As SPOCandidateDiagnostics
    Dim diagnosticsRecorded As Boolean
    Dim found As Boolean
    Dim mutationBackup As Curve
    Dim mutationPending As Boolean
    Dim mutationOperation As String
    Dim insertsMidpoint As Boolean
    Dim startValid As Boolean
    Dim startClockwise As Boolean
    Dim targetAngle As Double
    Dim usingReverseFallback As Boolean
    Dim fallbackNoCandidate As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim restoreNumber As Long
    Dim restoreSource As String
    Dim restoreDescription As String

    On Error GoTo SubPathFail

    Set sp = crv.SubPaths(subpathIndex)
    stats.SubpathsSeen = stats.SubpathsSeen + 1

    If Not sp.Closed Then
        stats.OpenSubpathsSkipped = stats.OpenSubpathsSkipped + 1
        Exit Sub
    End If

    If sp.IsClockwise <> targetClockwise Then
        mutationOperation = "pembalikan arah"
        Set mutationBackup = crv.GetCopy
        If mutationBackup Is Nothing Then Err.Raise ERR_DIRECTION_FAILED, "ProcessSubPath", "Backup arah tidak tersedia."
        mutationPending = True
        sp.ReverseDirection
        Set sp = crv.SubPaths(subpathIndex)
        If Not sp.Closed Or sp.IsClockwise <> targetClockwise Then
            Err.Raise ERR_DIRECTION_FAILED, "ProcessSubPath", "Validasi pembalikan arah gagal."
        End If
        mutationPending = False
        Set mutationBackup = Nothing
        stats.DirectionReversed = stats.DirectionReversed + 1
    End If

    ' Arah tangent mengikuti mode pilihan, termasuk ketika arah fallback dibalik.
    If targetClockwise Then
        targetAngle = 0#
    Else
        targetAngle = 180#
    End If
    found = SPOFindStartCandidate(sp, sourceUnit, candidate, diagnostics, False, targetAngle)
    RecordCandidateDiagnostics stats, diagnostics
    diagnosticsRecorded = True
    startClockwise = targetClockwise
    If Not found Then
        If diagnostics.ReadErrors > 0 Then
            stats.IncompleteSearches = stats.IncompleteSearches + 1
            Exit Sub
        End If

        ' Arah fallback dan startpoint-nya harus commit atau pulih bersama.
        usingReverseFallback = True
        mutationOperation = "fallback arah terbalik"
        Set mutationBackup = crv.GetCopy
        If mutationBackup Is Nothing Then Err.Raise ERR_DIRECTION_FAILED, "ProcessSubPath", "Backup fallback tidak tersedia."
        mutationPending = True
        startClockwise = Not targetClockwise
        sp.ReverseDirection
        Set sp = crv.SubPaths(subpathIndex)
        If Not sp.Closed Or sp.IsClockwise <> startClockwise Then
            Err.Raise ERR_DIRECTION_FAILED, "ProcessSubPath", "Validasi arah fallback gagal."
        End If
        diagnosticsRecorded = False
        found = SPOFindStartCandidate(sp, sourceUnit, candidate, diagnostics, True, targetAngle)
        RecordCandidateDiagnostics stats, diagnostics
        diagnosticsRecorded = True
        If Not found Then
            If diagnostics.ReadErrors > 0 Then
                stats.IncompleteSearches = stats.IncompleteSearches + 1
            Else
                stats.NoCandidateFound = stats.NoCandidateFound + 1
            End If
            fallbackNoCandidate = True
            errorNumber = diagnostics.FirstNumber
            errorSource = diagnostics.FirstSource
            errorDescription = "Tidak ada kandidat valid pada arah fallback. " & diagnostics.FirstDescription
            GoTo RollbackMutation
        End If
    End If

    If candidate.IsCurrentStart And Not usingReverseFallback Then
        stats.StartAlreadyCorrect = stats.StartAlreadyCorrect + 1
        Exit Sub
    End If

    ' Transaksi startpoint terpisah dari arah yang sudah berhasil diterapkan.
    insertsMidpoint = (candidate.SegmentIndex > 0)
    If Not usingReverseFallback Then
        mutationOperation = "pemindahan startpoint"
        Set mutationBackup = crv.GetCopy
        If mutationBackup Is Nothing Then Err.Raise ERR_START_FAILED, "ProcessSubPath", "Backup startpoint tidak tersedia."
        mutationPending = True
    End If

    If Not candidate.IsCurrentStart Then
        candidateIndex = SPOMaterializeStartCandidate(crv, subpathIndex, candidate)
        If Not SetSubPathStartNode(crv, subpathIndex, candidateIndex, startClockwise) Then
            Err.Raise ERR_START_FAILED, "ProcessSubPath", "Pemindahan startpoint gagal."
        End If
    End If

    startValid = ValidateSubPathStart(crv, subpathIndex, startClockwise, sourceUnit, candidate)
    If Not startValid Then
        stats.ValidationWarnings = stats.ValidationWarnings + 1
        Err.Raise ERR_START_VALIDATION_FAILED, "ProcessSubPath", "Validasi startpoint gagal."
    End If
    mutationPending = False
    Set mutationBackup = Nothing
    If usingReverseFallback Then
        stats.DirectionReversed = stats.DirectionReversed + 1
        stats.ReverseFallbackStarts = stats.ReverseFallbackStarts + 1
    End If
    If candidate.IsCurrentStart Then
        stats.StartAlreadyCorrect = stats.StartAlreadyCorrect + 1
        Exit Sub
    End If
    stats.StartChanged = stats.StartChanged + 1
    Select Case candidate.Tier
        Case spoStartRelativeHandle
            stats.RelativeHandleStarts = stats.RelativeHandleStarts + 1
        Case spoStartLineMidpoint
            stats.LineMidpointStarts = stats.LineMidpointStarts + 1
    End Select
    Exit Sub

SubPathFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    SaveFirstError "ProcessSubPath / direction / candidate processing"
    If Not diagnosticsRecorded Then RecordCandidateDiagnostics stats, diagnostics

RollbackMutation:
    ' Termasuk jalur tanpa kandidat: error restore harus langsung diteruskan,
    ' bukan masuk lagi ke handler yang sama dan mencoba restore dua kali.
    On Error GoTo 0
    If mutationPending Then
        ' Pemulihan dibatasi pada curve ini; tidak memanggil Undo dokumen.
        If Not TryRestoreCurve(crv, mutationBackup, subpathIndex, _
                                      restoreNumber, restoreSource, restoreDescription) Then
            Err.Raise SPO_ERR_ROLLBACK_FAILED, "ProcessSubPath / rollback " & mutationOperation, _
                      "Operasi awal: " & errorNumber & " / " & errorSource & " / " & errorDescription & vbCrLf & _
                      "Pemulihan: " & restoreNumber & " / " & restoreSource & " / " & restoreDescription
        End If
        mutationPending = False
        If mutationOperation = "pembalikan arah" Then
            stats.DirectionRollbacks = stats.DirectionRollbacks + 1
        ElseIf usingReverseFallback Then
            stats.ReverseFallbackRollbacks = stats.ReverseFallbackRollbacks + 1
            If insertsMidpoint Then stats.MidpointRollbacks = stats.MidpointRollbacks + 1
        Else
            stats.StartRollbacks = stats.StartRollbacks + 1
            If insertsMidpoint Then stats.MidpointRollbacks = stats.MidpointRollbacks + 1
        End If
    End If
    If fallbackNoCandidate Then Exit Sub
    If SPOIsFatalError(errorNumber) Then Err.Raise errorNumber, errorSource, errorDescription
    stats.Failures = stats.Failures + 1
    Err.Clear
End Sub

Private Sub RecordCandidateDiagnostics(ByRef stats As OrganizerStats, ByRef diagnostics As SPOCandidateDiagnostics)
    stats.CandidateReadErrors = stats.CandidateReadErrors + diagnostics.ReadErrors
    If diagnostics.ReadErrors = 0 Or Len(mFirstCandidateError) > 0 Then Exit Sub
    mFirstCandidateError = "Pembacaan kandidat pertama yang gagal: " & diagnostics.FirstOperation & vbCrLf & _
                           "Err.Number: " & diagnostics.FirstNumber & vbCrLf & _
                           "Err.Source: " & diagnostics.FirstSource & vbCrLf & _
                           "Err.Description: " & diagnostics.FirstDescription
End Sub

Private Function TryRestoreCurve(ByVal crv As Curve, ByVal backup As Curve, _
                                         ByVal subpathIndex As Long, ByRef errorNumber As Long, _
                                         ByRef errorSource As String, ByRef errorDescription As String) As Boolean
    Dim restored As SubPath
    Dim original As SubPath
    On Error GoTo RestoreFailed
    crv.CopyAssign backup
    Set restored = crv.SubPaths(subpathIndex)
    Set original = backup.SubPaths(subpathIndex)
    If crv.SubPaths.Count <> backup.SubPaths.Count Or crv.Nodes.Count <> backup.Nodes.Count Then
        Err.Raise SPO_ERR_ROLLBACK_FAILED, "TryRestoreCurve", "Jumlah subpath/node setelah pemulihan tidak sesuai."
    End If
    If restored.Closed <> original.Closed Or restored.IsClockwise <> original.IsClockwise Then
        Err.Raise SPO_ERR_ROLLBACK_FAILED, "TryRestoreCurve", "Status closed/arah setelah pemulihan tidak sesuai."
    End If
    TryRestoreCurve = True
    Exit Function
RestoreFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
End Function

Private Function SetSubPathStartNode(ByVal crv As Curve, ByVal subpathIndex As Long, _
                                     ByVal nodeIndex As Long, _
                                     ByVal targetClockwise As Boolean) As Boolean
    Dim sp As SubPath
    Dim targetNode As Node
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo StartFail

    Set sp = crv.SubPaths(subpathIndex)
    Set targetNode = sp.Nodes(nodeIndex)

    targetNode.BreakApart

    Set sp = crv.SubPaths(subpathIndex)
    If Not sp.Closed Then
        On Error Resume Next
        sp.StartNode.JoinWith sp.EndNode
        errorNumber = Err.Number
        errorSource = Err.Source
        errorDescription = Err.Description
        Err.Clear
        On Error GoTo StartFail
        If SPOIsFatalError(errorNumber) Then Err.Raise errorNumber, errorSource, errorDescription

        Set sp = crv.SubPaths(subpathIndex)
        If Not sp.Closed Then
            sp.Closed = True
        End If
    End If

    Set sp = crv.SubPaths(subpathIndex)
    If sp.IsClockwise <> targetClockwise Then
        sp.ReverseDirection
    End If

    SetSubPathStartNode = sp.Closed And (sp.IsClockwise = targetClockwise)
    Exit Function

StartFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    SaveFirstError "SetSubPathStartNode / BreakApart / JoinWith / Close"
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Function ValidateSubPathStart(ByVal crv As Curve, ByVal subpathIndex As Long, _
                                      ByVal targetClockwise As Boolean, ByVal sourceUnit As cdrUnit, _
                                      ByRef candidate As SPOStartCandidate) As Boolean
    Dim sp As SubPath
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ValidateFail

    Set sp = crv.SubPaths(subpathIndex)
    If Not sp.Closed Then Exit Function
    If sp.IsClockwise <> targetClockwise Then Exit Function
    ValidateSubPathStart = SPOValidateStartCandidate(sp, sourceUnit, candidate)
    Exit Function

ValidateFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    SaveFirstError "ValidateSubPathStart"
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Function BuildSummary(ByRef stats As OrganizerStats) As String
    BuildSummary = "Ringkasan:" & vbCrLf & _
        "Shape dicek: " & stats.ShapesSeen & vbCrLf & _
        "Shape dikonversi ke curve: " & stats.ShapesConverted & vbCrLf & _
        "Shape non-curve diabaikan: " & stats.NonCurveSkipped & vbCrLf & _
        "Subpath dicek: " & stats.SubpathsSeen & vbCrLf & _
        "Open subpath diabaikan: " & stats.OpenSubpathsSkipped & vbCrLf & _
        "Arah subpath dibalik: " & stats.DirectionReversed & vbCrLf & _
        "Start point sudah benar: " & stats.StartAlreadyCorrect & vbCrLf & _
        "Start point diubah: " & stats.StartChanged & vbCrLf & _
        "Start point diubah via handle 5%: " & stats.RelativeHandleStarts & vbCrLf & _
        "Start point diubah via midpoint garis: " & stats.LineMidpointStarts & vbCrLf & _
        "Subpath via fallback arah terbalik: " & stats.ReverseFallbackStarts & vbCrLf & _
        "Tidak ada kandidat tangent sesuai mode +/- 15 derajat: " & stats.NoCandidateFound & vbCrLf & _
        "Pencarian tanpa hasil dengan error pembacaan: " & stats.IncompleteSearches & vbCrLf & _
        "Percobaan baca kandidat gagal: " & stats.CandidateReadErrors & vbCrLf & _
        "Pembalikan arah dipulihkan: " & stats.DirectionRollbacks & vbCrLf & _
        "Pemindahan startpoint dipulihkan: " & stats.StartRollbacks & vbCrLf & _
        "Fallback arah terbalik dipulihkan: " & stats.ReverseFallbackRollbacks & vbCrLf & _
        "Midpoint dipulihkan: " & stats.MidpointRollbacks & vbCrLf & _
        "Warning validasi: " & stats.ValidationWarnings & vbCrLf & _
        "Failure: " & stats.Failures

    If mFirstErrorNumber <> 0 Then
        BuildSummary = BuildSummary & vbCrLf & vbCrLf & _
            "First error operation: " & mFirstErrorOperation & vbCrLf & _
            "Err.Number: " & mFirstErrorNumber & vbCrLf & _
            "Err.Source: " & mFirstErrorSource & vbCrLf & _
            "Err.Description: " & mFirstErrorDescription
    End If
    If Len(mFirstCandidateError) > 0 Then BuildSummary = BuildSummary & vbCrLf & vbCrLf & mFirstCandidateError
End Function

Private Sub ClearFirstError()
    mFirstErrorOperation = ""
    mFirstErrorNumber = 0
    mFirstErrorSource = ""
    mFirstErrorDescription = ""
    mFirstCandidateError = ""
End Sub

Private Sub SaveFirstError(ByVal operationName As String)
    If mFirstErrorNumber <> 0 Then Exit Sub
    If Err.Number = 0 Then Exit Sub

    mFirstErrorOperation = operationName
    mFirstErrorNumber = Err.Number
    mFirstErrorSource = Err.Source
    mFirstErrorDescription = Err.Description
End Sub

Private Sub ShowMessageAfterClose(ByVal messageText As String, _
                                  ByVal messageStyle As VbMsgBoxStyle)
    Me.Hide
    DoEvents
    MsgBox messageText, messageStyle, "Start Point Organizer"
    Unload Me
End Sub

' Called only by MRTargetBridge; normal menu entry points remain unchanged.
Public Sub MRBindRunner(ByVal observer As Object, ByVal token As String)
    Set pMRObserver = observer
    pMRToken = token
End Sub

Public Sub MRDetachRunner()
    Set pMRObserver = Nothing
    pMRToken = vbNullString
End Sub

Private Sub UserForm_Terminate()
    Dim observer As Object, token As String
    On Error GoTo NotifyFailed
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
    Exit Sub
NotifyFailed:
    MsgBox "Gagal memberitahu Macro Runner bahwa form sudah ditutup (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub
