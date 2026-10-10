Option Explicit

' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String


Private mFirstErrorOperation As String
Private mFirstErrorNumber As Long
Private mFirstErrorDescription As String

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
    NoCandidateFound As Long
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
    Dim fatalDescription As String

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

CleanExit:
    If commandGroupOpen Then
        On Error Resume Next
        doc.EndCommandGroup
        On Error GoTo 0
    End If

    ShowMessageAfterClose BuildSummary(stats), vbInformation
    Exit Sub

FatalFail:
    fatalNumber = Err.Number
    fatalDescription = Err.Description
    stats.Failures = stats.Failures + 1
    If commandGroupOpen Then
        On Error Resume Next
        doc.EndCommandGroup
        On Error GoTo 0
    End If

    ShowMessageAfterClose "Macro berhenti karena error." & vbCrLf & _
                          "Err.Number: " & fatalNumber & vbCrLf & _
                          "Err.Description: " & fatalDescription & vbCrLf & vbCrLf & _
                          BuildSummary(stats), vbCritical
End Sub

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
        ProcessSubPath crv, subpathIndex, targetClockwise, sourceUnit, stats
    Next subpathIndex

    Exit Sub

ShapeFail:
    SaveFirstError "ProcessShape / ConvertToCurves / Curve access"
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

    On Error GoTo SubPathFail

    Set sp = crv.SubPaths(subpathIndex)
    stats.SubpathsSeen = stats.SubpathsSeen + 1

    If Not sp.Closed Then
        stats.OpenSubpathsSkipped = stats.OpenSubpathsSkipped + 1
        Exit Sub
    End If

    If sp.IsClockwise <> targetClockwise Then
        sp.ReverseDirection
        stats.DirectionReversed = stats.DirectionReversed + 1
        Set sp = crv.SubPaths(subpathIndex)
    End If

    If Not SPOFindStartCandidate(sp, sourceUnit, candidate) Then
        stats.NoCandidateFound = stats.NoCandidateFound + 1
        Exit Sub
    End If

    If candidate.SegmentIndex = 0 And candidate.NodeIndex = sp.StartNode.Index Then
        stats.StartAlreadyCorrect = stats.StartAlreadyCorrect + 1
        Exit Sub
    End If

    candidateIndex = SPOMaterializeStartCandidate(crv, subpathIndex, candidate)
    If SetSubPathStartNode(crv, subpathIndex, candidateIndex, targetClockwise) Then
        stats.StartChanged = stats.StartChanged + 1
        Select Case candidate.Tier
            Case spoStartRelativeHandle
                stats.RelativeHandleStarts = stats.RelativeHandleStarts + 1
            Case spoStartLineMidpoint
                stats.LineMidpointStarts = stats.LineMidpointStarts + 1
        End Select
        If Not ValidateSubPathStart(crv, subpathIndex, targetClockwise, sourceUnit, candidate) Then
            stats.ValidationWarnings = stats.ValidationWarnings + 1
        End If
    Else
        stats.Failures = stats.Failures + 1
    End If

    Exit Sub

SubPathFail:
    SaveFirstError "ProcessSubPath / direction / candidate processing"
    stats.Failures = stats.Failures + 1
    Err.Clear
End Sub

Private Function SetSubPathStartNode(ByVal crv As Curve, ByVal subpathIndex As Long, _
                                     ByVal nodeIndex As Long, _
                                     ByVal targetClockwise As Boolean) As Boolean
    Dim sp As SubPath
    Dim targetNode As Node
    Dim joinSucceeded As Boolean

    On Error GoTo StartFail

    Set sp = crv.SubPaths(subpathIndex)
    Set targetNode = sp.Nodes(nodeIndex)

    targetNode.BreakApart

    Set sp = crv.SubPaths(subpathIndex)
    If Not sp.Closed Then
        On Error Resume Next
        sp.StartNode.JoinWith sp.EndNode
        joinSucceeded = (Err.Number = 0)
        Err.Clear
        On Error GoTo StartFail

        Set sp = crv.SubPaths(subpathIndex)
        If Not sp.Closed Then
            sp.Closed = True
        End If
    Else
        joinSucceeded = True
    End If

    Set sp = crv.SubPaths(subpathIndex)
    If sp.IsClockwise <> targetClockwise Then
        sp.ReverseDirection
    End If

    SetSubPathStartNode = joinSucceeded Or sp.Closed
    Exit Function

StartFail:
    SaveFirstError "SetSubPathStartNode / BreakApart / JoinWith / Close"
    Err.Clear
    SetSubPathStartNode = False
End Function

Private Function ValidateSubPathStart(ByVal crv As Curve, ByVal subpathIndex As Long, _
                                      ByVal targetClockwise As Boolean, ByVal sourceUnit As cdrUnit, _
                                      ByRef candidate As SPOStartCandidate) As Boolean
    Dim sp As SubPath

    On Error GoTo ValidateFail

    Set sp = crv.SubPaths(subpathIndex)
    If Not sp.Closed Then Exit Function
    If sp.IsClockwise <> targetClockwise Then Exit Function
    ValidateSubPathStart = SPOValidateStartCandidate(sp, sourceUnit, candidate)
    Exit Function

ValidateFail:
    SaveFirstError "ValidateSubPathStart"
    Err.Clear
End Function

Private Function BuildSummary(ByRef stats As OrganizerStats) As String
    BuildSummary = "Selesai." & vbCrLf & _
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
        "Tidak ada kandidat 0 +/- 15 derajat: " & stats.NoCandidateFound & vbCrLf & _
        "Warning validasi: " & stats.ValidationWarnings & vbCrLf & _
        "Failure: " & stats.Failures

    If mFirstErrorNumber <> 0 Then
        BuildSummary = BuildSummary & vbCrLf & vbCrLf & _
            "First error operation: " & mFirstErrorOperation & vbCrLf & _
            "Err.Number: " & mFirstErrorNumber & vbCrLf & _
            "Err.Description: " & mFirstErrorDescription
    End If
End Function

Private Sub ClearFirstError()
    mFirstErrorOperation = ""
    mFirstErrorNumber = 0
    mFirstErrorDescription = ""
End Sub

Private Sub SaveFirstError(ByVal operationName As String)
    If mFirstErrorNumber <> 0 Then Exit Sub
    If Err.Number = 0 Then Exit Sub

    mFirstErrorOperation = operationName
    mFirstErrorNumber = Err.Number
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
