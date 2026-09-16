Option Explicit

' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String


Private Const ANGLE_TOLERANCE As Double = 15#
Private Const MIN_CONTROL_LENGTH As Double = 0.000001
Private Const MIN_START_CONTROL_LENGTH As Double = 0.5

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
    Dim previousUnit As cdrUnit
    Dim unitChanged As Boolean
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

    On Error Resume Next
    Err.Clear
    previousUnit = doc.Unit
    doc.Unit = cdrMillimeter
    unitChanged = (Err.Number = 0)
    Err.Clear
    doc.BeginCommandGroup "Start Point Organizer"
    commandGroupOpen = (Err.Number = 0)
    Err.Clear
    On Error GoTo FatalFail

    For Each shp In selectedShapes
        ProcessShape shp, targetClockwise, stats
    Next shp

CleanExit:
    If commandGroupOpen Then
        On Error Resume Next
        doc.EndCommandGroup
        On Error GoTo 0
    End If
    If unitChanged Then
        On Error Resume Next
        doc.Unit = previousUnit
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
    If unitChanged Then
        On Error Resume Next
        doc.Unit = previousUnit
        On Error GoTo 0
    End If

    ShowMessageAfterClose "Macro berhenti karena error." & vbCrLf & _
                          "Err.Number: " & fatalNumber & vbCrLf & _
                          "Err.Description: " & fatalDescription & vbCrLf & vbCrLf & _
                          BuildSummary(stats), vbCritical
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
                         ByRef stats As OrganizerStats)
    Dim childShape As Shape
    Dim crv As Curve
    Dim subpathCount As Long
    Dim subpathIndex As Long

    On Error GoTo ShapeFail

    stats.ShapesSeen = stats.ShapesSeen + 1

    If shp.Type = cdrGroupShape Then
        For Each childShape In shp.Shapes
            ProcessShape childShape, targetClockwise, stats
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

    subpathCount = crv.Subpaths.Count
    For subpathIndex = 1 To subpathCount
        ProcessSubPath crv, subpathIndex, targetClockwise, stats
    Next subpathIndex

    Exit Sub

ShapeFail:
    SaveFirstError "ProcessShape / ConvertToCurves / Curve access"
    stats.Failures = stats.Failures + 1
    Err.Clear
End Sub

Private Sub ProcessSubPath(ByVal crv As Curve, ByVal subpathIndex As Long, _
                           ByVal targetClockwise As Boolean, _
                           ByRef stats As OrganizerStats)
    Dim sp As SubPath
    Dim startAngle As Double
    Dim startDiff As Double
    Dim candidateIndex As Long
    Dim candidateDiff As Double

    On Error GoTo SubPathFail

    Set sp = crv.Subpaths(subpathIndex)
    stats.SubpathsSeen = stats.SubpathsSeen + 1

    If Not sp.Closed Then
        stats.OpenSubpathsSkipped = stats.OpenSubpathsSkipped + 1
        Exit Sub
    End If

    If sp.IsClockwise <> targetClockwise Then
        sp.ReverseDirection
        stats.DirectionReversed = stats.DirectionReversed + 1
        Set sp = crv.Subpaths(subpathIndex)
    End If

    If TryGetNodeOutgoingAngle(sp.StartNode, startAngle, MIN_START_CONTROL_LENGTH, False) Then
        startDiff = AngleDistanceFromZero(startAngle)
        If startDiff <= ANGLE_TOLERANCE Then
            stats.StartAlreadyCorrect = stats.StartAlreadyCorrect + 1
            Exit Sub
        End If
    End If

    candidateIndex = FindBestStartCandidate(sp, candidateDiff)
    If candidateIndex = 0 Then
        stats.NoCandidateFound = stats.NoCandidateFound + 1
        Exit Sub
    End If

    If SetSubPathStartNode(crv, subpathIndex, candidateIndex, targetClockwise) Then
        stats.StartChanged = stats.StartChanged + 1
        If Not ValidateSubPathStart(crv, subpathIndex, targetClockwise) Then
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

Private Function FindBestStartCandidate(ByVal sp As SubPath, _
                                        ByRef bestDiff As Double) As Long
    Dim n As Node
    Dim nodeAngle As Double
    Dim diff As Double
    Dim bestIndex As Long

    bestDiff = ANGLE_TOLERANCE + 1#
    bestIndex = 0

    For Each n In sp.Nodes
        If TryGetNodeOutgoingAngle(n, nodeAngle, MIN_START_CONTROL_LENGTH, False) Then
            diff = AngleDistanceFromZero(nodeAngle)
            If diff <= ANGLE_TOLERANCE Then
                If bestIndex = 0 Or diff < bestDiff Then
                    bestIndex = n.Index
                    bestDiff = diff
                End If
            End If
        End If
    Next n

    FindBestStartCandidate = bestIndex
End Function

Private Function SetSubPathStartNode(ByVal crv As Curve, ByVal subpathIndex As Long, _
                                     ByVal nodeIndex As Long, _
                                     ByVal targetClockwise As Boolean) As Boolean
    Dim sp As SubPath
    Dim targetNode As Node
    Dim joinSucceeded As Boolean

    On Error GoTo StartFail

    Set sp = crv.Subpaths(subpathIndex)
    Set targetNode = sp.Nodes(nodeIndex)

    targetNode.BreakApart

    Set sp = crv.Subpaths(subpathIndex)
    If Not sp.Closed Then
        On Error Resume Next
        sp.StartNode.JoinWith sp.EndNode
        joinSucceeded = (Err.Number = 0)
        Err.Clear
        On Error GoTo StartFail

        Set sp = crv.Subpaths(subpathIndex)
        If Not sp.Closed Then
            sp.Closed = True
        End If
    Else
        joinSucceeded = True
    End If

    Set sp = crv.Subpaths(subpathIndex)
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
                                      ByVal targetClockwise As Boolean) As Boolean
    Dim sp As SubPath
    Dim startAngle As Double

    On Error GoTo ValidateFail

    Set sp = crv.Subpaths(subpathIndex)
    If Not sp.Closed Then Exit Function
    If sp.IsClockwise <> targetClockwise Then Exit Function
    If Not TryGetNodeOutgoingAngle(sp.StartNode, startAngle, MIN_START_CONTROL_LENGTH, False) Then Exit Function

    ValidateSubPathStart = (AngleDistanceFromZero(startAngle) <= ANGLE_TOLERANCE)
    Exit Function

ValidateFail:
    SaveFirstError "ValidateSubPathStart"
    Err.Clear
End Function

Private Function TryGetNodeOutgoingAngle(ByVal n As Node, _
                                         ByRef angleValue As Double, _
                                         Optional ByVal minControlLength As Double = MIN_CONTROL_LENGTH, _
                                         Optional ByVal allowFallback As Boolean = True) As Boolean
    Dim seg As Segment

    On Error GoTo AngleFail

    Set seg = n.NextSegment
    If Not seg Is Nothing Then
        If seg.Type = cdrCurveSegment Then
            If Abs(seg.StartingControlPointLength) >= minControlLength Then
                angleValue = NormalizeAngle(seg.StartingControlPointAngle)
                TryGetNodeOutgoingAngle = True
                Exit Function
            End If
        End If
    End If

    If allowFallback Then
        TryGetNodeOutgoingAngle = TryGetAngleToNextNode(n, angleValue)
    Else
        TryGetNodeOutgoingAngle = False
    End If
    Exit Function

AngleFail:
    SaveFirstError "TryGetNodeOutgoingAngle / NextSegment / control point"
    Err.Clear
    TryGetNodeOutgoingAngle = False
End Function

Private Function TryGetAngleToNextNode(ByVal n As Node, _
                                       ByRef angleValue As Double) As Boolean
    Dim nextNode As Node
    Dim dx As Double
    Dim dy As Double

    On Error GoTo AngleFail

    Set nextNode = n.Next
    If nextNode Is Nothing Then
        If n.SubPath.Closed Then Set nextNode = n.SubPath.StartNode
    End If

    If nextNode Is Nothing Then Exit Function

    dx = nextNode.PositionX - n.PositionX
    dy = nextNode.PositionY - n.PositionY

    If Abs(dx) <= MIN_CONTROL_LENGTH And Abs(dy) <= MIN_CONTROL_LENGTH Then
        Exit Function
    End If

    angleValue = NormalizeAngle(Atan2Degrees(dy, dx))
    TryGetAngleToNextNode = True
    Exit Function

AngleFail:
    SaveFirstError "TryGetAngleToNextNode / node position fallback"
    Err.Clear
End Function

Private Function AngleDistanceFromZero(ByVal angleValue As Double) As Double
    Dim normalizedAngle As Double

    normalizedAngle = NormalizeAngle(angleValue)
    If normalizedAngle > 180# Then
        AngleDistanceFromZero = 360# - normalizedAngle
    Else
        AngleDistanceFromZero = normalizedAngle
    End If
End Function

Private Function NormalizeAngle(ByVal angleValue As Double) As Double
    Do While angleValue < 0#
        angleValue = angleValue + 360#
    Loop

    Do While angleValue >= 360#
        angleValue = angleValue - 360#
    Loop

    NormalizeAngle = angleValue
End Function

Private Function Atan2Degrees(ByVal y As Double, ByVal x As Double) As Double
    Const PI As Double = 3.14159265358979
    Dim angleRadians As Double

    If Abs(x) <= MIN_CONTROL_LENGTH Then
        If y >= 0# Then
            angleRadians = PI / 2#
        Else
            angleRadians = -PI / 2#
        End If
    Else
        angleRadians = Atn(y / x)
        If x < 0# Then
            angleRadians = angleRadians + PI
        End If
    End If

    Atan2Degrees = angleRadians * 180# / PI
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
