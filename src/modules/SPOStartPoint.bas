Option Explicit

' Experimental geometry settings. Length thresholds are always millimeters.
Private Const ANGLE_TOLERANCE As Double = 15#
Private Const MAX_TANGENT_DIFFERENCE As Double = 15#
Private Const MIN_START_CONTROL_LENGTH_MM As Double = 0.5
Private Const MIN_START_CONTROL_SEGMENT_RATIO As Double = 0.05
Private Const MIN_STRAIGHT_SEGMENT_LENGTH_MM As Double = 2#
Private Const GEOMETRY_EPSILON_MM As Double = 0.000001
' Corel rounds inserted coordinates and segment lengths to its internal grid.
Private Const MIDPOINT_TOLERANCE_MM As Double = 0.0002
' Absorb floating-point roundoff at inclusive length/angle boundaries only.
Private Const NUMERIC_EPSILON As Double = 0.000000001

Public Enum SPOStartTier
    spoStartNone = 0
    spoStartFixedHandle = 1
    spoStartRelativeHandle = 2
    spoStartLineMidpoint = 3
End Enum

Public Type SPOStartCandidate
    Tier As SPOStartTier
    NodeIndex As Long
    SegmentIndex As Long ' Nonzero only when a midpoint node must be inserted.
    AngleDifference As Double
    PositionX As Double
    PositionY As Double
End Type

Public Function SPOFindStartCandidate(ByVal sp As SubPath, ByVal sourceUnit As cdrUnit, _
                                      ByRef candidate As SPOStartCandidate) As Boolean
    Dim tier As Long
    Dim n As Node
    Dim seg As Segment
    Dim diff As Double
    Dim mmPerUnit As Double

    candidate.Tier = spoStartNone
    candidate.NodeIndex = 0
    candidate.SegmentIndex = 0
    candidate.AngleDifference = ANGLE_TOLERANCE + 1#
    candidate.PositionX = 0#
    candidate.PositionY = 0#
    If Not sp.Closed Then Exit Function
    mmPerUnit = Application.ConvertUnits(1#, sourceUnit, cdrMillimeter)

    ' Finish an entire tier before considering the next one. A current start
    ' qualifying only for 5% must not hide a 0.5 mm candidate elsewhere.
    For tier = spoStartFixedHandle To spoStartLineMidpoint
        If EvaluateStartCandidate(sp, sp.StartNode, tier, mmPerUnit, diff) Then
            StoreNodeCandidate candidate, sp.StartNode, tier, diff
            SPOFindStartCandidate = True
            Exit Function
        End If

        For Each n In sp.Nodes
            If EvaluateStartCandidate(sp, n, tier, mmPerUnit, diff) Then
                If diff < candidate.AngleDifference Then
                    StoreNodeCandidate candidate, n, tier, diff
                End If
            End If
        Next n

        If tier = spoStartLineMidpoint Then
            For Each seg In sp.Segments
                If EvaluateLineSegment(seg, mmPerUnit, diff) Then
                    If diff < candidate.AngleDifference Then
                        candidate.Tier = spoStartLineMidpoint
                        candidate.NodeIndex = 0
                        candidate.SegmentIndex = seg.Index
                        candidate.AngleDifference = diff
                        candidate.PositionX = (seg.StartNode.PositionX + seg.EndNode.PositionX) / 2#
                        candidate.PositionY = (seg.StartNode.PositionY + seg.EndNode.PositionY) / 2#
                    End If
                End If
            Next seg
        End If

        If candidate.Tier <> spoStartNone Then
            SPOFindStartCandidate = True
            Exit Function
        End If
    Next tier
End Function

Public Function SPOMaterializeStartCandidate(ByVal crv As Curve, ByVal subpathIndex As Long, _
                                             ByRef candidate As SPOStartCandidate) As Long
    Dim sp As SubPath
    Dim midpointNode As Node

    If candidate.SegmentIndex > 0 Then
        Set sp = crv.SubPaths(subpathIndex)
        Set midpointNode = sp.Segments(candidate.SegmentIndex).AddNodeAt(0.5, cdrRelativeSegmentOffset)
        candidate.NodeIndex = midpointNode.Index
        candidate.SegmentIndex = 0
    End If
    SPOMaterializeStartCandidate = candidate.NodeIndex
End Function

Public Function SPOValidateStartCandidate(ByVal sp As SubPath, ByVal sourceUnit As cdrUnit, _
                                          ByRef expected As SPOStartCandidate) As Boolean
    Dim diff As Double
    Dim mmPerUnit As Double
    Dim n As Node

    If Not sp.Closed Then Exit Function
    If expected.Tier = spoStartNone Then Exit Function
    mmPerUnit = Application.ConvertUnits(1#, sourceUnit, cdrMillimeter)
    Set n = sp.StartNode
    ' Node indices can change after BreakApart/JoinWith; check the position.
    If Abs(n.PositionX - expected.PositionX) * mmPerUnit > MIDPOINT_TOLERANCE_MM Then Exit Function
    If Abs(n.PositionY - expected.PositionY) * mmPerUnit > MIDPOINT_TOLERANCE_MM Then Exit Function
    SPOValidateStartCandidate = EvaluateStartCandidate(sp, n, expected.Tier, mmPerUnit, diff)
End Function

Private Function EvaluateStartCandidate(ByVal sp As SubPath, ByVal n As Node, _
                                         ByVal tier As SPOStartTier, ByVal mmPerUnit As Double, _
                                         ByRef diff As Double) As Boolean
    Dim incoming As Segment
    Dim outgoing As Segment
    Dim incomingAngle As Double
    Dim outgoingAngle As Double
    Dim incomingLengthMM As Double
    Dim outgoingLengthMM As Double
    Dim handleLengthMM As Double
    Dim minimumHandleMM As Double

    Set incoming = n.PrevSegment
    Set outgoing = n.NextSegment
    ' Resolve the closed seam explicitly if Corel omits an adjacent reference.
    If incoming Is Nothing Then Set incoming = sp.LastSegment
    If outgoing Is Nothing Then Set outgoing = sp.FirstSegment
    If incoming Is Nothing Then Exit Function
    If outgoing Is Nothing Then Exit Function

    incomingLengthMM = Abs(incoming.Length) * mmPerUnit
    outgoingLengthMM = Abs(outgoing.Length) * mmPerUnit
    If incomingLengthMM <= GEOMETRY_EPSILON_MM Then Exit Function
    If outgoingLengthMM <= GEOMETRY_EPSILON_MM Then Exit Function

    If tier = spoStartLineMidpoint Then
        ' Reuse a midpoint inserted on an earlier run: two equal, collinear
        ' line halves together must still meet the original 2 mm threshold.
        If incoming.Type <> cdrLineSegment Then Exit Function
        If outgoing.Type <> cdrLineSegment Then Exit Function
        If Abs(incomingLengthMM - outgoingLengthMM) > MIDPOINT_TOLERANCE_MM + NUMERIC_EPSILON Then Exit Function
        If Abs(n.PositionX - (incoming.StartNode.PositionX + outgoing.EndNode.PositionX) / 2#) * _
                mmPerUnit > MIDPOINT_TOLERANCE_MM Then Exit Function
        If Abs(n.PositionY - (incoming.StartNode.PositionY + outgoing.EndNode.PositionY) / 2#) * _
                mmPerUnit > MIDPOINT_TOLERANCE_MM Then Exit Function
        If Not MeetsMinimum(incomingLengthMM + outgoingLengthMM, MIN_STRAIGHT_SEGMENT_LENGTH_MM) Then Exit Function
        incomingAngle = incoming.GetTangentAt(1#, cdrRelativeSegmentOffset)
        outgoingAngle = outgoing.GetTangentAt(0#, cdrRelativeSegmentOffset)
        If AngleDistance(incomingAngle, outgoingAngle) > MAX_TANGENT_DIFFERENCE + NUMERIC_EPSILON Then Exit Function
    Else
        If outgoing.Type <> cdrCurveSegment Then Exit Function
        handleLengthMM = Abs(outgoing.StartingControlPointLength) * mmPerUnit
        If handleLengthMM <= GEOMETRY_EPSILON_MM Then Exit Function
        Select Case tier
            Case spoStartFixedHandle
                minimumHandleMM = MIN_START_CONTROL_LENGTH_MM
            Case spoStartRelativeHandle
                minimumHandleMM = outgoingLengthMM * MIN_START_CONTROL_SEGMENT_RATIO
            Case Else
                Exit Function
        End Select
        If Not MeetsMinimum(handleLengthMM, minimumHandleMM) Then Exit Function

        ' The incoming handle establishes a defined arrival tangent; it does
        ' not need to satisfy the outgoing handle's 0.5 mm / 5% threshold.
        If incoming.Type = cdrCurveSegment Then
            If Abs(incoming.EndingControlPointLength) * mmPerUnit <= GEOMETRY_EPSILON_MM Then Exit Function
        End If
        outgoingAngle = OutgoingTangent(outgoing)
        incomingAngle = incoming.GetTangentAt(1#, cdrRelativeSegmentOffset)
        If AngleDistance(incomingAngle, outgoingAngle) > MAX_TANGENT_DIFFERENCE + NUMERIC_EPSILON Then Exit Function
    End If

    diff = AngleDistance(outgoingAngle, 0#)
    EvaluateStartCandidate = (diff <= ANGLE_TOLERANCE + NUMERIC_EPSILON)
End Function

Private Function EvaluateLineSegment(ByVal seg As Segment, ByVal mmPerUnit As Double, _
                                      ByRef diff As Double) As Boolean
    If seg.Type <> cdrLineSegment Then Exit Function
    If Not MeetsMinimum(Abs(seg.Length) * mmPerUnit, MIN_STRAIGHT_SEGMENT_LENGTH_MM) Then Exit Function
    ' An interior point of a straight segment has no corner at the new start.
    diff = AngleDistance(seg.GetTangentAt(0.5, cdrRelativeSegmentOffset), 0#)
    EvaluateLineSegment = (diff <= ANGLE_TOLERANCE + NUMERIC_EPSILON)
End Function

Private Function OutgoingTangent(ByVal seg As Segment) As Double
    On Error GoTo ControlAngleFallback
    OutgoingTangent = seg.GetTangentAt(0#, cdrRelativeSegmentOffset)
    Exit Function

ControlAngleFallback:
    ' Eligibility has already rejected a zero-length starting handle.
    ' Keep the previous control-angle API as a narrow fallback.
    Err.Clear
    OutgoingTangent = seg.StartingControlPointAngle
End Function

Private Sub StoreNodeCandidate(ByRef candidate As SPOStartCandidate, ByVal n As Node, _
                                ByVal tier As SPOStartTier, ByVal diff As Double)
    candidate.Tier = tier
    candidate.NodeIndex = n.Index
    candidate.SegmentIndex = 0
    candidate.AngleDifference = diff
    candidate.PositionX = n.PositionX
    candidate.PositionY = n.PositionY
End Sub

Private Function MeetsMinimum(ByVal value As Double, ByVal minimum As Double) As Boolean
    MeetsMinimum = (value + NUMERIC_EPSILON >= minimum)
End Function

Private Function AngleDistance(ByVal firstAngle As Double, ByVal secondAngle As Double) As Double
    Dim distance As Double
    distance = firstAngle - secondAngle
    distance = distance - 360# * Int(distance / 360#)
    If distance > 180# Then distance = 360# - distance
    AngleDistance = distance
End Function
