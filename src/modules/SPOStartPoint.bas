Option Explicit

' Pengaturan geometri eksperimental. Ambang batas panjang selalu dalam satuan milimeter.
Private Const ANGLE_TOLERANCE As Double = 15#
Private Const MAX_TANGENT_DIFFERENCE As Double = 15#
Private Const MIN_START_CONTROL_LENGTH_MM As Double = 0.5
Private Const MIN_START_CONTROL_SEGMENT_RATIO As Double = 0.05
Private Const MIN_STRAIGHT_SEGMENT_LENGTH_MM As Double = 2#
Private Const GEOMETRY_EPSILON_MM As Double = 0.000001
' Corel membulatkan koordinat dan panjang segmen yang disisipkan ke grid internalnya.
Private Const MIDPOINT_TOLERANCE_MM As Double = 0.0002
' Tangani pembulatan floating-point hanya pada batas panjang/sudut inklusif.
Private Const NUMERIC_EPSILON As Double = 0.000000001

Public Const SPO_ERR_ROLLBACK_FAILED As Long = vbObjectError + 5120

Private Enum SPOCandidateResult
    spoCandidateIneligible = 0
    spoCandidateEligible = 1
    spoCandidateReadError = 2
End Enum

Public Type SPOCandidateDiagnostics
    ReadErrors As Long ' Jumlah percobaan baca yang gagal, bukan jumlah node unik.
    FirstNumber As Long
    FirstSource As String
    FirstDescription As String
    FirstOperation As String
End Type

Public Enum SPOStartTier
    spoStartNone = 0
    spoStartFixedHandle = 1
    spoStartRelativeHandle = 2
    spoStartLineMidpoint = 3
    spoStartReverseLineMidpoint = 4
End Enum

Public Type SPOStartCandidate
    Tier As SPOStartTier
    NodeIndex As Long
    SegmentIndex As Long ' Nonzero only when a midpoint node must be inserted.
    IsCurrentStart As Boolean
    AngleDifference As Double
    TargetAngle As Double
    PositionX As Double
    PositionY As Double
End Type

Public Function SPOFindStartCandidate(ByVal sp As SubPath, ByVal sourceUnit As cdrUnit, _
                                      ByRef candidate As SPOStartCandidate, _
                                      ByRef diagnostics As SPOCandidateDiagnostics, _
                                      Optional ByVal reverseFallback As Boolean = False, _
                                      Optional ByVal targetAngle As Double = 0#) As Boolean
    Dim tier As Long
    Dim firstTier As Long
    Dim lastTier As Long
    Dim itemIndex As Long
    Dim trial As SPOStartCandidate
    Dim mmPerUnit As Double
    Dim readErrorsBeforeTier As Long

    candidate.Tier = spoStartNone
    candidate.NodeIndex = 0
    candidate.SegmentIndex = 0
    candidate.IsCurrentStart = False
    candidate.AngleDifference = ANGLE_TOLERANCE + 1#
    candidate.TargetAngle = targetAngle
    candidate.PositionX = 0#
    candidate.PositionY = 0#
    diagnostics.ReadErrors = 0
    diagnostics.FirstNumber = 0
    diagnostics.FirstSource = ""
    diagnostics.FirstDescription = ""
    diagnostics.FirstOperation = ""
    If Not sp.Closed Then Exit Function
    mmPerUnit = Application.ConvertUnits(1#, sourceUnit, cdrMillimeter)

    ' Selesaikan satu tingkatan secara utuh sebelum mempertimbangkan tingkatan berikutnya.
    ' Temuan awal yang hanya memenuhi syarat 5% tidak boleh menutupi kandidat berukuran 0,5 mm di tempat lain.
    firstTier = spoStartFixedHandle
    lastTier = spoStartLineMidpoint
    If reverseFallback Then
        ' Caller sudah membalik arah dalam transaksi; tier ini hanya mencari garis.
        firstTier = spoStartReverseLineMidpoint
        lastTier = spoStartReverseLineMidpoint
    End If
    For tier = firstTier To lastTier
        readErrorsBeforeTier = diagnostics.ReadErrors
        If TryReadNodeCandidate(sp, 0, tier, mmPerUnit, targetAngle, trial, diagnostics) = spoCandidateEligible Then
            CopyCandidate candidate, trial
        End If

        For itemIndex = 1 To sp.Nodes.Count
            If TryReadNodeCandidate(sp, itemIndex, tier, mmPerUnit, targetAngle, trial, diagnostics) = spoCandidateEligible Then
                If IsBetterCandidate(trial, candidate, reverseFallback) Then
                    CopyCandidate candidate, trial
                End If
            End If
        Next itemIndex

        If tier = spoStartLineMidpoint Or tier = spoStartReverseLineMidpoint Then
            For itemIndex = 1 To sp.Segments.Count
                If TryReadLineCandidate(sp, itemIndex, tier, mmPerUnit, targetAngle, trial, diagnostics) = spoCandidateEligible Then
                    If IsBetterCandidate(trial, candidate, reverseFallback) Then
                        CopyCandidate candidate, trial
                    End If
                End If
            Next itemIndex
        End If

        If candidate.Tier <> spoStartNone Then
            SPOFindStartCandidate = True
            Exit Function
        End If
        ' Error baca bukan bukti bahwa tier ini tidak memiliki kandidat valid.
        ' Tetap cari kandidat sehat dalam tier yang sama, tetapi jangan turun tier.
        If diagnostics.ReadErrors > readErrorsBeforeTier Then Exit Function
    Next tier
End Function

Private Function IsBetterCandidate(ByRef trial As SPOStartCandidate, ByRef current As SPOStartCandidate, _
                                   ByVal lowestFirst As Boolean) As Boolean
    If current.Tier = spoStartNone Then
        IsBetterCandidate = True
    ElseIf trial.PositionY <> current.PositionY Then
        If lowestFirst Then
            IsBetterCandidate = (trial.PositionY < current.PositionY)
        Else
            IsBetterCandidate = (trial.PositionY > current.PositionY)
        End If
    Else
        IsBetterCandidate = (trial.PositionX < current.PositionX)
    End If
End Function

Private Function TryReadNodeCandidate(ByVal sp As SubPath, ByVal itemIndex As Long, _
                                      ByVal tier As SPOStartTier, ByVal mmPerUnit As Double, _
                                      ByVal targetAngle As Double, _
                                      ByRef trial As SPOStartCandidate, _
                                      ByRef diagnostics As SPOCandidateDiagnostics) As SPOCandidateResult
    Dim operation As String
    On Error GoTo ReadFailed
    If ReadNodeCandidate(sp, itemIndex, tier, mmPerUnit, targetAngle, trial, operation) Then
        TryReadNodeCandidate = spoCandidateEligible
    End If
    Exit Function
ReadFailed:
    RecordCandidateReadError diagnostics, "Node " & itemIndex & ", tier " & tier & " / " & operation
    TryReadNodeCandidate = spoCandidateReadError
End Function

Private Function ReadNodeCandidate(ByVal sp As SubPath, ByVal itemIndex As Long, _
                                   ByVal tier As SPOStartTier, ByVal mmPerUnit As Double, _
                                   ByVal targetAngle As Double, _
                                   ByRef trial As SPOStartCandidate, ByRef operation As String) As Boolean
    Dim n As Node
    Dim diff As Double
    operation = "Membaca node (index 0 = start lama)"
    If itemIndex = 0 Then
        Set n = sp.StartNode
    Else
        Set n = sp.Nodes(itemIndex)
    End If
    operation = "EvaluateStartCandidate"
    If Not EvaluateStartCandidate(sp, n, tier, mmPerUnit, targetAngle, diff) Then Exit Function
    operation = "Membaca index dan posisi node"
    StoreNodeCandidate trial, n, tier, targetAngle, diff
    trial.IsCurrentStart = (itemIndex = 0)
    ReadNodeCandidate = True
End Function

Private Function TryReadLineCandidate(ByVal sp As SubPath, ByVal itemIndex As Long, _
                                      ByVal tier As SPOStartTier, ByVal mmPerUnit As Double, _
                                      ByVal targetAngle As Double, _
                                      ByRef trial As SPOStartCandidate, _
                                      ByRef diagnostics As SPOCandidateDiagnostics) As SPOCandidateResult
    Dim operation As String
    On Error GoTo ReadFailed
    If ReadLineCandidate(sp, itemIndex, tier, mmPerUnit, targetAngle, trial, operation) Then
        TryReadLineCandidate = spoCandidateEligible
    End If
    Exit Function
ReadFailed:
    RecordCandidateReadError diagnostics, "Segment " & itemIndex & ", tier " & tier & " / " & operation
    TryReadLineCandidate = spoCandidateReadError
End Function

Private Function ReadLineCandidate(ByVal sp As SubPath, ByVal itemIndex As Long, _
                                   ByVal tier As SPOStartTier, ByVal mmPerUnit As Double, _
                                   ByVal targetAngle As Double, _
                                   ByRef trial As SPOStartCandidate, _
                                   ByRef operation As String) As Boolean
    Dim seg As Segment
    Dim diff As Double
    Dim nodeDiff As Double
    operation = "Membaca segment"
    Set seg = sp.Segments(itemIndex)
    operation = "EvaluateLineSegment"
    If Not EvaluateLineSegment(seg, mmPerUnit, targetAngle, diff) Then Exit Function
    ' Node midpoint lama sudah diperiksa dalam loop node. Jangan memilih
    ' midpoint dari separuh garisnya karena urutan Y/X, pada pencarian atas maupun bawah.
    operation = "Mengenali endpoint midpoint lama"
    If EvaluateStartCandidate(sp, seg.StartNode, tier, mmPerUnit, targetAngle, nodeDiff) Then Exit Function
    If EvaluateStartCandidate(sp, seg.EndNode, tier, mmPerUnit, targetAngle, nodeDiff) Then Exit Function
    operation = "Membaca index dan endpoint segment"
    trial.Tier = tier
    trial.NodeIndex = 0
    trial.SegmentIndex = seg.Index
    trial.IsCurrentStart = False
    trial.AngleDifference = diff
    trial.TargetAngle = targetAngle
    trial.PositionX = (seg.StartNode.PositionX + seg.EndNode.PositionX) / 2#
    trial.PositionY = (seg.StartNode.PositionY + seg.EndNode.PositionY) / 2#
    ReadLineCandidate = True
End Function

Private Sub CopyCandidate(ByRef destination As SPOStartCandidate, ByRef source As SPOStartCandidate)
    destination.Tier = source.Tier
    destination.NodeIndex = source.NodeIndex
    destination.SegmentIndex = source.SegmentIndex
    destination.IsCurrentStart = source.IsCurrentStart
    destination.AngleDifference = source.AngleDifference
    destination.TargetAngle = source.TargetAngle
    destination.PositionX = source.PositionX
    destination.PositionY = source.PositionY
End Sub

Private Sub RecordCandidateReadError(ByRef diagnostics As SPOCandidateDiagnostics, ByVal operation As String)
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    If SPOIsFatalError(errorNumber) Then Err.Raise errorNumber, errorSource, errorDescription
    diagnostics.ReadErrors = diagnostics.ReadErrors + 1
    If diagnostics.FirstNumber = 0 Then
        diagnostics.FirstNumber = errorNumber
        diagnostics.FirstSource = errorSource
        diagnostics.FirstDescription = errorDescription
        diagnostics.FirstOperation = operation
    End If
    Err.Clear
End Sub

Public Function SPOIsFatalError(ByVal errorNumber As Long) As Boolean
    Select Case errorNumber
        Case SPO_ERR_ROLLBACK_FAILED, 7, 14, 18, _
             &H80010108, &H80010007, &H80010012, &H800401FD, &H800706BA, &H800706BE
            ' Host terputus, server mati, kehabisan memori, atau interupsi pengguna.
            SPOIsFatalError = True
    End Select
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
    SPOValidateStartCandidate = EvaluateStartCandidate(sp, n, expected.Tier, mmPerUnit, expected.TargetAngle, diff)
End Function

Private Function EvaluateStartCandidate(ByVal sp As SubPath, ByVal n As Node, _
                                         ByVal tier As SPOStartTier, ByVal mmPerUnit As Double, _
                                         ByVal targetAngle As Double, _
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

    If tier = spoStartLineMidpoint Or tier = spoStartReverseLineMidpoint Then
        ' Gunakan kembali titik tengah yang disisipkan pada tahap pengerjaan sebelumnya:
        ' gabungan dua bagian garis yang sama panjang dan segaris harus tetap memenuhi ambang batas awal sebesar 2 mm.
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

        ' Handle masuk menetapkan garis singgung kedatangan yang pasti;
        ' handle ini tidak perlu memenuhi ambang batas 0,5 mm / 5% milik handle keluar.
        If incoming.Type = cdrCurveSegment Then
            If Abs(incoming.EndingControlPointLength) * mmPerUnit <= GEOMETRY_EPSILON_MM Then Exit Function
        End If
        outgoingAngle = OutgoingTangent(outgoing)
        incomingAngle = incoming.GetTangentAt(1#, cdrRelativeSegmentOffset)
        If AngleDistance(incomingAngle, outgoingAngle) > MAX_TANGENT_DIFFERENCE + NUMERIC_EPSILON Then Exit Function
    End If

    diff = AngleDistance(outgoingAngle, targetAngle)
    EvaluateStartCandidate = (diff <= ANGLE_TOLERANCE + NUMERIC_EPSILON)
End Function

Private Function EvaluateLineSegment(ByVal seg As Segment, ByVal mmPerUnit As Double, _
                                      ByVal targetAngle As Double, _
                                      ByRef diff As Double) As Boolean
    If seg.Type <> cdrLineSegment Then Exit Function
    If Not MeetsMinimum(Abs(seg.Length) * mmPerUnit, MIN_STRAIGHT_SEGMENT_LENGTH_MM) Then Exit Function
    ' Titik interior dari suatu ruas garis tidak memiliki sudut pada titik awal yang baru.
    diff = AngleDistance(seg.GetTangentAt(0.5, cdrRelativeSegmentOffset), targetAngle)
    EvaluateLineSegment = (diff <= ANGLE_TOLERANCE + NUMERIC_EPSILON)
End Function

Private Function OutgoingTangent(ByVal seg As Segment) As Double
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    On Error GoTo ControlAngleFallback
    OutgoingTangent = seg.GetTangentAt(0#, cdrRelativeSegmentOffset)
    Exit Function

ControlAngleFallback:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    If SPOIsFatalError(errorNumber) Then Err.Raise errorNumber, errorSource, errorDescription
    ' Proses validasi kelayakan telah menolak handle awal dengan panjang nol.
    ' Pertahankan API control-angle sebelumnya sebagai mekanisme cadangan yang terbatas.
    Err.Clear
    OutgoingTangent = seg.StartingControlPointAngle
End Function

Private Sub StoreNodeCandidate(ByRef candidate As SPOStartCandidate, ByVal n As Node, _
                                ByVal tier As SPOStartTier, ByVal targetAngle As Double, ByVal diff As Double)
    candidate.Tier = tier
    candidate.NodeIndex = n.Index
    candidate.SegmentIndex = 0
    candidate.AngleDifference = diff
    candidate.TargetAngle = targetAngle
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
