Option Explicit

Private Const SPO_ROW_TOLERANCE_MM As Double = 10#
Private Const SPO_MIN_SUBPATHS As Long = 3

Private Type SPOZOrderStats
    GroupsProcessed As Long
    SubgroupsFlattened As Long
    GroupCurvesSplit As Long
    GroupMembersOrdered As Long
    SingleCurvesOrdered As Long
    SmallSingleCurvesSkipped As Long
    NonCurvesSkipped As Long
    SelectionRestoreFailed As Boolean
End Type

Public Function SPOZOrderSelected(ByVal doc As Document, _
                                  ByVal selectedShapes As ShapeRange) As String
    Dim targets As Collection
    Dim processedTargets As Collection
    Dim shp As Shape
    Dim processedGroup As Shape
    Dim stats As SPOZOrderStats
    Dim previousUnit As cdrUnit
    Dim unitChanged As Boolean
    Dim commandGroupOpen As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim operation As String

    On Error GoTo RunFailed

    Set targets = New Collection
    Set processedTargets = New Collection
    For Each shp In selectedShapes
        targets.Add shp
    Next shp

    operation = "Membaca unit dokumen"
    previousUnit = doc.Unit
    operation = "Memulai CommandGroup"
    doc.BeginCommandGroup "Start Point Organizer - ZOrder"
    commandGroupOpen = True

    If previousUnit <> cdrMillimeter Then
        operation = "Mengatur unit milimeter"
        unitChanged = True
        doc.Unit = cdrMillimeter
    End If

    For Each shp In targets
        Select Case shp.Type
            Case cdrGroupShape
                operation = "Mengurutkan anggota group"
                Set processedGroup = SPOProcessGroup(shp, stats, operation)
                processedTargets.Add processedGroup
            Case cdrCurveShape
                operation = "Mengurutkan subpath curve tunggal"
                If shp.Curve.SubPaths.Count >= SPO_MIN_SUBPATHS Then
                    SPOProcessSingleCurve shp, doc, stats
                Else
                    stats.SmallSingleCurvesSkipped = stats.SmallSingleCurvesSkipped + 1
                End If
                processedTargets.Add shp
            Case Else
                stats.NonCurvesSkipped = stats.NonCurvesSkipped + 1
                processedTargets.Add shp
        End Select
    Next shp

    operation = "Memulihkan selection"
    If Not SPORestoreSelection(processedTargets) Then stats.SelectionRestoreFailed = True

    If unitChanged Then
        operation = "Memulihkan unit dokumen"
        doc.Unit = previousUnit
        unitChanged = False
    End If
    operation = "Menutup CommandGroup"
    doc.EndCommandGroup
    commandGroupOpen = False

    SPOZOrderSelected = SPOBuildZOrderSummary(stats)
    Exit Function

RunFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not targets Is Nothing Then SPORestoreSelection targets
    If unitChanged Then doc.Unit = previousUnit
    If commandGroupOpen Then doc.EndCommandGroup
    On Error GoTo 0
    Err.Raise errorNumber, "SPOZOrderSelected", operation & ": " & errorDescription
End Function

Private Function SPOProcessGroup(ByVal groupShape As Shape, _
                                 ByRef stats As SPOZOrderStats, _
                                 ByRef operation As String) As Shape
    Dim nestedCount As Long
    Dim child As Shape
    Dim children As Collection
    Dim subpathCount As Long
    Dim shapeCountBefore As Long
    Dim subpathTotalBefore As Long
    Dim expectedShapeCount As Long
    Dim i As Long

    stats.GroupsProcessed = stats.GroupsProcessed + 1

    operation = "Meratakan subgroup"
    nestedCount = SPOCountNestedGroups(groupShape)
    If nestedCount > 0 Then
        groupShape.Shapes.All.UngroupAll
        stats.SubgroupsFlattened = stats.SubgroupsFlattened + nestedCount
        If SPOCountNestedGroups(groupShape) <> 0 Then
            Err.Raise vbObjectError + 5101, "SPOProcessGroup", _
                      "Subgroup masih ada setelah UngroupAll."
        End If
    End If

    operation = "Membaca anggota group"
    shapeCountBefore = groupShape.Shapes.Count
    subpathTotalBefore = SPOCountGroupSubpaths(groupShape)
    expectedShapeCount = shapeCountBefore
    Set children = New Collection
    For Each child In groupShape.Shapes
        children.Add child
    Next child

    For i = 1 To children.Count
        Set child = children(i)
        If child.Type = cdrCurveShape Then
            subpathCount = child.Curve.SubPaths.Count
            If subpathCount >= SPO_MIN_SUBPATHS Then
                operation = "Memisahkan subpath curve anggota " & i & _
                            " dari " & children.Count
                SPOExpandCurveInGroup groupShape, child, subpathCount
                expectedShapeCount = expectedShapeCount + subpathCount - 1
                stats.GroupCurvesSplit = stats.GroupCurvesSplit + 1
            End If
        End If
    Next i

    operation = "Memeriksa hasil pemisahan group"
    If groupShape.Shapes.Count <> expectedShapeCount Then
        Err.Raise vbObjectError + 5104, "SPOProcessGroup", _
                  "Jumlah anggota group tidak sesuai. Sebelum: " & _
                  shapeCountBefore & "; seharusnya: " & expectedShapeCount & _
                  "; sesudah: " & groupShape.Shapes.Count & "."
    End If
    If SPOCountGroupSubpaths(groupShape) <> subpathTotalBefore Then
        Err.Raise vbObjectError + 5109, "SPOProcessGroup", _
                  "Total subpath dalam group berubah setelah pemisahan."
    End If

    operation = "Mengurutkan anggota group"
    SPOOrderGroupMembers groupShape, stats
    Set SPOProcessGroup = groupShape
End Function

Private Sub SPOExpandCurveInGroup(ByVal groupShape As Shape, _
                                  ByVal sourceShape As Shape, _
                                  ByVal subpathCount As Long)
    Dim sourceCurve As Curve
    Dim subCurve As Curve
    Dim duplicate As Shape
    Dim parent As Shape
    Dim created As Collection
    Dim createdShape As Shape
    Dim sourceChanged As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim i As Long

    On Error GoTo SplitFailed
    Set created = New Collection
    Set sourceCurve = sourceShape.Curve.GetCopy
    For i = 2 To subpathCount
        Set duplicate = sourceShape.Duplicate(0, 0)
        If duplicate Is Nothing Then
            Err.Raise vbObjectError + 5106, "SPOExpandCurveInGroup", _
                      "Duplicate tidak mengembalikan curve."
        End If
        created.Add duplicate
        Set parent = duplicate.ParentGroup
        If parent Is Nothing Then
            Err.Raise vbObjectError + 5107, "SPOExpandCurveInGroup", _
                      "Salinan curve tidak masuk ke group induk."
        End If
        If Not (parent Is groupShape) Then
            Err.Raise vbObjectError + 5107, "SPOExpandCurveInGroup", _
                      "Salinan curve masuk ke group yang berbeda."
        End If
        Set subCurve = sourceCurve.SubPaths(i).GetCopy
        duplicate.Curve.CopyAssign subCurve
        If duplicate.Curve.SubPaths.Count <> 1 Then
            Err.Raise vbObjectError + 5108, "SPOExpandCurveInGroup", _
                      "Salinan curve masih memiliki beberapa subpath."
        End If
    Next i

    Set subCurve = sourceCurve.SubPaths(1).GetCopy
    sourceChanged = True
    sourceShape.Curve.CopyAssign subCurve
    If sourceShape.Curve.SubPaths.Count <> 1 Then
        Err.Raise vbObjectError + 5110, "SPOExpandCurveInGroup", _
                  "Curve sumber masih memiliki beberapa subpath."
    End If
    Exit Sub

SplitFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If sourceChanged Then sourceShape.Curve.CopyAssign sourceCurve
    If Not created Is Nothing Then
        For Each createdShape In created
            createdShape.Delete
        Next createdShape
    End If
    On Error GoTo 0
    Err.Raise errorNumber, "SPOExpandCurveInGroup", errorDescription
End Sub

Private Function SPOCountNestedGroups(ByVal groupShape As Shape) As Long
    Dim child As Shape

    For Each child In groupShape.Shapes
        If child.Type = cdrGroupShape Then
            SPOCountNestedGroups = SPOCountNestedGroups + 1 + _
                                   SPOCountNestedGroups(child)
        End If
    Next child
End Function

Private Function SPOCountGroupSubpaths(ByVal groupShape As Shape) As Long
    Dim child As Shape

    For Each child In groupShape.Shapes
        If child.Type = cdrCurveShape Then
            SPOCountGroupSubpaths = SPOCountGroupSubpaths + child.Curve.SubPaths.Count
        End If
    Next child
End Function

Private Sub SPOOrderGroupMembers(ByVal groupShape As Shape, _
                                 ByRef stats As SPOZOrderStats)
    Dim member As Shape
    Dim eligible() As Shape
    Dim centerX() As Double
    Dim centerY() As Double
    Dim order() As Long
    Dim eligibleCount As Long
    Dim x As Double
    Dim y As Double
    Dim width As Double
    Dim height As Double
    Dim previous As Shape
    Dim parent As Shape
    Dim i As Long

    If groupShape.Shapes.Count = 0 Then Exit Sub
    ReDim eligible(1 To groupShape.Shapes.Count)
    ReDim centerX(1 To groupShape.Shapes.Count)
    ReDim centerY(1 To groupShape.Shapes.Count)

    For Each member In groupShape.Shapes
        If member.Type = cdrCurveShape Then
            eligibleCount = eligibleCount + 1
            Set eligible(eligibleCount) = member
            member.GetBoundingBox x, y, width, height
            centerX(eligibleCount) = x + width / 2#
            centerY(eligibleCount) = y + height / 2#
        Else
            stats.NonCurvesSkipped = stats.NonCurvesSkipped + 1
        End If
    Next member

    If eligibleCount < 2 Then Exit Sub

    SPOBuildOrder centerX, centerY, eligibleCount, order
    Set previous = eligible(order(1))
    For i = 2 To eligibleCount
        If Not eligible(order(i)).OrderIsInFrontOf(previous) Then
            eligible(order(i)).OrderFrontOf previous
        End If
        Set parent = eligible(order(i)).ParentGroup
        If parent Is Nothing Then
            Err.Raise vbObjectError + 5107, "SPOOrderGroupMembers", _
                      "Anggota keluar dari group setelah pengurutan ZOrder."
        End If
        If Not (parent Is groupShape) Then
            Err.Raise vbObjectError + 5107, "SPOOrderGroupMembers", _
                      "Anggota keluar dari group setelah pengurutan ZOrder."
        End If
        If Not eligible(order(i)).OrderIsInFrontOf(previous) Then
            Err.Raise vbObjectError + 5108, "SPOOrderGroupMembers", _
                      "Urutan ZOrder anggota group tidak sesuai."
        End If
        Set previous = eligible(order(i))
    Next i
    stats.GroupMembersOrdered = stats.GroupMembersOrdered + eligibleCount
End Sub

Private Sub SPOProcessSingleCurve(ByVal shp As Shape, ByVal doc As Document, _
                                  ByRef stats As SPOZOrderStats)
    Dim sourceCurve As Curve
    Dim rebuiltCurve As Curve
    Dim subCurve As Curve
    Dim sp As SubPath
    Dim centerX() As Double
    Dim centerY() As Double
    Dim order() As Long
    Dim subpathCount As Long
    Dim x As Double
    Dim y As Double
    Dim width As Double
    Dim height As Double
    Dim i As Long

    Set sourceCurve = shp.Curve.GetCopy
    subpathCount = sourceCurve.SubPaths.Count
    ReDim centerX(1 To subpathCount)
    ReDim centerY(1 To subpathCount)

    For i = 1 To subpathCount
        Set sp = sourceCurve.SubPaths(i)
        If Not sp.GetBoundingBox(x, y, width, height) Then
            Err.Raise vbObjectError + 5105, "SPOProcessSingleCurve", _
                      "Bounding box subpath tidak tersedia."
        End If
        centerX(i) = x + width / 2#
        centerY(i) = y + height / 2#
    Next i

    SPOBuildOrder centerX, centerY, subpathCount, order
    Set rebuiltCurve = CreateCurve(doc)
    ' Urutan subpath hasil AppendCurve terbaca terbalik di CorelDRAW.
    ' Append dari depan ke belakang agar hasilnya kiri-atas ke kanan-bawah.
    For i = subpathCount To 1 Step -1
        Set subCurve = sourceCurve.SubPaths(order(i)).GetCopy
        rebuiltCurve.AppendCurve subCurve
    Next i

    shp.Curve.CopyAssign rebuiltCurve
    If shp.Curve.SubPaths.Count <> subpathCount Then
        Err.Raise vbObjectError + 5106, "SPOProcessSingleCurve", _
                  "Jumlah subpath berubah setelah AppendCurve."
    End If
    stats.SingleCurvesOrdered = stats.SingleCurvesOrdered + 1
End Sub

Private Sub SPOBuildOrder(ByRef centerX() As Double, _
                          ByRef centerY() As Double, ByVal itemCount As Long, _
                          ByRef order() As Long)
    Dim i As Long
    Dim j As Long
    Dim firstInRow As Long
    Dim lastInRow As Long
    Dim candidate As Long
    Dim rowY As Double

    ReDim order(1 To itemCount)
    For i = 1 To itemCount
        order(i) = i
    Next i

    ' Tetapkan baris dari Y tertinggi. Ikat toleransi ke Y awal baris,
    ' bukan ke anggota terakhir, agar baris yang miring tidak merambat.
    For i = 2 To itemCount
        candidate = order(i)
        j = i - 1
        Do While j >= 1
            If centerY(order(j)) >= centerY(candidate) Then Exit Do
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = candidate
    Next i

    firstInRow = 1
    Do While firstInRow <= itemCount
        rowY = centerY(order(firstInRow))
        lastInRow = firstInRow
        Do While lastInRow < itemCount
            If rowY - centerY(order(lastInRow + 1)) > SPO_ROW_TOLERANCE_MM Then Exit Do
            lastInRow = lastInRow + 1
        Loop

        For i = firstInRow + 1 To lastInRow
            candidate = order(i)
            j = i - 1
            Do While j >= firstInRow
                If centerX(order(j)) <= centerX(candidate) Then Exit Do
                order(j + 1) = order(j)
                j = j - 1
            Loop
            order(j + 1) = candidate
        Next i
        firstInRow = lastInRow + 1
    Loop
End Sub

Private Function SPORestoreSelection(ByVal targets As Collection) As Boolean
    Dim restored As ShapeRange
    Dim shp As Shape

    On Error GoTo RestoreFailed
    Set restored = CreateShapeRange
    For Each shp In targets
        restored.Add shp
    Next shp
    restored.CreateSelection
    SPORestoreSelection = True
    Exit Function

RestoreFailed:
    Err.Clear
End Function

Private Function SPOBuildZOrderSummary(ByRef stats As SPOZOrderStats) As String
    SPOBuildZOrderSummary = "ZOrder selesai." & vbCrLf & _
        "Group diproses: " & stats.GroupsProcessed & vbCrLf & _
        "Subgroup diratakan: " & stats.SubgroupsFlattened & vbCrLf & _
        "Curve dalam group dipecah: " & stats.GroupCurvesSplit & vbCrLf & _
        "Anggota group diurutkan: " & stats.GroupMembersOrdered & vbCrLf & _
        "Curve tunggal diurutkan: " & stats.SingleCurvesOrdered & vbCrLf & _
        "Curve tunggal kurang dari " & SPO_MIN_SUBPATHS & " subpath dilewati: " & _
            stats.SmallSingleCurvesSkipped

    If stats.NonCurvesSkipped > 0 Then
        SPOBuildZOrderSummary = SPOBuildZOrderSummary & vbCrLf & vbCrLf & _
            "Peringatan: terdapat " & stats.NonCurvesSkipped & _
            " objek non-curve di dalam pilihan; objek tersebut dilewati."
    End If
    If stats.SelectionRestoreFailed Then
        SPOBuildZOrderSummary = SPOBuildZOrderSummary & vbCrLf & _
            "Peringatan: selection awal tidak dapat dipulihkan seluruhnya."
    End If
End Function
