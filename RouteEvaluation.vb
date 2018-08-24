Imports jakStd20_MathExt
Imports jakStd20_Datums

Partial Public Class Route

    Private Structure ProjectionResults
        Dim X, Y, KP, DCC As Double
        Dim OnSegment As PointPosition
    End Structure

    Private Enum PointPosition
        Before = 0
        Between = 1
        After = 2
    End Enum

    Private Shared Function CreateGridKpLists(ReferenceRoute As Route) As List(Of RoutePoint)
        Dim segrad, chord1, chord2, swept1, swept2, chord1hdg, chord2hdg, h2c, KPoffset As Double
        Dim KPpoint As New RoutePoint
        Dim totKP, maxKP As Integer
        Dim KPlist As New List(Of RoutePoint)
        'Dim tmpPoint As New RoutePoint

        'Define the dimension of the coordinates array
        maxKP = CInt(ReferenceRoute.CompSegment(ReferenceRoute.SegmentCount - 1).EndPoint.ForwardGridKP)
        If maxKP > ReferenceRoute.CompSegment(ReferenceRoute.SegmentCount - 1).EndPoint.ForwardGridKP Then maxKP = maxKP - 1
        maxKP = maxKP + 1

        'Set the start point of the route
        KPpoint.Easting = ReferenceRoute.CompSegment(0).StartPoint.Easting
        KPpoint.Northing = ReferenceRoute.CompSegment(0).StartPoint.Northing
        KPpoint.Heading = CalcHeading(KPpoint.Easting, KPpoint.Northing, ReferenceRoute.CompSegment(0).EndPoint.Easting, ReferenceRoute.CompSegment(0).EndPoint.Northing)
        KPlist.Add(KPpoint)
        'start to evaluate the KPs coordinate along the route
        For nSeg = 0 To ReferenceRoute.SegmentCount - 1
            'evaluate the first kp of the segment
            KPoffset = ReferenceRoute.CompSegment(nSeg).StartPoint.ForwardGridKP - Math.Truncate(ReferenceRoute.CompSegment(nSeg).StartPoint.ForwardGridKP)
            For incrKP As Integer = 1 To CInt(Math.Truncate(ReferenceRoute.CompSegment(nSeg).Length))
                Dim tmpKPpoint As New RoutePoint
                If ReferenceRoute.CompSegment(nSeg).IsLine Then
                    'evaluate along a line
                    tmpKPpoint.Easting = ReferenceRoute.CompSegment(nSeg).StartPoint.Easting - ((incrKP - KPoffset) / ReferenceRoute.CompSegment(nSeg).Length *
                                                                        (ReferenceRoute.CompSegment(nSeg).StartPoint.Easting - ReferenceRoute.CompSegment(nSeg).EndPoint.Easting))
                    tmpKPpoint.Northing = ReferenceRoute.CompSegment(nSeg).StartPoint.Northing - ((incrKP - KPoffset) / ReferenceRoute.CompSegment(nSeg).Length *
                                                                        (ReferenceRoute.CompSegment(nSeg).StartPoint.Northing - ReferenceRoute.CompSegment(nSeg).EndPoint.Northing))
                    'calculate the overall kp of the point
                    totKP = CInt(ReferenceRoute.CompSegment(nSeg).StartPoint.ForwardGridKP + incrKP - KPoffset)
                Else
                    'evaluate along an arc
                    segrad = ReferenceRoute.CompSegment(nSeg).Radious
                    chord1 = Math.Sqrt((ReferenceRoute.CompSegment(nSeg).StartPoint.Easting - ReferenceRoute.CompSegment(nSeg).EndPoint.Easting) ^ 2 +
                                       (ReferenceRoute.CompSegment(nSeg).StartPoint.Northing - ReferenceRoute.CompSegment(nSeg).EndPoint.Northing) ^ 2)
                    swept1 = 2 * Math.Asin(chord1 / 2 / segrad)
                    chord1hdg = CalcHeading(ReferenceRoute.CompSegment(nSeg).StartPoint.Easting, ReferenceRoute.CompSegment(nSeg).StartPoint.Northing,
                                            ReferenceRoute.CompSegment(nSeg).EndPoint.Easting, ReferenceRoute.CompSegment(nSeg).EndPoint.Northing)
                    swept2 = (incrKP - KPoffset) / (2 * Math.PI * segrad) * (Math.PI * 2)
                    chord2 = segrad * Math.Sin(swept2 / 2) * 2
                    'evaluate the heading of the actual chord
                    If ReferenceRoute.CompSegment(nSeg).IsArcClockwise Then
                        h2c = chord1hdg - (((2 * Math.PI) - swept1) / 2) + Math.PI
                        chord2hdg = AngleFit2Pi(h2c + ((Math.PI * 2 - swept2) / 2) - Math.PI)
                    Else
                        'Left bend need to reverse the angles to draw correct
                        h2c = chord1hdg + (((2 * Math.PI) - swept1) / 2) + Math.PI
                        chord2hdg = AngleFit2Pi(h2c - ((Math.PI * 2 - swept2) / 2) - Math.PI)
                    End If
                    'calculate the coords of the point on arc
                    tmpKPpoint.ENfromPoint3D(RangeBearing(ReferenceRoute.CompSegment(nSeg).StartPoint.Easting, ReferenceRoute.CompSegment(nSeg).StartPoint.Northing, chord2, chord2hdg))
                    totKP = CInt(ReferenceRoute.CompSegment(nSeg).StartPoint.ForwardGridKP + incrKP - KPoffset)
                End If
                tmpKPpoint.Heading = CalcHeading(KPlist.Last.Easting, KPlist.Last.Northing, tmpKPpoint.Easting, tmpKPpoint.Northing)
                KPlist.Add(tmpKPpoint)
            Next 'incrKP
            'eventual extra point
            If (ReferenceRoute.CompSegment(nSeg).EndPoint.ForwardGridKP - totKP) > 1 Then
                Dim tmpKPpointL As New RoutePoint
                Dim extraKP As Integer = CInt(Math.Truncate(ReferenceRoute.CompSegment(nSeg).Length)) + 1
                If ReferenceRoute.CompSegment(nSeg).IsLine Then
                    'evaluate along a line
                    tmpKPpointL.Easting = ReferenceRoute.CompSegment(nSeg).StartPoint.Easting - ((extraKP - KPoffset) / ReferenceRoute.CompSegment(nSeg).Length *
                                                                        (ReferenceRoute.CompSegment(nSeg).StartPoint.Easting - ReferenceRoute.CompSegment(nSeg).EndPoint.Easting))
                    tmpKPpointL.Northing = ReferenceRoute.CompSegment(nSeg).StartPoint.Northing - ((extraKP - KPoffset) / ReferenceRoute.CompSegment(nSeg).Length *
                                                                        (ReferenceRoute.CompSegment(nSeg).StartPoint.Northing - ReferenceRoute.CompSegment(nSeg).EndPoint.Northing))
                    'calculate the overall kp of the point
                    totKP = CInt(ReferenceRoute.CompSegment(nSeg).StartPoint.ForwardGridKP + extraKP - KPoffset)
                Else
                    'evaluate along an arc
                    segrad = ReferenceRoute.CompSegment(nSeg).Radious
                    chord1 = Math.Sqrt((ReferenceRoute.CompSegment(nSeg).StartPoint.Easting - ReferenceRoute.CompSegment(nSeg).EndPoint.Easting) ^ 2 +
                                       (ReferenceRoute.CompSegment(nSeg).StartPoint.Northing - ReferenceRoute.CompSegment(nSeg).EndPoint.Northing) ^ 2)
                    swept1 = 2 * Math.Asin(chord1 / 2 / segrad)
                    chord1hdg = CalcHeading(ReferenceRoute.CompSegment(nSeg).StartPoint.Easting, ReferenceRoute.CompSegment(nSeg).StartPoint.Northing,
                                            ReferenceRoute.CompSegment(nSeg).EndPoint.Easting, ReferenceRoute.CompSegment(nSeg).EndPoint.Northing)
                    swept2 = (extraKP - KPoffset) / (2 * Math.PI * segrad) * (Math.PI * 2)
                    chord2 = segrad * Math.Sin(swept2 / 2) * 2
                    'evaluate the heading of the actual chord
                    If ReferenceRoute.CompSegment(nSeg).IsArcClockwise Then
                        h2c = chord1hdg - (((2 * Math.PI) - swept1) / 2) + Math.PI
                        chord2hdg = AngleFit2Pi(h2c + ((Math.PI * 2 - swept2) / 2) - Math.PI)
                    Else
                        'Left bend need to reverse the angles to draw correct
                        h2c = chord1hdg + (((2 * Math.PI) - swept1) / 2) + Math.PI
                        chord2hdg = AngleFit2Pi(h2c - ((Math.PI * 2 - swept2) / 2) - Math.PI)
                    End If
                    'calculate the coords of the point on arc
                    tmpKPpointL.ENfromPoint3D(RangeBearing(ReferenceRoute.CompSegment(nSeg).StartPoint.Easting, ReferenceRoute.CompSegment(nSeg).StartPoint.Northing, chord2, chord2hdg))
                    totKP = CInt(ReferenceRoute.CompSegment(nSeg).StartPoint.ForwardGridKP + extraKP - KPoffset)
                End If
                tmpKPpointL.Heading = CalcHeading(KPlist.Last.Easting, KPlist.Last.Northing, tmpKPpointL.Easting, tmpKPpointL.Northing)
                KPlist.Add(tmpKPpointL)
            End If
        Next 'nSeg
        'insert the end point of the route
        Dim KPpointLast As New RoutePoint
        KPpointLast.Easting = ReferenceRoute.CompSegment(ReferenceRoute.SegmentCount - 1).EndPoint.Easting
        KPpointLast.Northing = ReferenceRoute.CompSegment(ReferenceRoute.SegmentCount - 1).EndPoint.Northing
        KPpointLast.Heading = CalcHeading(KPlist.Last.Easting, KPlist.Last.Northing, KPpointLast.Easting, KPpointLast.Northing)
        KPlist.Add(KPpointLast)

        'Fill the kp fields
        Dim RteLength As Double = ReferenceRoute.CompSegment(ReferenceRoute.SegmentCount - 1).EndPoint.ForwardGridKP
        For k = 1 To KPlist.Count - 1
            KPlist(k).ForwardGridKP = k
            KPlist(k).ReverseGridKP = RteLength - k
        Next
        KPlist(0).ForwardGridKP = 0
        KPlist(0).ReverseGridKP = RteLength
        KPlist(KPlist.Count - 1).ForwardGridKP = RteLength
        KPlist(KPlist.Count - 1).ReverseGridKP = 0
        'return the maxKP value
        Return KPlist
    End Function

    Public Function ProjectPointOnRoute(dataPoint As Point3D, ByRef projectedPoint As Point3D, ByRef KP As Double, ByRef DCC As Double) As String
        If dataPoint Is Nothing Then
            Return ""
        Else
            Return ProjectPointOnRoute(dataPoint.X, dataPoint.Y, projectedPoint.X, projectedPoint.Y, KP, DCC)
        End If
    End Function

    Public Function ProjectPointOnRoute(dataPoint As Point3D) As RoutePoint
        Dim ppX, ppY, ppKP, ppDCC As Double
        Dim resPnt As New RoutePoint
        If dataPoint Is Nothing Then Return resPnt

        ProjectPointOnRoute(dataPoint.X, dataPoint.Y, ppX, ppY, ppKP, ppDCC)
        resPnt = GetRoutePointFromProgressiveDistance(ppKP, DistanceType.ForwardGrid)
        Return resPnt

    End Function

    Public Function ProjectPointOnRoute(dataPointX As Double, dataPointY As Double, ByRef projectedPointX As Double, ByRef projectedPointY As Double,
                                        ByRef KP As Double, ByRef DCC As Double) As String
        Dim A0, A1, A2, Hdg1, Hdg2, Hdg0, SpX, SpY, EpX, EpY, CpX, CpY, k As Double
        Dim DeltaAngle, DeltaDistance, SegRad, DpRadious, TmpDCC, TmpKP, TmpDist1, TmpDist2, D0, D1, D2, HdgLine, HdgPrj As Double
        Dim TmpPoint As New Point3D
        Dim ResultPoints(SegmentCount - 1) As ProjectionResults
        Dim AngleTollerance As Double = 0.000000001
        Dim DistanceTollerance As Double = 0.000001
        Dim dbgString As String

        dbgString = "SpX, SpY, EpX, EpY, CpX, CpY, Hdg0, Hdg1, Hdg2, A0, A1, A2, SegRad, DpRadious, SegID, OnSegment" & vbCrLf

        For s = 0 To SegmentCount - 1
            If CompSegment(s).IsLine Then
                'Define a shorter name version for the two vertex of the segment
                SpX = CompSegment(s).StartPoint.Easting
                SpY = CompSegment(s).StartPoint.Northing
                EpX = CompSegment(s).EndPoint.Easting
                EpY = CompSegment(s).EndPoint.Northing
                'Calculate the intersection point
                k = ((EpY - SpY) * (dataPointX - SpX) - (EpX - SpX) * (dataPointY - SpY)) / ((EpY - SpY) ^ 2 + (EpX - SpX) ^ 2)
                ResultPoints(s).X = dataPointX - k * (EpY - SpY)
                ResultPoints(s).Y = dataPointY + k * (EpX - SpX)
                'Calculate the KP of the Projected point
                TmpKP = Distance2D(SpX, SpY, ResultPoints(s).X, ResultPoints(s).Y)
                'Calculate the DCC
                TmpDCC = Distance2D(ResultPoints(s).X, ResultPoints(s).Y, dataPointX, dataPointY)
                'Calculate the distances to find the projected point position
                D0 = CompSegment(s).Length
                D1 = Distance2D(ResultPoints(s).X, ResultPoints(s).Y, SpX, SpY)
                D2 = Distance2D(ResultPoints(s).X, ResultPoints(s).Y, EpX, EpY)
                'If the sum of the distances of the projpoint from the two vertex is equal to the segment lenght then
                ' the projpoint is between the two vertex.
                DeltaDistance = Math.Abs(D0 - (D1 + D2))
                If DeltaDistance < DistanceTollerance Then
                    ResultPoints(s).OnSegment = PointPosition.Between
                    ResultPoints(s).KP = CompSegment(s).StartPoint.ForwardGridKP + TmpKP
                Else
                    If D1 > D2 Then 'Other segments
                        If s = SegmentCount - 1 Then 'Last segment
                            ResultPoints(s).OnSegment = PointPosition.Between
                            ResultPoints(s).KP = CompSegment(s).StartPoint.ForwardGridKP + TmpKP
                        Else
                            ResultPoints(s).OnSegment = PointPosition.After
                            ResultPoints(s).KP = CompSegment(s).StartPoint.ForwardGridKP + TmpKP
                        End If
                    Else
                        If s = 0 Then 'First segment
                            ResultPoints(s).OnSegment = PointPosition.Between
                            ResultPoints(s).KP = CompSegment(s).StartPoint.ForwardGridKP - TmpKP
                        Else
                            ResultPoints(s).OnSegment = PointPosition.Before
                            ResultPoints(s).KP = CompSegment(s).StartPoint.ForwardGridKP - TmpKP
                        End If
                    End If
                End If
                'Found on witch side is the data point
                If TmpDCC = 0 Then
                    'Point on line
                    ResultPoints(s).DCC = TmpDCC
                Else
                    HdgLine = Math.Atan2(SpY - EpY, SpX - EpX)
                    HdgPrj = Math.Atan2(ResultPoints(s).Y - dataPointY, ResultPoints(s).X - dataPointX)
                    If HdgLine >= -Math.PI And HdgLine <= -(Math.PI / 2) Then
                        ResultPoints(s).DCC = If(HdgPrj >= 0, TmpDCC, -TmpDCC)
                    ElseIf HdgLine >= -(Math.PI / 2) And HdgLine <= 0 Then
                        ResultPoints(s).DCC = If(HdgPrj <= 0, TmpDCC, -TmpDCC)
                    ElseIf HdgLine >= 0 And HdgLine <= (Math.PI / 2) Then
                        ResultPoints(s).DCC = If(HdgPrj <= 0, TmpDCC, -TmpDCC)
                    Else
                        ResultPoints(s).DCC = If(HdgPrj >= 0, TmpDCC, -TmpDCC)
                    End If

                End If
            Else 'Arc Segment
                'Define a shorter name version for the two vertex of the segment
                SpX = CompSegment(s).StartPoint.Easting
                SpY = CompSegment(s).StartPoint.Northing
                EpX = CompSegment(s).EndPoint.Easting
                EpY = CompSegment(s).EndPoint.Northing
                CpX = CompSegment(s).CenterPoint.Easting
                CpY = CompSegment(s).CenterPoint.Northing
                'Check if the intersection will fall on this arc
                Hdg1 = Math.Atan2(SpY - CpY, SpX - CpX)
                Hdg2 = Math.Atan2(EpY - CpY, EpX - CpX)
                Hdg0 = Math.Atan2(dataPointY - CpY, dataPointX - CpX)
                A0 = Hdg2 - Hdg1 'Segment sweep angle
                A1 = Hdg0 - Hdg1 'Sweep angle between V1 and DataPoint
                A2 = Hdg0 - Hdg2 'Sweep angle between V2 and DataPoint

                If Hdg1 > Math.PI / 2 And Hdg2 < -Math.PI / 2 Then
                    If Hdg0 < -Math.PI / 2 Then
                        DeltaAngle = Math.Abs(Math.Abs(A0) - (Math.Abs(A1) - Math.Abs(A2)))
                    Else
                        DeltaAngle = Math.Abs(Math.Abs(A0) - (Math.Abs(A2) - Math.Abs(A1)))
                    End If
                Else
                    A0 = Hdg2 - Hdg1 'Segment sweep angle
                    A1 = Hdg0 - Hdg1 'Sweep angle between V1 and DataPoint
                    A2 = Hdg0 - Hdg2 'Sweep angle between V2 and DataPoint
                    DeltaAngle = Math.Abs(Math.Abs(A0) - (Math.Abs(A1) + Math.Abs(A2)))

                End If

                If DeltaAngle < AngleTollerance Then
                    'Point perpendicular to this arc
                    'evaluate along an arc
                    SegRad = CompSegment(s).Radious
                    'Calculate the coord of the intersection point
                    If CompSegment(s).IsArcClockwise Then
                        TmpPoint = RangeBearing(CpX, CpY, -SegRad, Hdg0 + Math.PI)
                    Else
                        TmpPoint = RangeBearing(CpX, CpY, SegRad, Hdg0)
                    End If
                    ResultPoints(s).X = TmpPoint.X
                    ResultPoints(s).Y = TmpPoint.Y
                    'Calculate the KP of the Projected point
                    If Hdg0 < -Math.PI / 2 And Hdg1 > Math.PI / 2 Then
                        ResultPoints(s).KP = CompSegment(s).StartPoint.ForwardGridKP + (SegRad * Math.Abs((Hdg0 + Math.PI * 2) - Hdg1))
                    Else
                        ResultPoints(s).KP = CompSegment(s).StartPoint.ForwardGridKP + (SegRad * Math.Abs(Hdg0 - Hdg1))
                    End If

                    'Calculate the DCC
                    ResultPoints(s).DCC = Distance2D(ResultPoints(s).X, ResultPoints(s).Y, dataPointX, dataPointY)
                    DpRadious = Distance2D(CpX, CpY, dataPointX, dataPointY)
                    If CompSegment(s).IsArcClockwise Then
                        If DpRadious > SegRad Then ResultPoints(s).DCC = -ResultPoints(s).DCC
                    Else
                        If DpRadious < SegRad Then ResultPoints(s).DCC = -ResultPoints(s).DCC
                    End If
                    ResultPoints(s).OnSegment = PointPosition.Between
                Else
                    'Point do NOT intersect on this arc
                    If A1 >= 0 And A2 >= 0 Then
                        If CompSegment(s).IsArcClockwise Then
                            ResultPoints(s).OnSegment = PointPosition.After
                            If s = SegmentCount - 1 Then
                                'Last segment of the route and data point after V2
                                'Compute intersection on route extension
                                'Define a dummy segment tangent to the end point of the arc
                                TmpPoint = RangeBearing(EpX, EpY, 1000, AngleFit2Pi(Hdg2 + 90))
                                'Check if the intersection will fall on this segment
                                Hdg0 = Math.Atan2(TmpPoint.Y - EpY, TmpPoint.X - EpX)
                                Hdg1 = Math.Atan2(dataPointY - EpY, dataPointX - EpX)
                                A1 = Hdg1 - Hdg0
                                'Compute intersection on route extension as negative KP
                                k = ((TmpPoint.Y - EpY) * (dataPointX - EpX) - (TmpPoint.X - EpX) * (dataPointY - EpY)) / ((TmpPoint.Y - EpY) ^ 2 + (TmpPoint.X - EpX) ^ 2)
                                ResultPoints(s).X = dataPointX - k * (TmpPoint.Y - EpY)
                                ResultPoints(s).Y = dataPointY + k * (TmpPoint.X - EpX)
                                'Calculate the KP of the Projected point
                                ResultPoints(s).KP = -Distance2D(EpX, EpY, ResultPoints(s).X, ResultPoints(s).Y)
                                'Calculate the DCC
                                ResultPoints(s).DCC = Distance2D(ResultPoints(s).X, ResultPoints(s).Y, dataPointX, dataPointY)
                                If A1 > 0 Then ResultPoints(s).DCC = -ResultPoints(s).DCC
                                ResultPoints(s).OnSegment = PointPosition.Between
                            End If
                        Else
                            ResultPoints(s).OnSegment = PointPosition.Before
                            If s = 0 Then
                                'First segment of the route and data point before V1
                                'Define a dummy segment tangent to the start point of the arc
                                TmpPoint = RangeBearing(SpX, SpY, 1000, AngleFit2Pi(Hdg1 + 90))
                                'Check if the intersection will fall on this segment
                                Hdg0 = Math.Atan2(SpY - TmpPoint.Y, SpX - TmpPoint.X)
                                Hdg1 = Math.Atan2(dataPointY - TmpPoint.Y, dataPointX - TmpPoint.X)
                                A1 = Hdg1 - Hdg0
                                'Compute intersection on route extension as negative KP
                                k = ((SpY - TmpPoint.Y) * (dataPointX - TmpPoint.X) - (SpX - TmpPoint.X) * (dataPointY - TmpPoint.Y)) / ((SpY - TmpPoint.Y) ^ 2 + (SpX - TmpPoint.X) ^ 2)
                                ResultPoints(s).X = dataPointX - k * (SpY - TmpPoint.Y)
                                ResultPoints(s).Y = dataPointY + k * (SpX - TmpPoint.X)
                                'Calculate the KP of the Projected point
                                ResultPoints(s).KP = -Distance2D(SpX, SpY, ResultPoints(s).X, ResultPoints(s).Y)
                                'Calculate the DCC
                                ResultPoints(s).DCC = Distance2D(ResultPoints(s).X, ResultPoints(s).Y, dataPointX, dataPointY)
                                If A1 > 0 Then ResultPoints(s).DCC = -ResultPoints(s).DCC
                                ResultPoints(s).OnSegment = PointPosition.Between
                            End If

                        End If
                    Else
                        If A1 <= 0 And A2 <= 0 Then
                            If CompSegment(s).IsArcClockwise Then
                                ResultPoints(s).OnSegment = PointPosition.Before
                                If s = 0 Then
                                    'First segment of the route and data point before V1
                                    'Define a dummy segment tangent to the start point of the arc
                                    TmpPoint = RangeBearing(SpX, SpY, 1000, AngleFit2Pi(Hdg1 + 90))
                                    'Check if the intersection will fall on this segment
                                    Hdg0 = Math.Atan2(SpY - TmpPoint.Y, SpX - TmpPoint.X)
                                    Hdg1 = Math.Atan2(dataPointY - TmpPoint.Y, dataPointX - TmpPoint.X)
                                    A1 = Hdg1 - Hdg0
                                    'Compute intersection on route extension as negative KP
                                    k = ((SpY - TmpPoint.Y) * (dataPointX - TmpPoint.X) - (SpX - TmpPoint.X) * (dataPointY - TmpPoint.Y)) / ((SpY - TmpPoint.Y) ^ 2 + (SpX - TmpPoint.X) ^ 2)
                                    ResultPoints(s).X = dataPointX - k * (SpY - TmpPoint.Y)
                                    ResultPoints(s).Y = dataPointY + k * (SpX - TmpPoint.X)
                                    'Calculate the KP of the Projected point
                                    ResultPoints(s).KP = -Distance2D(SpX, SpY, ResultPoints(s).X, ResultPoints(s).Y)
                                    'Calculate the DCC
                                    ResultPoints(s).DCC = Distance2D(ResultPoints(s).X, ResultPoints(s).Y, dataPointX, dataPointY)
                                    If A1 > 0 Then ResultPoints(s).DCC = -ResultPoints(s).DCC
                                    ResultPoints(s).OnSegment = PointPosition.Between
                                End If
                            Else
                                ResultPoints(s).OnSegment = PointPosition.After
                                If s = SegmentCount - 1 Then
                                    'Last segment of the route and data point after V2
                                    'Compute intersection on route extension
                                    'Define a dummy segment tangent to the end point of the arc
                                    TmpPoint = RangeBearing(EpX, EpY, 1000, AngleFit2Pi(Hdg2 + 90))
                                    'Check if the intersection will fall on this segment
                                    Hdg0 = Math.Atan2(TmpPoint.Y - EpY, TmpPoint.X - EpX)
                                    Hdg1 = Math.Atan2(dataPointY - EpY, dataPointX - EpX)
                                    A1 = Hdg1 - Hdg0
                                    'Compute intersection on route extension as negative KP
                                    k = ((TmpPoint.Y - EpY) * (dataPointX - EpX) - (TmpPoint.X - EpX) * (dataPointY - EpY)) / ((TmpPoint.Y - EpY) ^ 2 + (TmpPoint.X - EpX) ^ 2)
                                    ResultPoints(s).X = dataPointX - k * (TmpPoint.Y - EpY)
                                    ResultPoints(s).Y = dataPointY + k * (TmpPoint.X - EpX)
                                    'Calculate the KP of the Projected point
                                    ResultPoints(s).KP = -Distance2D(EpX, EpY, ResultPoints(s).X, ResultPoints(s).Y)
                                    'Calculate the DCC
                                    ResultPoints(s).DCC = Distance2D(ResultPoints(s).X, ResultPoints(s).Y, dataPointX, dataPointY)
                                    If A1 > 0 Then ResultPoints(s).DCC = -ResultPoints(s).DCC
                                    ResultPoints(s).OnSegment = PointPosition.Between
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            dbgString = dbgString & SpX & "," & SpY & "," & EpX & "," & EpY & "," & CpX & "," & CpY & "," & Hdg0 & "," & Hdg1 & "," & Hdg2 & ","
            dbgString = dbgString & A0 & "," & A1 & "," & A2 & "," & SegRad & "," & DpRadious & "," & s & "," & ResultPoints(s).OnSegment & vbCrLf
        Next

        'Retrieve the minimum DCC value OnSegment
        TmpDCC = Double.PositiveInfinity
        Dim TmpSeg As Integer = -1
        For r = 0 To SegmentCount - 1
            If ResultPoints(r).OnSegment = PointPosition.Between Then
                If TmpDCC > Math.Abs(ResultPoints(r).DCC) Then
                    TmpDCC = Math.Abs(ResultPoints(r).DCC)
                    TmpSeg = r
                End If
            End If
        Next

        'Return the results with the minimum DCC
        If TmpSeg > -1 Then
            projectedPointX = ResultPoints(TmpSeg).X
            projectedPointY = ResultPoints(TmpSeg).Y
            KP = ResultPoints(TmpSeg).KP
            DCC = ResultPoints(TmpSeg).DCC
        Else
            'Calculate the distances from all the vertex
            TmpDist2 = Double.PositiveInfinity
            For r = 0 To SegmentCount - 1
                TmpDist1 = Distance2D(CompSegment(r).StartPoint.Easting, CompSegment(r).StartPoint.Northing, dataPointX, dataPointY)
                If TmpDist1 < TmpDist2 Then
                    TmpDist2 = TmpDist1
                    TmpSeg = r
                End If
            Next
            TmpDist1 = Distance2D(CompSegment(SegmentCount - 1).EndPoint.Easting, CompSegment(SegmentCount - 1).EndPoint.Northing, dataPointX, dataPointY)
            If TmpDist1 < TmpDist2 Then
                TmpDist2 = TmpDist1
                TmpSeg = SegmentCount
            End If
            'Return the info relative to the nearest TP
            If TmpSeg = SegmentCount Then
                projectedPointX = CompSegment(SegmentCount - 1).EndPoint.Easting
                projectedPointY = CompSegment(SegmentCount - 1).EndPoint.Northing
                KP = CompSegment(SegmentCount - 1).StartPoint.ForwardGridKP
                DCC = TmpDist2
            Else
                projectedPointX = CompSegment(TmpSeg).StartPoint.Easting
                projectedPointY = CompSegment(TmpSeg).StartPoint.Northing
                KP = CompSegment(TmpSeg).StartPoint.ForwardGridKP
                DCC = TmpDist2
            End If
        End If

        dbgString = dbgString & dataPointX & "," & dataPointY & "," & projectedPointX & "," & projectedPointY & "," & KP & "," & DCC & "," & TmpSeg & vbCrLf

        Return dbgString

    End Function

    Private Shared Function ComputeEllipsoidDistances(RefRoute As Route, CRS As Projections) As Route
        Dim inPoint, outPoint As New Point3D
        Dim GeoPnt, GeoPnt2 As New Point3D
        Dim tmpEllDist, ArcReduction, TpEll As Double
        Dim nearestTPid As Integer
        Dim nearestTPkp As Double
        Dim MaxGridPointsID As Integer

        '------------------------------------------------------------------------------------------------------------------
        ' First step: Calculate the geographic coordinates of the TPs, CPs and IPs
        '------------------------------------------------------------------------------------------------------------------
        For Each seg In RefRoute.ComputedSegments
            inPoint.X = seg.StartPoint.Easting
            inPoint.Y = seg.StartPoint.Northing
            outPoint = CRS.ToGeographic(inPoint)
            seg.StartPoint.LLfromPoint3D(outPoint)

            inPoint.X = seg.EndPoint.Easting
            inPoint.Y = seg.EndPoint.Northing
            outPoint = CRS.ToGeographic(inPoint)
            seg.EndPoint.LLfromPoint3D(outPoint)

            inPoint.X = seg.CenterPoint.Easting
            inPoint.Y = seg.CenterPoint.Northing
            outPoint = CRS.ToGeographic(inPoint)
            seg.CenterPoint.LLfromPoint3D(outPoint)

            inPoint.X = seg.VertexPoint.Easting
            inPoint.Y = seg.VertexPoint.Northing
            outPoint = CRS.ToGeographic(inPoint)
            seg.VertexPoint.LLfromPoint3D(outPoint)

        Next

        '------------------------------------------------------------------------------------------------------------------
        ' Second step: Calculate the geographic coordinates of the point inside the GFKP array
        '------------------------------------------------------------------------------------------------------------------
        For pn = 0 To RefRoute.ForwardGridPoints.Count - 1
            inPoint.X = RefRoute.ForwardGridPoints(pn).Easting
            inPoint.Y = RefRoute.ForwardGridPoints(pn).Northing
            outPoint = CRS.ToGeographic(inPoint)
            RefRoute.ForwardGridPoints(pn).LLfromPoint3D(outPoint)
        Next

        '------------------------------------------------------------------------------------------------------------------
        ' Fourth step: Calculate the ellipsoidical distances of the points inside the GFKP array
        '------------------------------------------------------------------------------------------------------------------
        nearestTPid = 0
        nearestTPkp = RefRoute.ComputedSegments(nearestTPid).EndPoint.ForwardGridKP
        MaxGridPointsID = RefRoute.ForwardGridPoints.Count - 1
        RefRoute.ForwardGridPoints(0).ForwardEllipsoidKP = 0.0
        RefRoute.ComputedSegments(0).StartPoint.ForwardEllipsoidKP = 0.0

        For epnt = 0 To MaxGridPointsID - 1
            If RefRoute.ComputedSegments(nearestTPid).IsLine Then
                'Calculate distances along a straigth line
                If nearestTPkp - epnt < 1 Then
                    'TP nearer then the next point
                    GeoPnt.X = RefRoute.ComputedSegments(nearestTPid).StartPoint.Longitude
                    GeoPnt.Y = RefRoute.ComputedSegments(nearestTPid).StartPoint.Latitude
                    GeoPnt.Z = 0
                    GeoPnt2.X = RefRoute.ComputedSegments(nearestTPid).EndPoint.Longitude
                    GeoPnt2.Y = RefRoute.ComputedSegments(nearestTPid).EndPoint.Latitude
                    GeoPnt2.Z = 0
                    'TpEll = CRS.BaseEllipsoid.Distance(GeoPnt, GeoPnt2)
                    TpEll = CRS.BaseEllipsoid.DistanceKarney(GeoPnt, GeoPnt2)
                    If TpEll = 0 Then
                        inPoint.X = RefRoute.ForwardGridPoints(epnt).Easting
                        inPoint.Y = RefRoute.ForwardGridPoints(epnt).Northing
                        outPoint.X = RefRoute.ComputedSegments(nearestTPid).EndPoint.Easting
                        outPoint.Y = RefRoute.ComputedSegments(nearestTPid).EndPoint.Northing
                        TpEll = RefRoute.ForwardGridPoints(epnt).ForwardEllipsoidKP + Distance2D(inPoint, outPoint)
                    Else
                        TpEll = TpEll + RefRoute.ComputedSegments(nearestTPid).StartPoint.ForwardEllipsoidKP
                    End If
                    'Add the TP ellKP to the array
                    RefRoute.ComputedSegments(nearestTPid).EndPoint.ForwardEllipsoidKP = TpEll
                    If nearestTPid < RefRoute.SegmentCount - 1 Then
                        RefRoute.ComputedSegments(nearestTPid + 1).StartPoint.ForwardEllipsoidKP = TpEll
                    End If
                    'Calculate the remaining part of the segment
                    GeoPnt.X = RefRoute.ComputedSegments(nearestTPid).EndPoint.Longitude
                    GeoPnt.Y = RefRoute.ComputedSegments(nearestTPid).EndPoint.Latitude
                    GeoPnt.Z = 0
                    GeoPnt2.X = RefRoute.ForwardGridPoints(epnt + 1).Longitude
                    GeoPnt2.Y = RefRoute.ForwardGridPoints(epnt + 1).Latitude
                    GeoPnt2.Z = 0
                    'tmpEllDist = CRS.BaseEllipsoid.Distance(GeoPnt, GeoPnt2)
                    tmpEllDist = CRS.BaseEllipsoid.DistanceKarney(GeoPnt, GeoPnt2)
                    If tmpEllDist = 0 Then
                        inPoint.X = RefRoute.ForwardGridPoints(epnt + 1).Easting
                        inPoint.Y = RefRoute.ForwardGridPoints(epnt + 1).Northing
                        outPoint.X = RefRoute.ComputedSegments(nearestTPid).EndPoint.Easting
                        outPoint.Y = RefRoute.ComputedSegments(nearestTPid).EndPoint.Northing
                        tmpEllDist = TpEll + Distance2D(inPoint, outPoint)
                    Else
                        tmpEllDist = TpEll + tmpEllDist
                    End If
                    'Update nearest TP info
                    If nearestTPid < RefRoute.SegmentCount - 1 Then
                        nearestTPid = nearestTPid + 1
                        nearestTPkp = RefRoute.ComputedSegments(nearestTPid).EndPoint.ForwardGridKP
                    End If
                Else
                    'Calculate the distance between start TP and actual point
                    GeoPnt.X = RefRoute.ComputedSegments(nearestTPid).StartPoint.Longitude
                    GeoPnt.Y = RefRoute.ComputedSegments(nearestTPid).StartPoint.Latitude
                    GeoPnt.Z = 0
                    GeoPnt2.X = RefRoute.ForwardGridPoints(epnt + 1).Longitude
                    GeoPnt2.Y = RefRoute.ForwardGridPoints(epnt + 1).Latitude
                    GeoPnt2.Z = 0
                    'tmpEllDist = CRS.BaseEllipsoid.Distance(GeoPnt, GeoPnt2)
                    tmpEllDist = CRS.BaseEllipsoid.DistanceKarney(GeoPnt, GeoPnt2)
                    If tmpEllDist = 0 Then
                        inPoint.X = RefRoute.ForwardGridPoints(epnt).Easting
                        inPoint.Y = RefRoute.ForwardGridPoints(epnt).Northing
                        outPoint.X = RefRoute.ForwardGridPoints(epnt + 1).Easting
                        outPoint.Y = RefRoute.ForwardGridPoints(epnt + 1).Northing
                        tmpEllDist = RefRoute.ForwardGridPoints(epnt).ForwardEllipsoidKP + Distance2D(inPoint, outPoint)
                    Else
                        tmpEllDist = tmpEllDist + RefRoute.ComputedSegments(nearestTPid).StartPoint.ForwardEllipsoidKP
                    End If
                End If
                'Store the ellkp inside the RefRoute.ForwardGridPoints points array
                RefRoute.ForwardGridPoints(epnt + 1).ForwardEllipsoidKP = tmpEllDist
            Else
                'Calculate distances along an arc segment (Sum elldist for each meter)
                If nearestTPkp - epnt < 1 Then
                    'Last point of this segment
                    GeoPnt.X = RefRoute.ForwardGridPoints(epnt).Longitude
                    GeoPnt.Y = RefRoute.ForwardGridPoints(epnt).Latitude
                    GeoPnt.Z = 0
                    GeoPnt2.X = RefRoute.ComputedSegments(nearestTPid).EndPoint.Longitude
                    GeoPnt2.Y = RefRoute.ComputedSegments(nearestTPid).EndPoint.Latitude
                    GeoPnt2.Z = 0
                    'TpEll = CRS.BaseEllipsoid.Distance(GeoPnt, GeoPnt2)
                    TpEll = CRS.BaseEllipsoid.DistanceKarney(GeoPnt, GeoPnt2)
                    If TpEll = 0 Then
                        inPoint.X = RefRoute.ForwardGridPoints(epnt).Easting
                        inPoint.Y = RefRoute.ForwardGridPoints(epnt).Northing
                        outPoint.X = RefRoute.ComputedSegments(nearestTPid).EndPoint.Easting
                        outPoint.Y = RefRoute.ComputedSegments(nearestTPid).EndPoint.Northing
                        TpEll = RefRoute.ForwardGridPoints(epnt).ForwardEllipsoidKP + Distance2D(inPoint, outPoint)
                    Else
                        TpEll = RefRoute.ForwardGridPoints(epnt).ForwardEllipsoidKP + TpEll
                    End If
                    'Add the TP ellKP to the array
                    RefRoute.ComputedSegments(nearestTPid).EndPoint.ForwardEllipsoidKP = TpEll
                    If nearestTPid < RefRoute.SegmentCount - 1 Then
                        RefRoute.ComputedSegments(nearestTPid + 1).StartPoint.ForwardEllipsoidKP = TpEll
                    End If
                    'Calculate the remaining part of the segment
                    GeoPnt.X = RefRoute.ComputedSegments(nearestTPid).EndPoint.Longitude
                    GeoPnt.Y = RefRoute.ComputedSegments(nearestTPid).EndPoint.Latitude
                    GeoPnt.Z = 0
                    GeoPnt2.X = RefRoute.ForwardGridPoints(epnt + 1).Longitude
                    GeoPnt2.Y = RefRoute.ForwardGridPoints(epnt + 1).Latitude
                    GeoPnt2.Z = 0
                    'tmpEllDist = CRS.BaseEllipsoid.Distance(GeoPnt, GeoPnt2)
                    tmpEllDist = CRS.BaseEllipsoid.DistanceKarney(GeoPnt, GeoPnt2)
                    If tmpEllDist = 0 Then
                        inPoint.X = RefRoute.ForwardGridPoints(epnt + 1).Easting
                        inPoint.Y = RefRoute.ForwardGridPoints(epnt + 1).Northing
                        outPoint.X = RefRoute.ComputedSegments(nearestTPid).EndPoint.Easting
                        outPoint.Y = RefRoute.ComputedSegments(nearestTPid).EndPoint.Northing
                        tmpEllDist = TpEll + Distance2D(inPoint, outPoint)
                    Else
                        tmpEllDist = TpEll + tmpEllDist
                    End If
                    'Store the ellkp inside the RefRoute.ForwardGridPoints points array
                    RefRoute.ForwardGridPoints(epnt + 1).ForwardEllipsoidKP = tmpEllDist
                    'Update nearest TP info
                    If nearestTPid < RefRoute.SegmentCount - 1 Then
                        nearestTPid = nearestTPid + 1
                        nearestTPkp = RefRoute.ComputedSegments(nearestTPid).EndPoint.ForwardGridKP
                    End If
                Else
                    'Calculate the distance between two KP points
                    GeoPnt.X = RefRoute.ForwardGridPoints(epnt).Longitude
                    GeoPnt.Y = RefRoute.ForwardGridPoints(epnt).Latitude
                    GeoPnt.Z = 0
                    GeoPnt2.X = RefRoute.ForwardGridPoints(epnt + 1).Longitude
                    GeoPnt2.Y = RefRoute.ForwardGridPoints(epnt + 1).Latitude
                    GeoPnt2.Z = 0
                    'tmpEllDist = CRS.BaseEllipsoid.Distance(GeoPnt, GeoPnt2)
                    tmpEllDist = CRS.BaseEllipsoid.DistanceKarney(GeoPnt, GeoPnt2)
                    'Calculate the difference between the arc and the chord
                    inPoint.X = RefRoute.ForwardGridPoints(epnt).Easting
                    inPoint.Y = RefRoute.ForwardGridPoints(epnt).Northing
                    outPoint.X = RefRoute.ForwardGridPoints(epnt + 1).Easting
                    outPoint.Y = RefRoute.ForwardGridPoints(epnt + 1).Northing
                    ArcReduction = 1.0 - Distance2D(inPoint, outPoint)
                    tmpEllDist = tmpEllDist + (ArcReduction * tmpEllDist)
                    'Store the ellkp inside the RefRoute.ForwardGridPoints points array
                    RefRoute.ForwardGridPoints(epnt + 1).ForwardEllipsoidKP = RefRoute.ForwardGridPoints(epnt).ForwardEllipsoidKP + tmpEllDist
                End If
            End If
        Next
        'Compute the last point
        GeoPnt.X = RefRoute.ForwardGridPoints(MaxGridPointsID - 1).Longitude
        GeoPnt.Y = RefRoute.ForwardGridPoints(MaxGridPointsID - 1).Latitude
        GeoPnt.Z = 0
        GeoPnt2.X = RefRoute.ForwardGridPoints(MaxGridPointsID).Longitude
        GeoPnt2.Y = RefRoute.ForwardGridPoints(MaxGridPointsID).Latitude
        GeoPnt2.Z = 0
        'tmpEllDist = CRS.BaseEllipsoid.Distance(GeoPnt, GeoPnt2)
        tmpEllDist = CRS.BaseEllipsoid.DistanceKarney(GeoPnt, GeoPnt2)
        If tmpEllDist = 0 Then
            inPoint.X = RefRoute.ForwardGridPoints(MaxGridPointsID).Easting
            inPoint.Y = RefRoute.ForwardGridPoints(MaxGridPointsID).Northing
            outPoint.X = RefRoute.ForwardGridPoints(MaxGridPointsID - 1).Easting
            outPoint.Y = RefRoute.ForwardGridPoints(MaxGridPointsID - 1).Northing
            RefRoute.ForwardGridPoints(MaxGridPointsID).ForwardEllipsoidKP = RefRoute.ForwardGridPoints(MaxGridPointsID - 1).ForwardEllipsoidKP + Distance2D(inPoint, outPoint)
        Else
            RefRoute.ForwardGridPoints(MaxGridPointsID).ForwardEllipsoidKP = RefRoute.ForwardGridPoints(MaxGridPointsID - 1).ForwardEllipsoidKP + tmpEllDist
        End If
        'Calculate and store also the Rev ellkp
        For epnt = 0 To MaxGridPointsID
            RefRoute.ForwardGridPoints(epnt).ReverseEllipsoidKP = RefRoute.ForwardGridPoints(MaxGridPointsID).ForwardEllipsoidKP - RefRoute.ForwardGridPoints(epnt).ForwardEllipsoidKP
        Next

        '------------------------------------------------------------------------------------------------------------------
        ' Sixth step: Calculate the ellipsoidical KP of the TP points
        '------------------------------------------------------------------------------------------------------------------
        Dim ekp1 As Double
        Dim ekp2 As Double
        Dim ekp0 As Double
        'Calculate the ellkp for the TPs
        RefRoute.ComputedSegments(0).StartPoint.ForwardEllipsoidKP = RefRoute.ForwardGridPoints(0).ForwardEllipsoidKP
        RefRoute.ComputedSegments(0).StartPoint.ReverseEllipsoidKP = RefRoute.ForwardGridPoints(0).ReverseEllipsoidKP
        For t = 1 To RefRoute.SegmentCount - 1
            'Forward Ell KP
            ekp1 = RefRoute.ForwardGridPoints(CInt(Math.Truncate(RefRoute.ComputedSegments(t).StartPoint.ForwardGridKP))).ForwardEllipsoidKP
            ekp2 = RefRoute.ForwardGridPoints(CInt(Math.Truncate(RefRoute.ComputedSegments(t).StartPoint.ForwardGridKP)) + 1).ForwardEllipsoidKP
            ekp0 = ((ekp2 - ekp1) * (RefRoute.ComputedSegments(t).StartPoint.ForwardGridKP - RefRoute.ForwardGridPoints(CInt(Math.Truncate(RefRoute.ComputedSegments(t).StartPoint.ForwardGridKP))).ForwardGridKP)) + ekp1
            RefRoute.ComputedSegments(t).StartPoint.ForwardEllipsoidKP = ekp0
            RefRoute.ComputedSegments(t - 1).EndPoint.ForwardEllipsoidKP = ekp0
            'Reverse Ell KP
            ekp1 = RefRoute.ForwardGridPoints(CInt(Math.Truncate(RefRoute.ComputedSegments(t).StartPoint.ForwardGridKP))).ReverseEllipsoidKP
            ekp2 = RefRoute.ForwardGridPoints(CInt(Math.Truncate(RefRoute.ComputedSegments(t).StartPoint.ForwardGridKP)) + 1).ReverseEllipsoidKP
            ekp0 = ((ekp2 - ekp1) * (RefRoute.ComputedSegments(t).StartPoint.ForwardGridKP - RefRoute.ForwardGridPoints(CInt(Math.Truncate(RefRoute.ComputedSegments(t).StartPoint.ForwardGridKP))).ForwardGridKP)) + ekp1
            RefRoute.ComputedSegments(t).StartPoint.ReverseEllipsoidKP = ekp0
            RefRoute.ComputedSegments(t - 1).EndPoint.ReverseEllipsoidKP = ekp0
        Next
        RefRoute.ComputedSegments(RefRoute.SegmentCount - 1).EndPoint.ForwardEllipsoidKP = RefRoute.ForwardGridPoints(MaxGridPointsID).ForwardEllipsoidKP
        RefRoute.ComputedSegments(RefRoute.SegmentCount - 1).EndPoint.ReverseEllipsoidKP = RefRoute.ForwardGridPoints(MaxGridPointsID).ReverseEllipsoidKP

        'Calculate the ellkp for the CP and VP
        For Each seg In RefRoute.ComputedSegments
            seg.CenterPoint.ForwardEllipsoidKP = seg.StartPoint.ForwardEllipsoidKP + (seg.EndPoint.ForwardEllipsoidKP - seg.StartPoint.ForwardEllipsoidKP) / 2
            seg.VertexPoint.ForwardEllipsoidKP = seg.StartPoint.ForwardEllipsoidKP + (seg.EndPoint.ForwardEllipsoidKP - seg.StartPoint.ForwardEllipsoidKP) / 2
            seg.CenterPoint.ReverseEllipsoidKP = seg.StartPoint.ReverseEllipsoidKP + (seg.EndPoint.ReverseEllipsoidKP - seg.StartPoint.ReverseEllipsoidKP) / 2
            seg.VertexPoint.ReverseEllipsoidKP = seg.StartPoint.ReverseEllipsoidKP + (seg.EndPoint.ReverseEllipsoidKP - seg.StartPoint.ReverseEllipsoidKP) / 2
        Next

        Return RefRoute

    End Function

    Public Function GetRoutePointFromProgressiveDistance(distance As Double, inputType As DistanceType) As RoutePoint
        Dim resPoint As New RoutePoint
        Dim ratio As Double
        Dim resID As Integer = -1

        If IsFinite(distance) Then

            Select Case inputType
                Case DistanceType.ForwardGrid
                    If distance <= 0 Then
                        'Point outside of the computed route
                        resID = -1111
                    ElseIf distance >= iFwdGrdKps(iFwdGrdKps.Count - 1).ForwardGridKP Then
                        'Point outside of the computed route
                        resID = -9999
                    Else
                        For p = 0 To iFwdGrdKps.Count - 2
                            If distance <= iFwdGrdKps(p + 1).ForwardGridKP And distance > iFwdGrdKps(p).ForwardGridKP Then
                                ratio = (distance - iFwdGrdKps(p).ForwardGridKP) / (iFwdGrdKps(p + 1).ForwardGridKP - iFwdGrdKps(p).ForwardGridKP)
                                resID = p
                                Exit For
                            End If
                        Next
                    End If

                Case DistanceType.ReverseGrid
                    If distance <= 0 Then
                        'Point outside of the computed route
                        resID = -1111
                    ElseIf distance >= iFwdGrdKps(0).ReverseGridKP Then
                        'Point outside of the computed route
                        resID = -9999
                    Else
                        For p = 0 To iFwdGrdKps.Count - 2
                            If distance >= iFwdGrdKps(p + 1).ReverseGridKP And distance < iFwdGrdKps(p).ReverseGridKP Then
                                ratio = (distance - iFwdGrdKps(p).ReverseGridKP) / (iFwdGrdKps(p + 1).ReverseGridKP - iFwdGrdKps(p).ReverseGridKP)
                                resID = p
                                Exit For
                            End If
                        Next
                    End If

                Case DistanceType.ForwardEllipsoidical
                    If distance <= 0 Then
                        'Point outside of the computed route
                        resID = -1111
                    ElseIf distance >= iFwdGrdKps(iFwdGrdKps.Count - 1).ForwardEllipsoidKP Then
                        'Point outside of the computed route
                        resID = -9999
                    Else
                        For p = 0 To iFwdGrdKps.Count - 2
                            If distance <= iFwdGrdKps(p + 1).ForwardEllipsoidKP And distance > iFwdGrdKps(p).ForwardEllipsoidKP Then
                                ratio = (distance - iFwdGrdKps(p).ForwardEllipsoidKP) / (iFwdGrdKps(p + 1).ForwardEllipsoidKP - iFwdGrdKps(p).ForwardEllipsoidKP)
                                resID = p
                                Exit For
                            End If
                        Next
                    End If

                Case DistanceType.ReverseEllipsoidical
                    If distance <= 0 Then
                        'Point outside of the computed route
                        resID = -1111
                    ElseIf distance >= iFwdGrdKps(0).ReverseEllipsoidKP Then
                        'Point outside of the computed route
                        resID = -9999
                    Else
                        For p = 0 To iFwdGrdKps.Count - 2
                            If distance >= iFwdGrdKps(p + 1).ReverseEllipsoidKP And distance < iFwdGrdKps(p).ReverseEllipsoidKP Then
                                ratio = (distance - iFwdGrdKps(p).ReverseEllipsoidKP) / (iFwdGrdKps(p + 1).ReverseEllipsoidKP - iFwdGrdKps(p).ReverseEllipsoidKP)
                                resID = p
                                Exit For
                            End If
                        Next
                    End If

            End Select

            If resID > -1 Then
                resPoint.Easting = ((iFwdGrdKps(resID + 1).Easting - iFwdGrdKps(resID).Easting) * ratio) + iFwdGrdKps(resID).Easting
                resPoint.Northing = ((iFwdGrdKps(resID + 1).Northing - iFwdGrdKps(resID).Northing) * ratio) + iFwdGrdKps(resID).Northing
                resPoint.Latitude = ((iFwdGrdKps(resID + 1).Latitude - iFwdGrdKps(resID).Latitude) * ratio) + iFwdGrdKps(resID).Latitude
                resPoint.Longitude = ((iFwdGrdKps(resID + 1).Longitude - iFwdGrdKps(resID).Longitude) * ratio) + iFwdGrdKps(resID).Longitude
                resPoint.ForwardGridKP = ((iFwdGrdKps(resID + 1).ForwardGridKP - iFwdGrdKps(resID).ForwardGridKP) * ratio) + iFwdGrdKps(resID).ForwardGridKP
                resPoint.ReverseGridKP = ((iFwdGrdKps(resID + 1).ReverseGridKP - iFwdGrdKps(resID).ReverseGridKP) * ratio) + iFwdGrdKps(resID).ReverseGridKP
                resPoint.ForwardEllipsoidKP = ((iFwdGrdKps(resID + 1).ForwardEllipsoidKP - iFwdGrdKps(resID).ForwardEllipsoidKP) * ratio) + iFwdGrdKps(resID).ForwardEllipsoidKP
                resPoint.ReverseEllipsoidKP = ((iFwdGrdKps(resID + 1).ReverseEllipsoidKP - iFwdGrdKps(resID).ReverseEllipsoidKP) * ratio) + iFwdGrdKps(resID).ReverseEllipsoidKP
                resPoint.Heading = CalcHeading(iFwdGrdKps(resID).Easting, iFwdGrdKps(resID).Northing, iFwdGrdKps(resID + 1).Easting, iFwdGrdKps(resID + 1).Northing)
                Return resPoint

            ElseIf resID = -1111 Then
                Return iFwdGrdKps(0)

            ElseIf resID = -9999 Then
                Return iFwdGrdKps(iFwdGrdKps.Count - 1)
            Else
                Return New RoutePoint

            End If
        Else
            Return New RoutePoint
        End If

    End Function

    Private Shared Function IsArcClockwise(P1 As Point3D, V As Point3D, P2 As Point3D) As Boolean
        Dim HdgS1, HdgS2 As Double
        HdgS1 = CalcHeading(P1, V)
        HdgS2 = CalcHeading(V, P2)

        If HdgS1 - HdgS2 < 0 Then
            Return False
        Else
            Return True
        End If

    End Function

End Class