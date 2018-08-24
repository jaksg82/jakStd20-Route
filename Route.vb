Imports System.Math
Imports jakStd20_MathExt
Imports System.IO
Imports jakStd20_Datums

<Assembly: CLSCompliant(False)>
<Assembly: System.Runtime.InteropServices.ComVisible(False)>

Public Class Route
    Dim IsLoaded As LoadingStatus
    Dim CompSegment As New List(Of RouteSegment)
    Dim iRouteMin, iRouteMax As New Point3D

    'Dim iFwdGrdKps, iRevGrdKps, iFwdEllKps, iRevEllKps As New List(Of RoutePoint)
    Dim iFwdGrdKps As New List(Of RoutePoint)

    Dim iMastPoints As New List(Of MasterPoint)
    Dim iMastSegs As New List(Of MasterSegment)

    Public Property Name As String

    ReadOnly Property SegmentCount As Integer
        Get
            If CompSegment Is Nothing Then
                Return -1
            Else
                Return CompSegment.Count
            End If
        End Get
    End Property

    ReadOnly Property Lenght As Double
        Get
            If CompSegment Is Nothing Then
                Return 0.0
            Else
                Return CompSegment.Last.EndPoint.ForwardGridKP
            End If
        End Get
    End Property

    ReadOnly Property IsRouteLoaded As LoadingStatus
        Get
            Return IsLoaded
        End Get
    End Property

    ReadOnly Property ComputedSegments As List(Of RouteSegment)
        Get
            Return CompSegment
        End Get
    End Property

    ReadOnly Property RouteMin As Point3D
        Get
            Return iRouteMin
        End Get
    End Property

    ReadOnly Property RouteMax As Point3D
        Get
            Return iRouteMax
        End Get
    End Property

    ReadOnly Property ForwardGridPoints As List(Of RoutePoint)
        Get
            Return iFwdGrdKps
        End Get
    End Property

    Public Enum LoadingStatus
        NotLoaded
        Segments
        GridKPs
        All
    End Enum

    Public Enum DistanceType
        ForwardGrid
        ReverseGrid
        ForwardEllipsoidical
        ReverseEllipsoidical
    End Enum

    Public Function GetMasterPoints() As List(Of MasterPoint)
        Return iMastPoints
    End Function

    Public Function GetMasterSegments() As List(Of MasterSegment)
        Return iMastSegs
    End Function

    Public Sub New()
        IsLoaded = LoadingStatus.NotLoaded
        Name = "Route_" & Date.Now.ToString("yyyyMMddHHmmss", Globalization.CultureInfo.InvariantCulture)

    End Sub

    ''' <summary>
    ''' Create a route from the given master points and segments
    ''' </summary>
    ''' <param name="points">Master points of the route</param>
    ''' <param name="segments">Master segments of the route</param>
    Public Sub New(points As List(Of MasterPoint), segments As List(Of MasterSegment))
        Name = "Route_" & Date.Now.ToString("yyyyMMddHHmmss", Globalization.CultureInfo.InvariantCulture)
        FromMasterRoute(points, segments, False)
    End Sub

    ''' <summary>
    ''' Create a route from the given master points and segments
    ''' </summary>
    ''' <param name="points">Master points of the route</param>
    ''' <param name="segments">Master segments of the route</param>
    Public Sub New(points As List(Of MasterPoint), segments As List(Of MasterSegment), loadOnlySegments As Boolean)
        Name = "Route_" & Date.Now.ToString("yyyyMMddHHmmss", Globalization.CultureInfo.InvariantCulture)
        FromMasterRoute(points, segments, loadOnlySegments)
    End Sub

    ''' <summary>
    ''' Create a route from the given virtual file
    ''' </summary>
    ''' <param name="routeStream">Stream object from where to read</param>
    ''' <remarks></remarks>
    Public Sub New(routeStream As Stream)
        Dim oldline, tmpline As String
        Dim RteList As New List(Of RouteFileEntry)
        Name = "Route_" & Date.Now.ToString("yyyyMMddHHmmss", Globalization.CultureInfo.InvariantCulture)

        Dim NumPoints As Integer = -1
        tmpline = "P,X,Y"
        oldline = tmpline

        'Open the given file to fill the route array
        Using rte As New StreamReader(routeStream)
            Do Until tmpline = "" Or rte.EndOfStream
                'Read a new line
                tmpline = rte.ReadLine
                'Verify if is double point
                If tmpline = oldline Then
                    'Read a new line
                    tmpline = rte.ReadLine
                    If tmpline = oldline Then
                        'Read a new line
                        tmpline = rte.ReadLine
                        If tmpline = oldline Then
                            'Read a new line
                            tmpline = rte.ReadLine
                            If tmpline = oldline Then
                                'Read a new line
                                tmpline = rte.ReadLine
                                If tmpline = oldline Then
                                    'Read a new line
                                    tmpline = rte.ReadLine
                                    oldline = tmpline
                                Else
                                    oldline = tmpline
                                End If
                            Else
                                oldline = tmpline
                            End If
                        Else
                            oldline = tmpline
                        End If
                    Else
                        oldline = tmpline
                    End If
                Else
                    oldline = tmpline
                End If

                NumPoints = NumPoints + 1

                Dim newline As New RouteFileEntry(tmpline)
                If newline.Code <> "" Then
                    RteList.Add(newline)
                End If

            Loop
        End Using
        Dim tmpMaster As New MasterRoute
        tmpMaster = ReadRouteArray(RteList)
        For Each p In tmpMaster.Points
            iMastPoints.Add(p)
        Next
        For Each s In tmpMaster.Segments
            iMastSegs.Add(s)
        Next
        FromMasterRoute(tmpMaster.Points, tmpMaster.Segments)

    End Sub

    ''' <summary>
    ''' Create a route from the given master points and segments
    ''' </summary>
    ''' <param name="Points">Master points of the route</param>
    ''' <param name="Segments">Master segments of the route</param>
    ''' <param name="LoadOnlySegments">True to skip the KP computation</param>
    Private Sub FromMasterRoute(Points As List(Of MasterPoint), Segments As List(Of MasterSegment), Optional LoadOnlySegments As Boolean = False)
        iMastPoints = Points
        iMastSegs = Segments
        CompSegment.Clear()
        If Points.Count < 2 Then
            'Not enough points to create a route
            'Dim tmpSeg As New RouteSegment
            Dim startPoint, endPoint As New RoutePoint()

            If Points.Count = 1 Then
                startPoint.Easting = Points(0).Easting
                startPoint.Northing = Points(0).Northing
                startPoint.Name = Points(0).Name
                startPoint.ForwardGridKP = 0
                endPoint.Easting = Points(0).Easting
                endPoint.Northing = Points(0).Northing + 1
                endPoint.Name = Points(0).Name & "-2"
                endPoint.ForwardGridKP = 1
            Else
                startPoint.Easting = 0
                startPoint.Northing = 0
                startPoint.Name = ""
                startPoint.ForwardGridKP = 0
                endPoint.Easting = 0
                endPoint.Northing = 1
                endPoint.Name = ""
                endPoint.ForwardGridKP = 1
            End If

            CompSegment.Add(ComputeLineSegment(startPoint, endPoint))

        ElseIf Points.Count = 2 Then
            'Single segment route
            Dim tmpSeg As New RouteSegment
            Dim startPoint, endPoint As New RoutePoint()
            startPoint.Easting = Points(0).Easting
            startPoint.Northing = Points(0).Northing
            startPoint.Name = Points(0).Name
            startPoint.ForwardGridKP = 0
            endPoint.Easting = Points(1).Easting
            endPoint.Northing = Points(1).Northing
            endPoint.Name = Points(1).Name
            If Segments(0).IsArc Then
                'Arc
                tmpSeg = ComputeArcSegment(startPoint, endPoint, Segments(0))
            Else
                'Line
                tmpSeg = ComputeLineSegment(startPoint, endPoint)
            End If

            CompSegment.Add(tmpSeg)
        Else
            'Multi segment route
            For p = 1 To Points.Count - 1
                If Points(p).Type = MasterPoint.PointType.VertexPoint Then
                    If Points(p).Radious > 0 Then
                        'Rounded vertex
                        If p = Points.Count - 1 Then
                            'Last point so ignore the radious
                            Dim startPoint As New RoutePoint(Points(p - 1))
                            Dim endPoint As New RoutePoint(Points(p))
                            Dim tmpSeg As New RouteSegment
                            startPoint.ForwardGridKP = If(p = 1, 0.0, CompSegment.Last.EndPoint.ForwardGridKP)
                            tmpSeg = ComputeLineSegment(startPoint, endPoint)
                            CompSegment.Add(tmpSeg)
                        Else
                            Dim startPoint As New RoutePoint(Points(p - 1))
                            Dim vertPoint As New RoutePoint(Points(p))
                            Dim endPoint As New RoutePoint(Points(p + 1))
                            Dim ArcStart, ArcEnd As New RoutePoint
                            Dim tmpSeg1, tmpSeg2 As New RouteSegment
                            Dim tmpArcDef As New MasterSegment
                            Dim tmpRadious, alpha1, alpha2, beta, dist As Double

                            'Check for a previus vertex
                            If Points(p - 1).Type = MasterPoint.PointType.VertexPoint Then
                                If Points(p - 1).Radious > 0 Then
                                    startPoint = CompSegment.Last.EndPoint
                                End If
                            End If

                            'Prepare the points
                            tmpRadious = Points(p).Radious
                            alpha1 = Atan2(startPoint.Northing - vertPoint.Northing, startPoint.Easting - vertPoint.Easting)
                            alpha2 = Atan2(vertPoint.Northing - endPoint.Northing, vertPoint.Easting - endPoint.Easting)
                            beta = PI - Abs(alpha1 - alpha2)
                            'Evaluate the bending direction
                            If (alpha1 - alpha2) < 0 Then
                                'Left Bend
                                tmpArcDef.IsArcClockwise = False
                            Else
                                'Right Bend
                                tmpArcDef.IsArcClockwise = True
                            End If
                            tmpArcDef.IsArc = True
                            tmpArcDef.Radious = tmpRadious
                            'Evaluate the distance between TP and IP
                            dist = Tan((PI / 2) - (beta / 2)) * tmpRadious
                            'Define the start and the end point of the arc segment
                            ArcStart.ENfromPoint3D(RangeBearing(vertPoint.Easting, vertPoint.Northing, dist, AngleFit2Pi(alpha1)))
                            ArcEnd.ENfromPoint3D(RangeBearing(vertPoint.Easting, vertPoint.Northing, dist, AngleFit2Pi(alpha2 - PI)))

                            '------DEFINE THE STRAIGHT SEGMENT BEFORE THE ARC-------
                            startPoint.ForwardGridKP = If(p = 1, 0.0, CompSegment.Last.EndPoint.ForwardGridKP)
                            tmpSeg1 = ComputeLineSegment(startPoint, ArcStart)
                            CompSegment.Add(tmpSeg1)

                            '------DEFINE THE ARC SEGMENT------------------
                            ArcStart.ForwardGridKP = tmpSeg1.EndPoint.ForwardGridKP
                            tmpSeg2 = ComputeArcSegment(ArcStart, ArcEnd, tmpArcDef)
                            CompSegment.Add(tmpSeg2)

                        End If
                    Else
                        'Sharp vertex
                        Dim startPoint As New RoutePoint(Points(p - 1))
                        Dim endPoint As New RoutePoint(Points(p))
                        Dim tmpSeg As New RouteSegment
                        startPoint.ForwardGridKP = If(p = 1, 0.0, CompSegment.Last.EndPoint.ForwardGridKP)
                        tmpSeg = ComputeLineSegment(startPoint, endPoint)
                        CompSegment.Add(tmpSeg)
                    End If
                Else
                    Dim startPoint As New RoutePoint(Points(p - 1))
                    Dim endPoint As New RoutePoint(Points(p))
                    Dim tmpSeg As New RouteSegment
                    startPoint.ForwardGridKP = If(p = 1, 0.0, CompSegment.Last.EndPoint.ForwardGridKP)

                    'Check for arc segment
                    If Points(p - 1).Type = MasterPoint.PointType.VertexPoint Then
                        If CompSegment.Last.IsLine Then
                            'Line segment
                            tmpSeg = ComputeLineSegment(startPoint, endPoint)
                        Else
                            startPoint = CompSegment.Last.EndPoint
                            tmpSeg = ComputeLineSegment(startPoint, endPoint)

                        End If
                    Else
                        If Segments(p - 1).IsArc Then
                            'Arc segment
                            tmpSeg = ComputeArcSegment(startPoint, endPoint, Segments(p - 1))
                        Else
                            'Line segment
                            tmpSeg = ComputeLineSegment(startPoint, endPoint)
                        End If
                    End If
                    CompSegment.Add(tmpSeg)
                End If
            Next
        End If
        SetDataRange(Me)
        SetGridKps(Me)
        IsLoaded = LoadingStatus.Segments
        If LoadOnlySegments = False Then
            'Compute the KP lists
            iFwdGrdKps = CreateGridKpLists(Me)
            If iFwdGrdKps IsNot Nothing And iFwdGrdKps.Count > 0 Then
                IsLoaded = LoadingStatus.GridKPs
                'Dim revRoute As New Route(Me.GetMasterPoints, Me.GetMasterSegments, True)
                'revRoute = GetReversedRoute(revRoute)
                'iRevGrdKps = CreateGridKpLists(revRoute)
                'If iRevGrdKps IsNot Nothing And iRevGrdKps.Count > 0 Then
                '    IsLoaded = LoadingStatus.GridKPs
                'End If

            End If

        End If

    End Sub

    ''' <summary>
    ''' Build the route from the given informations
    ''' </summary>
    ''' <param name="RouteList">Array of the parsed route informations</param>
    ''' <returns>True if builded correctly</returns>
    ''' <remarks></remarks>
    Private Shared Function ReadRouteArray(RouteList As List(Of RouteFileEntry)) As MasterRoute

        Dim tmpMastRoute As New MasterRoute

        'Verify the start point
        If RouteList(0).Code <> "P" And RouteList(0).Code <> "I" Then
            'arc definition
            'ignore the entry
        Else
            Dim tmpPnt As New MasterPoint
            tmpPnt.Easting = RouteList(0).East
            tmpPnt.Northing = RouteList(0).North
            tmpPnt.Radious = RouteList(0).Radious
            tmpPnt.Type = If(RouteList(0).Code = "I", MasterPoint.PointType.VertexPoint, MasterPoint.PointType.TangentPoint)
            tmpMastRoute.Points.Add(tmpPnt)
        End If

        'Verify the sequence
        For p = 1 To RouteList.Count - 2
            If RouteList(p).Code = "A" Then
                If RouteList(p - 1).Code = "P" Then
                    If RouteList(p + 1).Code = "P" Then
                        'correct definition of an arc segment
                        Dim tmpsegarc As New MasterSegment
                        tmpsegarc.IsArc = True
                        tmpsegarc.IsArcClockwise = RouteList(p).IsArcClockwise
                        tmpsegarc.Radious = RouteList(p).Radious
                        tmpMastRoute.Segments.Add(tmpsegarc)
                        Dim tmpPnt As New MasterPoint
                        tmpPnt.Easting = RouteList(p + 1).East
                        tmpPnt.Northing = RouteList(p + 1).North
                        tmpPnt.Radious = RouteList(p + 1).Radious
                        tmpPnt.Type = MasterPoint.PointType.TangentPoint
                        tmpMastRoute.Points.Add(tmpPnt)
                        'skip the entry already computed
                        p = p + 1
                    ElseIf RouteList(p + 1).Code = "I" Then
                        If RouteList(p + 1).Radious <= 0 Or Double.IsNaN(RouteList(p + 1).Radious) Then
                            'end point is not a tangent but is ok
                            Dim tmpsegarc As New MasterSegment
                            tmpsegarc.IsArc = True
                            tmpsegarc.IsArcClockwise = RouteList(p).IsArcClockwise
                            tmpsegarc.Radious = RouteList(p).Radious
                            tmpMastRoute.Segments.Add(tmpsegarc)
                            Dim tmpPnt As New MasterPoint
                            tmpPnt.Easting = RouteList(p + 1).East
                            tmpPnt.Northing = RouteList(p + 1).North
                            tmpPnt.Radious = RouteList(p + 1).Radious
                            tmpPnt.Type = MasterPoint.PointType.TangentPoint
                            tmpMastRoute.Points.Add(tmpPnt)
                            'skip the entry already computed
                            p = p + 1
                        Else
                            'double radious definition
                            'ignore the arc definition and keep the vertex
                            Dim tmpsegarc As New MasterSegment
                            tmpMastRoute.Segments.Add(tmpsegarc)
                            Dim tmpPnt As New MasterPoint
                            tmpPnt.Easting = RouteList(p + 1).East
                            tmpPnt.Northing = RouteList(p + 1).North
                            tmpPnt.Radious = RouteList(p + 1).Radious
                            tmpPnt.Type = MasterPoint.PointType.VertexPoint
                            tmpMastRoute.Points.Add(tmpPnt)
                            'skip the entry already computed
                            p = p + 1

                        End If
                    Else
                        'another arc definition
                        'skip the entry
                        p = p + 1

                    End If

                ElseIf RouteList(p - 1).Code = "I" Then
                    If RouteList(p + 1).Code = "P" Then
                        If RouteList(p - 1).Radious <= 0 Or Double.IsNaN(RouteList(p - 1).Radious) Then
                            'start point is not a tangent but is ok
                            Dim tmpsegarc As New MasterSegment
                            tmpsegarc.IsArc = True
                            tmpsegarc.IsArcClockwise = RouteList(p).IsArcClockwise
                            tmpsegarc.Radious = RouteList(p).Radious
                            tmpMastRoute.Segments.Add(tmpsegarc)
                            Dim tmpPnt As New MasterPoint
                            tmpPnt.Easting = RouteList(p + 1).East
                            tmpPnt.Northing = RouteList(p + 1).North
                            tmpPnt.Radious = RouteList(p + 1).Radious
                            tmpPnt.Type = MasterPoint.PointType.TangentPoint
                            tmpMastRoute.Points.Add(tmpPnt)
                            'skip the entry already computed
                            p = p + 1
                        Else
                            'double radious definition
                            'ignore the arc definition and keep the vertex
                            Dim tmpsegarc As New MasterSegment
                            tmpMastRoute.Segments.Add(tmpsegarc)
                            Dim tmpPnt As New MasterPoint
                            tmpPnt.Easting = RouteList(p + 1).East
                            tmpPnt.Northing = RouteList(p + 1).North
                            tmpPnt.Radious = RouteList(p + 1).Radious
                            tmpPnt.Type = MasterPoint.PointType.VertexPoint
                            tmpMastRoute.Points.Add(tmpPnt)
                            'skip the entry already computed
                            p = p + 1

                        End If

                    ElseIf RouteList(p + 1).Code = "I" Then
                        'double vertex definition
                        If RouteList(p - 1).Radious <= 0 Or Double.IsNaN(RouteList(p - 1).Radious) Then
                            If RouteList(p - 1).Radious <= 0 Or Double.IsNaN(RouteList(p - 1).Radious) Then
                                'both start and end points can be used as tangent
                                Dim tmpsegarc As New MasterSegment
                                tmpsegarc.IsArc = True
                                tmpsegarc.IsArcClockwise = RouteList(p).IsArcClockwise
                                tmpsegarc.Radious = RouteList(p).Radious
                                tmpMastRoute.Segments.Add(tmpsegarc)
                                Dim tmpPnt As New MasterPoint
                                tmpPnt.Easting = RouteList(p + 1).East
                                tmpPnt.Northing = RouteList(p + 1).North
                                tmpPnt.Radious = RouteList(p + 1).Radious
                                tmpPnt.Type = MasterPoint.PointType.TangentPoint
                                tmpMastRoute.Points.Add(tmpPnt)
                                'skip the entry already computed
                                p = p + 1

                            End If
                        Else
                            'skip the entry
                            p = p + 1

                        End If
                    Else
                        'another arc definition
                        'skip the entry
                        p = p + 1

                    End If

                End If
            Else
                'point definition
                Dim tmpsegline As New MasterSegment
                tmpMastRoute.Segments.Add(tmpsegline)
                Dim tmpPnt As New MasterPoint
                tmpPnt.Easting = RouteList(p).East
                tmpPnt.Northing = RouteList(p).North
                tmpPnt.Radious = RouteList(p).Radious
                tmpPnt.Type = If(RouteList(p).Code = "I", MasterPoint.PointType.VertexPoint, MasterPoint.PointType.TangentPoint)
                tmpMastRoute.Points.Add(tmpPnt)

            End If
        Next

        'Verify the end point
        If RouteList(RouteList.Count - 1).Code <> "P" And RouteList(RouteList.Count - 1).Code <> "I" Then
            'last entry is an arc
            'ignore the entry
        Else
            'point definition
            Dim tmpsegline As New MasterSegment
            tmpMastRoute.Segments.Add(tmpsegline)
            Dim tmpPnt As New MasterPoint
            tmpPnt.Easting = RouteList(RouteList.Count - 1).East
            tmpPnt.Northing = RouteList(RouteList.Count - 1).North
            tmpPnt.Radious = RouteList(RouteList.Count - 1).Radious
            tmpPnt.Type = If(RouteList(RouteList.Count - 1).Code = "I", MasterPoint.PointType.VertexPoint, MasterPoint.PointType.TangentPoint)
            tmpMastRoute.Points.Add(tmpPnt)

        End If

        Return tmpMastRoute

    End Function

    Private Shared Function ComputeArcSegment(startPoint As RoutePoint, endPoint As RoutePoint, Definition As MasterSegment) As RouteSegment
        'Arc
        'Verify the existence of a solution with given radious
        Dim chord, sweptangle, tmpKP, seghdg, hdg2center, hdg2vertex As Double
        Dim centerPoint, vertexPoint As New RoutePoint
        Dim tmpSeg As New RouteSegment

        chord = Distance2D(startPoint.Easting, startPoint.Northing, endPoint.Easting, endPoint.Northing)
        'Calculate the lenght of the arc
        sweptangle = 2 * Asin(chord / 2 / Definition.Radious)
        tmpKP = sweptangle * Definition.Radious
        'Calculate the coordinates of the center and vertex points
        seghdg = CalcHeading(startPoint.Easting, startPoint.Northing, endPoint.Easting, endPoint.Northing)
        If Definition.IsArcClockwise Then
            hdg2center = AngleFit2Pi(seghdg - (((2 * PI) - sweptangle) / 2) + (PI / 2))
            hdg2vertex = AngleFit2Pi(hdg2center + (PI / 2))
        Else
            'Left bend need to reverse the angles to draw correct
            hdg2center = AngleFit2Pi(seghdg + ((PI / 2) - (sweptangle / 2)))
            hdg2vertex = AngleFit2Pi(hdg2center - (PI / 2))
        End If
        centerPoint.ENfromPoint3D(RangeBearing(startPoint.Easting, startPoint.Northing, Definition.Radious, hdg2center))
        vertexPoint.ENfromPoint3D(RangeBearing(startPoint.Easting, startPoint.Northing, (Definition.Radious * Tan(sweptangle / 2)), hdg2vertex))

        'Put all together
        endPoint.ForwardGridKP = startPoint.ForwardGridKP + tmpKP
        centerPoint.ForwardGridKP = startPoint.ForwardGridKP + (tmpKP / 2)
        vertexPoint.ForwardGridKP = startPoint.ForwardGridKP + (tmpKP / 2)

        tmpSeg.StartPoint = startPoint
        tmpSeg.EndPoint = endPoint
        tmpSeg.VertexPoint = vertexPoint
        tmpSeg.CenterPoint = centerPoint
        tmpSeg.Length = tmpKP
        tmpSeg.IsLine = False
        tmpSeg.IsArcClockwise = Definition.IsArcClockwise
        tmpSeg.Radious = Definition.Radious

        Return tmpSeg

    End Function

    Private Shared Function ComputeLineSegment(startPoint As RoutePoint, endPoint As RoutePoint) As RouteSegment
        Dim tmpSeg As New RouteSegment
        Dim centerPoint As New RoutePoint

        tmpSeg.Length = Distance2D(startPoint.Easting, startPoint.Northing, endPoint.Easting, endPoint.Northing)
        endPoint.ForwardGridKP = tmpSeg.Length + startPoint.ForwardGridKP
        centerPoint.Easting = (startPoint.Easting + endPoint.Easting) / 2
        centerPoint.Northing = (startPoint.Northing + endPoint.Northing) / 2
        centerPoint.Name = ""
        centerPoint.ForwardGridKP = startPoint.ForwardGridKP + (tmpSeg.Length / 2)
        tmpSeg.StartPoint = startPoint
        tmpSeg.EndPoint = endPoint
        tmpSeg.VertexPoint = centerPoint
        tmpSeg.CenterPoint = centerPoint
        tmpSeg.IsLine = True
        tmpSeg.IsArcClockwise = True
        tmpSeg.Radious = 0

        Return tmpSeg
    End Function

    Private Sub SetDataRange(ByRef RefRoute As Route)
        'Erase any previous values
        iRouteMin = New Point3D
        iRouteMax = New Point3D
        'Set initial values
        iRouteMin.X = RefRoute.CompSegment(0).StartPoint.Easting
        iRouteMin.Y = RefRoute.CompSegment(0).StartPoint.Northing
        iRouteMax.X = RefRoute.CompSegment(0).StartPoint.Easting
        iRouteMax.Y = RefRoute.CompSegment(0).StartPoint.Northing
        'Go through the route point to calculate the min/max values
        For s = 0 To CompSegment.Count - 1
            If RefRoute.CompSegment(s).StartPoint.Easting < iRouteMin.X Then iRouteMin.X = RefRoute.CompSegment(s).StartPoint.Easting
            If RefRoute.CompSegment(s).StartPoint.Easting > iRouteMax.X Then iRouteMax.X = RefRoute.CompSegment(s).StartPoint.Easting
            If RefRoute.CompSegment(s).StartPoint.Northing < iRouteMin.Y Then iRouteMin.Y = RefRoute.CompSegment(s).StartPoint.Northing
            If RefRoute.CompSegment(s).StartPoint.Northing > iRouteMax.Y Then iRouteMax.Y = RefRoute.CompSegment(s).StartPoint.Northing
            If RefRoute.CompSegment(s).EndPoint.Easting < iRouteMin.X Then iRouteMin.X = RefRoute.CompSegment(s).EndPoint.Easting
            If RefRoute.CompSegment(s).EndPoint.Easting > iRouteMax.X Then iRouteMax.X = RefRoute.CompSegment(s).EndPoint.Easting
            If RefRoute.CompSegment(s).EndPoint.Northing < iRouteMin.Y Then iRouteMin.Y = RefRoute.CompSegment(s).EndPoint.Northing
            If RefRoute.CompSegment(s).EndPoint.Northing > iRouteMax.Y Then iRouteMax.Y = RefRoute.CompSegment(s).EndPoint.Northing
            If RefRoute.CompSegment(s).CenterPoint.Easting < iRouteMin.X Then iRouteMin.X = RefRoute.CompSegment(s).CenterPoint.Easting
            If RefRoute.CompSegment(s).CenterPoint.Easting > iRouteMax.X Then iRouteMax.X = RefRoute.CompSegment(s).CenterPoint.Easting
            If RefRoute.CompSegment(s).CenterPoint.Northing < iRouteMin.Y Then iRouteMin.Y = RefRoute.CompSegment(s).CenterPoint.Northing
            If RefRoute.CompSegment(s).CenterPoint.Northing > iRouteMax.Y Then iRouteMax.Y = RefRoute.CompSegment(s).CenterPoint.Northing
            If RefRoute.CompSegment(s).VertexPoint.Easting < iRouteMin.X Then iRouteMin.X = RefRoute.CompSegment(s).VertexPoint.Easting
            If RefRoute.CompSegment(s).VertexPoint.Easting > iRouteMax.X Then iRouteMax.X = RefRoute.CompSegment(s).VertexPoint.Easting
            If RefRoute.CompSegment(s).VertexPoint.Northing < iRouteMin.Y Then iRouteMin.Y = RefRoute.CompSegment(s).VertexPoint.Northing
            If RefRoute.CompSegment(s).VertexPoint.Northing > iRouteMax.Y Then iRouteMax.Y = RefRoute.CompSegment(s).VertexPoint.Northing
        Next

    End Sub

    Private Sub SetGridKps(ByRef RefRoute As Route)
        'Set initial KP
        RefRoute.CompSegment(0).StartPoint.ForwardGridKP = 0.0
        'Update all the grid KPs
        For s = 0 To RefRoute.SegmentCount - 1
            'Update forward KPs
            RefRoute.CompSegment(s).StartPoint.ForwardGridKP = If(s = 0, 0.0, RefRoute.CompSegment(s - 1).EndPoint.ForwardGridKP)
            RefRoute.CompSegment(s).EndPoint.ForwardGridKP = RefRoute.CompSegment(s).StartPoint.ForwardGridKP + RefRoute.CompSegment(s).Length
            RefRoute.CompSegment(s).CenterPoint.ForwardGridKP = RefRoute.CompSegment(s).StartPoint.ForwardGridKP + (RefRoute.CompSegment(s).Length / 2)
            RefRoute.CompSegment(s).VertexPoint.ForwardGridKP = RefRoute.CompSegment(s).StartPoint.ForwardGridKP + (RefRoute.CompSegment(s).Length / 2)
        Next
        For s = 0 To RefRoute.SegmentCount - 1
            'Update reverse KPs
            RefRoute.CompSegment(s).StartPoint.ReverseGridKP = If(s = 0, Lenght, RefRoute.CompSegment(s - 1).EndPoint.ReverseGridKP)
            RefRoute.CompSegment(s).EndPoint.ReverseGridKP = RefRoute.CompSegment(s).StartPoint.ReverseGridKP - RefRoute.CompSegment(s).Length
            RefRoute.CompSegment(s).CenterPoint.ReverseGridKP = RefRoute.CompSegment(s).StartPoint.ReverseGridKP - (RefRoute.CompSegment(s).Length / 2)
            RefRoute.CompSegment(s).VertexPoint.ReverseGridKP = RefRoute.CompSegment(s).StartPoint.ReverseGridKP - (RefRoute.CompSegment(s).Length / 2)
        Next
    End Sub

    'Private Function GetReversedRoute(ByVal RefRoute As Route) As Route
    '    Dim revRoute As New Route

    '    For seg = RefRoute.SegmentCount - 1 To 0 Step -1
    '        Dim exStart, exEnd As New RoutePoint
    '        Dim revSeg As New RouteSegment

    '        revSeg = RefRoute.CompSegment(seg)
    '        exStart = RefRoute.CompSegment(seg).StartPoint
    '        exEnd = RefRoute.CompSegment(seg).EndPoint
    '        revSeg.IsArcClockwise = Not RefRoute.CompSegment(seg).IsArcClockwise
    '        revSeg.StartPoint = exEnd
    '        revSeg.EndPoint = exStart
    '        revRoute.CompSegment.Add(revSeg)

    '    Next
    '    SetGridKps(revRoute)
    '    SetDataRange(revRoute)
    '    revRoute.IsLoaded = LoadingStatus.Segments
    '    Return revRoute

    'End Function

    Public Function ToXml() As String
        Dim resXdoc As New XDocument()
        Dim xRoot As New XElement("Route")
        xRoot.Add(New XElement("Name", Name))
        xRoot.Add(New XElement("IsRouteLoaded", IsRouteLoaded))

        Dim xSeg As New XElement("Segments")

        For Each seg In ComputedSegments
            xSeg.Add(seg.ToXElement)
        Next

        Dim xfgkg As New XElement("ForwardGridPoints")
        For Each fgkp In ForwardGridPoints
            xfgkg.Add(fgkp.ToXElement())
        Next

        xRoot.Add(xSeg)
        xRoot.Add(xfgkg)

        resXdoc.Add(xRoot)
        resXdoc.Declaration = New XDeclaration("1.0", "UTF-8", "yes")
        Return resXdoc.ToString(SaveOptions.None)

    End Function

    Public Sub New(routeXml As String)

        IsLoaded = LoadingStatus.NotLoaded

        Dim xnav As New XDocument()
        xnav = XDocument.Parse(routeXml)
        Dim xRoot As XElement = xnav.Root
        Name = xRoot.Element("Name").Value
        IsLoaded = DirectCast([Enum].Parse(GetType(LoadingStatus), xRoot.Element("IsRouteLoaded").Value), LoadingStatus)

        Dim xSeg = xRoot.Element("Segments")
        Dim xfgkg = xRoot.Element("ForwardGridPoints")

        If xSeg.HasElements Then
            For Each seg In xSeg.Elements("RouteSegment")
                Dim rs As New RouteSegment(seg)
                CompSegment.Add(rs)
                Dim ms As New MasterSegment
                Dim mp1 As New MasterPoint
                Dim mp2 As New MasterPoint
                mp1.Easting = rs.StartPoint.Easting
                mp1.Northing = rs.StartPoint.Northing
                mp2.Easting = rs.EndPoint.Easting
                mp2.Northing = rs.EndPoint.Northing
                If Not rs.IsLine Then
                    ms.IsArc = True
                    ms.IsArcClockwise = rs.IsArcClockwise
                    ms.Radious = rs.Radious
                End If
                If CompSegment.Count = 1 Then iMastPoints.Add(mp1)
                iMastPoints.Add(mp2)
                iMastSegs.Add(ms)
            Next
        End If

        If xfgkg.HasElements Then
            For Each pnt1 In xfgkg.Elements("RoutePoint")
                iFwdGrdKps.Add(New RoutePoint(pnt1))
            Next
        End If

    End Sub

    Public Function ToXmlDocument() As XDocument
        Dim resXdoc As New XDocument()
        Dim xRoot As New XElement("Route")
        xRoot.Add(New XElement("Name", Name))
        xRoot.Add(New XElement("IsRouteLoaded", IsRouteLoaded))

        Dim xSeg As New XElement("Segments")

        For Each seg In ComputedSegments
            xSeg.Add(seg.ToXElement)
        Next

        Dim xfgkg As New XElement("ForwardGridPoints")
        For Each fgkp In ForwardGridPoints
            xfgkg.Add(fgkp.ToXElement())
        Next

        xRoot.Add(xSeg)
        xRoot.Add(xfgkg)

        resXdoc.Add(xRoot)
        resXdoc.Declaration = New XDeclaration("1.0", "UTF-8", "yes")
        Return resXdoc

    End Function

    Public Sub New(routeXDocument As XDocument)

        IsLoaded = LoadingStatus.NotLoaded
        If routeXDocument Is Nothing Then Exit Sub

        Dim xnav As New XDocument()
        xnav = routeXDocument
        Dim xRoot As XElement = xnav.Root
        Name = xRoot.Element("Name").Value
        IsLoaded = DirectCast([Enum].Parse(GetType(LoadingStatus), xRoot.Element("IsRouteLoaded").Value), LoadingStatus)

        Dim xSeg = xRoot.Element("Segments")
        Dim xfgkg = xRoot.Element("ForwardGridPoints")

        If xSeg.HasElements Then
            For Each seg In xSeg.Elements("RouteSegment")
                Dim rs As New RouteSegment(seg)
                CompSegment.Add(rs)
                Dim ms As New MasterSegment
                Dim mp1 As New MasterPoint
                Dim mp2 As New MasterPoint
                mp1.Easting = rs.StartPoint.Easting
                mp1.Northing = rs.StartPoint.Northing
                mp2.Easting = rs.EndPoint.Easting
                mp2.Northing = rs.EndPoint.Northing
                If Not rs.IsLine Then
                    ms.IsArc = True
                    ms.IsArcClockwise = rs.IsArcClockwise
                    ms.Radious = rs.Radious
                End If
                If CompSegment.Count = 1 Then iMastPoints.Add(mp1)
                iMastPoints.Add(mp2)
                iMastSegs.Add(ms)
            Next
        End If

        If xfgkg.HasElements Then
            For Each pnt1 In xfgkg.Elements("RoutePoint")
                iFwdGrdKps.Add(New RoutePoint(pnt1))
            Next
        End If

    End Sub

    Public Sub ApplyProjection(CRS As Projections)
        Dim tmpRte As New Route

        Try
            tmpRte = ComputeEllipsoidDistances(Me, CRS)
            CompSegment = tmpRte.ComputedSegments
            iFwdGrdKps = tmpRte.ForwardGridPoints
            IsLoaded = LoadingStatus.All
        Catch ex As Exception

        End Try

    End Sub

End Class