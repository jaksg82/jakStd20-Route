Public Class RouteSegment
    Public Property StartPoint As RoutePoint
    Public Property EndPoint As RoutePoint
    Public Property CenterPoint As RoutePoint
    Public Property VertexPoint As RoutePoint
    Public Property IsLine As Boolean
    Public Property Radious As Double
    Public Property IsArcClockwise As Boolean
    Public Property Length As Double

    Public Sub New()
        StartPoint = New RoutePoint
        EndPoint = New RoutePoint
        CenterPoint = New RoutePoint
        VertexPoint = New RoutePoint
        IsLine = True
        Radious = Double.NaN
        IsArcClockwise = True
        Length = Double.NaN

    End Sub

    Public Sub New(segment As XElement)
        StartPoint = New RoutePoint
        EndPoint = New RoutePoint
        CenterPoint = New RoutePoint
        VertexPoint = New RoutePoint
        IsLine = True
        Radious = Double.NaN
        IsArcClockwise = True
        Length = Double.NaN

        If segment Is Nothing Then Exit Sub

        If segment.Name = "RouteSegment" Then
            StartPoint = New RoutePoint(segment.Element("StartPoint").Element("RoutePoint"))
            EndPoint = New RoutePoint(segment.Element("EndPoint").Element("RoutePoint"))
            CenterPoint = New RoutePoint(segment.Element("CenterPoint").Element("RoutePoint"))
            VertexPoint = New RoutePoint(segment.Element("VertexPoint").Element("RoutePoint"))
            If Not Boolean.TryParse(segment.Element("IsLine").Value, IsLine) Then IsLine = False
            If Not Double.TryParse(segment.Element("Radious").Value, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Radious) Then Radious = Double.NaN
            If Not Boolean.TryParse(segment.Element("IsArcClockwise").Value, IsArcClockwise) Then IsArcClockwise = True
            If Not Double.TryParse(segment.Element("Length").Value, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Length) Then Length = Double.NaN

        End If

    End Sub

    Public Function ToXElement() As XElement
        Dim xseg As New XElement("RouteSegment")
        xseg.Add(New XElement("StartPoint", StartPoint.ToXElement))
        xseg.Add(New XElement("EndPoint", EndPoint.ToXElement))
        xseg.Add(New XElement("CenterPoint", CenterPoint.ToXElement))
        xseg.Add(New XElement("VertexPoint", VertexPoint.ToXElement))
        xseg.Add(New XElement("IsLine", IsLine))
        xseg.Add(New XElement("Radious", Radious))
        xseg.Add(New XElement("IsArcClockwise", IsArcClockwise))
        xseg.Add(New XElement("Length", Length))

        Return xseg

    End Function

End Class
