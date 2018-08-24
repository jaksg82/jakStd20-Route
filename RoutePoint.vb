Imports jakStd20_MathExt

Public Class RoutePoint

    Public Property Easting As Double
    Public Property Northing As Double
    Public Property ForwardGridKP As Double
    Public Property ReverseGridKP As Double
    Public Property Heading As Double
    Public Property Latitude As Double
    Public Property Longitude As Double
    Public Property ForwardEllipsoidKP As Double
    Public Property ReverseEllipsoidKP As Double
    Public Property Name As String

    Public Sub New()
        Easting = Double.NaN
        Northing = Double.NaN
        ForwardGridKP = Double.NaN
        ReverseGridKP = Double.NaN
        Heading = Double.NaN
        Latitude = Double.NaN
        Longitude = Double.NaN
        ForwardEllipsoidKP = Double.NaN
        ReverseEllipsoidKP = Double.NaN
        Name = ""

    End Sub

    Public Sub New(point As MasterPoint)
        If point Is Nothing Then
            Easting = Double.NaN
            Northing = Double.NaN
            ForwardGridKP = Double.NaN
            ReverseGridKP = Double.NaN
            Heading = Double.NaN
            Latitude = Double.NaN
            Longitude = Double.NaN
            ForwardEllipsoidKP = Double.NaN
            ReverseEllipsoidKP = Double.NaN
            Name = ""
        Else
            Easting = point.Easting
            Northing = point.Northing
            ForwardGridKP = Double.NaN
            ReverseGridKP = Double.NaN
            Heading = Double.NaN
            Latitude = Double.NaN
            Longitude = Double.NaN
            ForwardEllipsoidKP = Double.NaN
            ReverseEllipsoidKP = Double.NaN
            Name = point.Name
        End If
    End Sub

    Public Sub New(point As XElement)
        Dim tmpTouple() As String
        If point Is Nothing Then
            Easting = Double.NaN
            Northing = Double.NaN
            ForwardGridKP = Double.NaN
            ReverseGridKP = Double.NaN
            Heading = Double.NaN
            Latitude = Double.NaN
            Longitude = Double.NaN
            ForwardEllipsoidKP = Double.NaN
            ReverseEllipsoidKP = Double.NaN
            Name = ""
        Else
            Easting = Double.NaN
            Northing = Double.NaN
            ForwardGridKP = Double.NaN
            ReverseGridKP = Double.NaN
            Heading = Double.NaN
            Latitude = Double.NaN
            Longitude = Double.NaN
            ForwardEllipsoidKP = Double.NaN
            ReverseEllipsoidKP = Double.NaN
            Name = ""

            If point.Name = "RoutePoint" Then
                If point.Elements("Name").Count > 0 Then
                    Name = point.Element("Name").Value
                End If
                tmpTouple = point.Element("EN").Value.Split("|"c)
                If tmpTouple.Count = 2 Then
                    If Not Double.TryParse(tmpTouple(0), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Easting) Then Easting = Double.NaN
                    If Not Double.TryParse(tmpTouple(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Northing) Then Northing = Double.NaN
                End If
                tmpTouple = point.Element("Grids").Value.Split("|"c)
                If tmpTouple.Count = 2 Then
                    If Not Double.TryParse(tmpTouple(0), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, ForwardGridKP) Then ForwardGridKP = Double.NaN
                    If Not Double.TryParse(tmpTouple(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, ReverseGridKP) Then ReverseGridKP = Double.NaN
                End If
                tmpTouple = point.Element("LL").Value.Split("|"c)
                If tmpTouple.Count = 2 Then
                    If Not Double.TryParse(tmpTouple(0), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Latitude) Then Latitude = Double.NaN
                    If Not Double.TryParse(tmpTouple(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Longitude) Then Longitude = Double.NaN
                End If
                tmpTouple = point.Element("Ellips").Value.Split("|"c)
                If tmpTouple.Count = 2 Then
                    If Not Double.TryParse(tmpTouple(0), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, ForwardEllipsoidKP) Then ForwardEllipsoidKP = Double.NaN
                    If Not Double.TryParse(tmpTouple(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, ReverseEllipsoidKP) Then ReverseEllipsoidKP = Double.NaN
                End If
                If Not Double.TryParse(point.Element("Hdg").Value, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Heading) Then Heading = Double.NaN
            End If
        End If
    End Sub

    Public Sub ENfromPoint3D(point As Point3D)
        If point IsNot Nothing Then
            Easting = point.X
            Northing = point.Y
        End If
    End Sub

    Public Sub LLfromPoint3D(point As Point3D)
        If point IsNot Nothing Then
            Latitude = point.Y
            Longitude = point.X
        End If
    End Sub

    Public Function ENToPoint3D() As Point3D
        Return New Point3D(Easting, Northing, 0.0)
    End Function

    Public Function LLToPoint3D() As Point3D
        Return New Point3D(Longitude, Latitude, 0.0)
    End Function

    Public Function ToXElement() As XElement
        Dim xpoint As New XElement("RoutePoint")

        xpoint.Add(New XElement("Name", Name))
        xpoint.Add(New XElement("EN", Easting & "|" & Northing))
        xpoint.Add(New XElement("Grids", ForwardGridKP & "|" & ReverseGridKP))
        xpoint.Add(New XElement("Hdg", Heading))
        xpoint.Add(New XElement("LL", Latitude & "|" & Longitude))
        xpoint.Add(New XElement("Ellips", ForwardEllipsoidKP & "|" & ReverseEllipsoidKP))

        Return xpoint

    End Function

End Class