Public Class MasterPoint
    Public Property Easting As Double
    Public Property Northing As Double
    Public Property Type As PointType
    Public Property Name As String
    Public Property Radious As Double

    Public Enum PointType
        TangentPoint
        VertexPoint
    End Enum

    Public Sub New()
        Type = PointType.VertexPoint
        Easting = 0.0
        Northing = 0.0
        Name = "Vertex"
        Radious = 0.0
    End Sub
End Class
