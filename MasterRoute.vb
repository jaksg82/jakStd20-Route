Public Class MasterRoute
    Public ReadOnly Property Points As List(Of MasterPoint)
    Public ReadOnly Property Segments As List(Of MasterSegment)

    Public Sub New()
        Points = New List(Of MasterPoint)
        Segments = New List(Of MasterSegment)
    End Sub
End Class
