Public Class RouteFileEntry
    Public Property Code As String
    Public Property East As Double
    Public Property North As Double
    Public Property Radious As Double
    Public Property IsArcClockwise As Boolean

    Public Sub New()
        Code = ""
        East = Double.NaN
        North = Double.NaN
        Radious = Double.NaN
        IsArcClockwise = False
    End Sub

    Public Sub New(line As String)

        If String.IsNullOrEmpty(line) Or String.IsNullOrWhiteSpace(line) Then
            Code = ""
            East = Double.NaN
            North = Double.NaN
            Radious = Double.NaN
            IsArcClockwise = False
        Else
            Dim tmpValues() As String
            tmpValues = line.Split(","c)
            If tmpValues.Count = 3 Then
                Select Case tmpValues(0).ToUpper
                    Case "P"
                        Code = "P"
                        If Not Double.TryParse(tmpValues(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, East) Then East = 0
                        If Not Double.TryParse(tmpValues(2), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, North) Then North = 0
                        Radious = 0.0
                        IsArcClockwise = False

                    Case "I"
                        Code = "I"
                        If Not Double.TryParse(tmpValues(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, East) Then East = 0
                        If Not Double.TryParse(tmpValues(2), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, North) Then North = 0
                        Radious = 0.0
                        IsArcClockwise = False

                    Case "A"
                        Code = "A"
                        If Not Double.TryParse(tmpValues(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Radious) Then Radious = 0
                        IsArcClockwise = tmpValues(2).Substring(0, 1).ToUpper = "R"
                        East = 0.0
                        North = 0.0

                    Case Else
                        Code = ""
                        East = Double.NaN
                        North = Double.NaN
                        Radious = Double.NaN
                        IsArcClockwise = False

                End Select

            ElseIf tmpValues.Count = 4 Then
                Select Case tmpValues(0).ToUpper
                    Case "P"
                        Code = "P"
                        If Not Double.TryParse(tmpValues(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, East) Then East = 0
                        If Not Double.TryParse(tmpValues(2), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, North) Then North = 0
                        If Double.TryParse(tmpValues(3), Radious) Then
                            Code = "I"
                        Else
                            Radious = 0.0
                        End If
                        IsArcClockwise = False

                    Case "I"
                        Code = "I"
                        If Not Double.TryParse(tmpValues(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, East) Then East = 0
                        If Not Double.TryParse(tmpValues(2), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, North) Then North = 0
                        If Not Double.TryParse(tmpValues(3), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Radious) Then Radious = 0
                        IsArcClockwise = False

                    Case "A"
                        Code = "A"
                        If Not Double.TryParse(tmpValues(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Radious) Then Radious = 0
                        IsArcClockwise = tmpValues(2).Substring(0, 1).ToUpper = "R"
                        East = 0.0
                        North = 0.0

                    Case Else
                        Code = ""
                        East = Double.NaN
                        North = Double.NaN
                        Radious = Double.NaN
                        IsArcClockwise = False

                End Select

            Else
                Code = ""
                East = Double.NaN
                North = Double.NaN
                Radious = Double.NaN
                IsArcClockwise = False
            End If

        End If
    End Sub
End Class
