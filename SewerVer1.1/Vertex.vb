Public Class Vertex
    Public wasVisited As Boolean
    Public label As String
    Public Sub New(ByVal label As String)
        Me.label = label
        wasVisited = False
    End Sub
End Class
