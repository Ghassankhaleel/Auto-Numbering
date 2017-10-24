Public Class Graph
    ' Public NUM_VERTICES As Integer
    Private vertices() As Vertex
    Private adjMatrix(,) As Integer
    Private numVerts As Integer
    Private n As Integer
    Public Sub New(ByVal NUM_VERTICES As Integer)
        ReDim vertices(NUM_VERTICES)
        ReDim adjMatrix(NUM_VERTICES, NUM_VERTICES)
        numVerts = 0
        Dim j, k As Integer
        'n = NUM_VERTICES
        For j = 0 To NUM_VERTICES - 1
            For k = 0 To NUM_VERTICES - 1
                adjMatrix(j, k) = 0
            Next
        Next
    End Sub
    Public Sub addVertex(ByVal label As String)
        vertices(numVerts) = New Vertex(label)
        numVerts += 1
    End Sub
    Public Sub addEdge(ByVal start As Integer, ByVal eend As Integer)
        adjMatrix(start, eend) = 1
        adjMatrix(eend, start) = 1
    End Sub
    Public Sub showVertex(ByVal v As Integer)
        'Console.Write(vertices(v).label)
    End Sub


    Public Function noSuccessors() As Integer
        Dim isEdge As Boolean
        Dim row, col As Integer
        For row = 0 To n - 1
            isEdge = False
            For col = 0 To n - 1
                If adjMatrix(row, col) > 0 Then
                    isEdge = True
                    Exit For
                End If
            Next
            If (Not (isEdge)) Then
                Return row
            End If
        Next
        Return -1
    End Function


    Public Sub delVertex(ByVal vert As Integer)
        Dim j, row, col As Integer
        If (vert <> n - 1) Then
            For j = vert To n - 1
                vertices(j) = vertices(j + 1)
            Next
            For row = vert To n - 1
                moveRow(row, n)
            Next
            For col = vert To n - 1
                moveCol(row, n - 1)
            Next
        End If
    End Sub
    Private Sub moveRow(ByVal row As Integer, ByVal _
    length As Integer)
        Dim col As Integer
        For col = 0 To length - 1
            adjMatrix(row, col) = adjMatrix(row + 1, col)
        Next
    End Sub
    Private Sub moveCol(ByVal col As Integer, ByVal _
    length As Integer)
        Dim row As Integer
        For row = 0 To length - 1
            adjMatrix(row, col) = adjMatrix(row, col + 1)
        Next
    End Sub

    'Public Sub TopSort()
    

End Class
