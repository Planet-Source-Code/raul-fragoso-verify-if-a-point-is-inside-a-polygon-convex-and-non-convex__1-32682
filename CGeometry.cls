VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
'This is a very useful class to store polygon's vertices and
'check if the polygon is convex or not and if a point is inside
'the polygon or not
'By Raul Fragoso on 14-March-2002
'****************************************************************

'Collection wich stores the vertexes
Private m_Vertices As Collection

'Adds a new vertex to the polygon
Public Sub AddVertex(X As Long, Y As Long)

    Dim pa As New CPoint
    pa.X = X
    pa.Y = Y
    m_Vertices.Add pa

End Sub

'Removes a vertex from the polygon
Public Sub RemoveVertex(iIndex As Long)

    If iIndex <> 0 And m_Vertices.Count <= iIndex Then
        m_Vertices.Remove iIndex
    End If

End Sub

'Returns the x coordinate
Public Function GetVertex(iIndex As Long) As CPoint

    Dim pa As New CPoint
    Set pa = m_Vertices(iIndex)
    Set GetVertex = pa

End Function

'Returns the number of points in the polygon
Public Function VertexCount() As Long

    VertexCount = m_Vertices.Count

End Function

'Clears all the points in the polygon
Public Sub ClearVertices()

   Set m_Vertices = Nothing
   Set m_Vertices = New Collection

End Sub

' Return True if the point is in the polygon.
Public Function PointInPolygon(ByVal X As Single, ByVal Y As Single) As Boolean

    Dim pt As Integer
    Dim total_angle As Single

    ' Get the angle between the point and the
    ' first and last vertices.
    total_angle = GetAngle( _
        m_Vertices(m_Vertices.Count).X, m_Vertices(m_Vertices.Count).Y, _
        X, Y, _
        m_Vertices(1).X, m_Vertices(1).Y)

    ' Add the angles from the point to each other
    ' pair of vertices.
    For pt = 1 To m_Vertices.Count - 1
        total_angle = total_angle + GetAngle(m_Vertices(pt).X, m_Vertices(pt).Y, X, Y, m_Vertices(pt + 1).X, m_Vertices(pt + 1).Y)
    Next pt

    ' The total angle should be 2 * PI or -2 * PI if
    ' the point is in the polygon and close to zero
    ' if the point is outside the polygon.
    PointInPolygon = (Abs(total_angle) > 0.000001)
End Function
' Return True if the polygon is convex.
Public Function PolygonIsConvex() As Boolean
Dim pt As Integer
Dim cross_product As Single
Dim got_negative As Boolean
Dim got_positive As Boolean

    ' For each set of three adjacent points A, B, C,
    ' find the dot product AB · BC. If the sign of
    ' all the dot products is the same, the angles
    ' are all positive or negative (depending on the
    ' order in which we visit them) so the polygon
    ' is convex.

    ' Assume the polygon is non-convex.
    PolygonIsConvex = False

    ' Look at the first set of points.
    cross_product = CrossProductLength( _
        m_Vertices(m_Vertices.Count - 1).X, m_Vertices(m_Vertices.Count - 1).Y, _
        m_Vertices(m_Vertices.Count).X, m_Vertices(m_Vertices.Count).Y, _
        m_Vertices(1).X, m_Vertices(1).Y)
    If cross_product < 0 Then
        got_negative = True
    ElseIf cross_product > 0 Then
        got_positive = True
    End If

    ' Look at the second set of points.
    cross_product = CrossProductLength( _
        m_Vertices(m_Vertices.Count).X, m_Vertices(m_Vertices.Count).Y, _
        m_Vertices(1).X, m_Vertices(1).Y, _
        m_Vertices(2).X, m_Vertices(2).Y)
    If cross_product < 0 Then
        got_negative = True
    ElseIf cross_product > 0 Then
        got_positive = True
    End If
    If got_negative And got_positive Then Exit Function

    ' Look at the remaining triples of points.
    For pt = 1 To m_Vertices.Count - 2
        cross_product = CrossProductLength( _
            m_Vertices(pt).X, m_Vertices(pt).Y, _
            m_Vertices(pt + 1).X, m_Vertices(pt + 1).Y, _
            m_Vertices(pt + 2).X, m_Vertices(pt + 2).Y)
        If cross_product < 0 Then
            got_negative = True
        ElseIf cross_product > 0 Then
            got_positive = True
        End If
        If got_negative And got_positive Then Exit Function
    Next pt

    ' If we got this far, the polygon is convex.
    PolygonIsConvex = True
End Function

' Return the angle with tangent opp/hyp. The returned
' value is between PI and -PI.
Private Function ATan2(ByVal opp As Single, ByVal adj As Single) As Single
    Dim angle As Single

    ' Get the basic angle.
    If Abs(adj) < 0.0001 Then
        angle = PI / 2
    Else
        angle = Abs(Atn(opp / adj))
    End If

    ' See if we are in quadrant 2 or 3.
    If adj < 0 Then
        ' angle > PI/2 or angle < -PI/2.
        angle = PI - angle
    End If

    ' See if we are in quadrant 3 or 4.
    If opp < 0 Then
        angle = -angle
    End If

    ' Return the result.
    ATan2 = angle
End Function
' Return the cross product AB x BC.
' The cross product is a vector perpendicular to AB
' and BC having length |AB| * |BC| * Sin(theta) and
' with direction given by the right-hand rule.
' For two vectors in the X-Y plane, the result is a
' vector with X and Y components 0 so the Z component
' gives the vector's length and direction.
Private Function CrossProductLength( _
    ByVal Ax As Single, ByVal Ay As Single, _
    ByVal Bx As Single, ByVal By As Single, _
    ByVal Cx As Single, ByVal Cy As Single _
  ) As Single
Dim BAx As Single
Dim BAy As Single
Dim BCx As Single
Dim BCy As Single

    ' Get the vectors' coordinates.
    BAx = Ax - Bx
    BAy = Ay - By
    BCx = Cx - Bx
    BCy = Cy - By

    ' Calculate the Z coordinate of the cross product.
    CrossProductLength = BAx * BCy - BAy * BCx
End Function
' Return the dot product AB · BC.
' Note that AB · BC = |AB| * |BC| * Cos(theta).
Private Function DotProduct( _
    ByVal Ax As Single, ByVal Ay As Single, _
    ByVal Bx As Single, ByVal By As Single, _
    ByVal Cx As Single, ByVal Cy As Single _
  ) As Single
Dim BAx As Single
Dim BAy As Single
Dim BCx As Single
Dim BCy As Single

    ' Get the vectors' coordinates.
    BAx = Ax - Bx
    BAy = Ay - By
    BCx = Cx - Bx
    BCy = Cy - By

    ' Calculate the dot product.
    DotProduct = BAx * BCx + BAy * BCy
End Function

' Return the angle ABC.
' Return a value between PI and -PI.
' Note that the value is the opposite of what you might
' expect because Y coordinates increase downward.
Private Function GetAngle(ByVal Ax As Single, ByVal Ay As Single, ByVal Bx As Single, ByVal By As Single, ByVal Cx As Single, ByVal Cy As Single) As Single
Dim dot_product As Single
Dim cross_product As Single

    ' Get the dot product and cross product.
    dot_product = DotProduct(Ax, Ay, Bx, By, Cx, Cy)
    cross_product = CrossProductLength(Ax, Ay, Bx, By, Cx, Cy)

    ' Calculate the angle.
    GetAngle = ATan2(cross_product, dot_product)
End Function

Private Sub Class_Initialize()

    Set m_Vertices = New Collection

End Sub

Private Sub Class_Terminate()

    Set m_Vertices = Nothing

End Sub
