VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geometry"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblInside 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label lblConvex 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":0000
      Height          =   1065
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If current state is drawing or not
Dim m_Drawing As Boolean
'The object wich holds the coordinates
Dim m_Geometry As CGeometry
'Last X,Y mouse coordinates
Private m_LastX As Long
Private m_LastY As Long

' Draw the polygon.
Private Sub DrawPolygon()
    Cls
    If m_Geometry.VertexCount > 2 Then
        Dim points() As POINTAPI
        Dim i&, lPoints&
        ReDim points(1 To 1)
        'Build an array of POINTAPI object since the Polygon function needs that
        For i = 1 To m_Geometry.VertexCount
            lPoints = lPoints + 1
            ReDim Preserve points(1 To i)
            points(i).X = m_Geometry.GetVertex(i).X
            points(i).Y = m_Geometry.GetVertex(i).Y
        Next i
        Polygon Me.hdc, points(1), lPoints
        If m_Geometry.PolygonIsConvex Then
            lblConvex.BackColor = vbGreen
            lblConvex.Caption = "Convex"
        Else
            lblConvex.BackColor = vbRed
            lblConvex.Caption = "Non-Convex"
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If m_Drawing = False Then
            m_Geometry.ClearVertices
            Me.Cls
            lblConvex.Caption = ""
            lblConvex.BackColor = &H8000000F
            lblInside.Caption = ""
            lblInside.BackColor = &H8000000F
        End If
    End If

End Sub

Private Sub Form_Load()

    Set m_Geometry = New CGeometry

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_Drawing Then
        If Button = vbRightButton Then
            ' Stop drawing.
            m_Drawing = False
            Me.DrawMode = vbCopyPen
            DrawPolygon
        Else
            ' Add another point to the polygon.
            Line (m_Geometry.GetVertex(m_Geometry.VertexCount).X, m_Geometry.GetVertex(m_Geometry.VertexCount).Y)-(m_LastX, m_LastY)
            m_Geometry.AddVertex CLng(X), CLng(Y)
            Me.DrawMode = vbCopyPen
            Line (m_Geometry.GetVertex(m_Geometry.VertexCount - 1).X, m_Geometry.GetVertex(m_Geometry.VertexCount - 1).Y)-(m_Geometry.GetVertex(m_Geometry.VertexCount).X, m_Geometry.GetVertex(m_Geometry.VertexCount).Y)
            Me.DrawMode = vbInvert
        End If
    Else
        m_Drawing = True
        Me.DrawMode = vbInvert
        Cls
        lblConvex.Caption = ""
        lblConvex.BackColor = &H8000000F
        lblInside.Caption = ""
        lblInside.BackColor = &H8000000F
        m_Geometry.ClearVertices
        m_Geometry.AddVertex CLng(X), CLng(Y)
        m_LastX = X
        m_LastY = Y
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_Geometry.VertexCount > 2 And m_Drawing = False Then
        If m_Geometry.PointInPolygon(X, Y) Then
            lblInside.BackColor = vbGreen
            lblInside.Caption = "Inside"
        Else
            lblInside.BackColor = vbRed
            lblInside.Caption = "Outside"
        End If
    End If
    
    If Not m_Drawing Then Exit Sub

    Line (m_Geometry.GetVertex(m_Geometry.VertexCount).X, m_Geometry.GetVertex(m_Geometry.VertexCount).Y)-(m_LastX, m_LastY)
    'Saves the last X,Y value
    m_LastX = X
    m_LastY = Y
    Line (m_Geometry.GetVertex(m_Geometry.VertexCount).X, m_Geometry.GetVertex(m_Geometry.VertexCount).Y)-(m_LastX, m_LastY)

End Sub
