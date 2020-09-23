VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************
'A simple class to hold a point
'*********************************************************

Private m_X As Long
Private m_Y As Long

Public Property Get X() As Long
    
    X = m_X

End Property

Public Property Get Y() As Long
    
    Y = m_Y

End Property

Public Property Let X(newX As Long)
    
    m_X = newX

End Property

Public Property Let Y(newY As Long)
    
    m_Y = newY

End Property

