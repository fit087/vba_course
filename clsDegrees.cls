VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDegrees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private angle As Double

Private graus As Integer
Private min As Integer
Private sec As Double

'Type graus_min_sec

'    graus As Integer
'    min As Integer
'    sec As Double

'End Type

Public Property Get degrees() As Integer

    'degrees = Int(angle)
    'degrees = graus_min_sec.graus
    degrees = graus
    
End Property
Public Property Get minutes() As Integer
'Public Property Get min() As Integer

    'degrees = Int(angle)
    'min = graus_min_sec.min
    minutes = min
    
End Property

Public Property Get seconds() As Double

    'degrees = Int(angle)
    'sec = graus_min_sec.sec
    seconds = sec
    
End Property

Public Property Let angle_dec(ByVal angulo As Double)

    Dim resto As Double
    angle = angulo
    Debug.Print "Angulo"; angulo
    
    'graus_min_sec.graus = Int(angulo)
    graus = Int(angulo)
    
    'resto = (angulo - graus_min_sec.graus) * 60
    resto = (angulo - graus) * 60
    'graus_min_sec.min = Int(resto)
    min = Int(resto)
    
    'resto = (resto - graus_min_sec.min) * 60
    resto = (resto - min) * 60
    'graus_min_sec.sec = resto
    sec = resto
    'Debug.Print "Os Segundos Restantes resto = "; resto
    'Debug.Print "Os Segundos Restantes sec = "; sec
    
End Property
