Option Explicit

Rem -----------------------------------------------------------------------------------------------------------------------
Rem --- Internal Variables                                                                           ---
Rem -----------------------------------------------------------------------------------------------------------------------

Public Xmax As Double
Public Cpx As Double
Public Cpy As Double
Public r As Double
Public v1 As Long
Public v2 As Long
Public v3 As Long

Rem -----------------------------------------------------------------------------------------------------------------------
Rem --- CONSTRUCTORS / DESTRUCTORS                                                                    ---
Rem -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()

End Sub      'Constructor

Rem -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    
End Sub 'Destructor

Rem -----------------------------------------------------------------------------------------------------------------------
Rem --- CLASS GET/LET/SET PROPERTIES                                                                 ---
Rem -----------------------------------------------------------------------------------------------------------------------

Rem ---
Rem --- clone the the triangle
Rem ---
Public Property Get Clone() As clsTriangle
    Dim myclone As clsTriangle

    Set myclone = New clsTriangle

    myclone.Xmax = Xmax
    myclone.Cpx = Cpx
    myclone.Cpy = Cpy
    myclone.r = r
    myclone.v1 = v1
    myclone.v2 = v2
    myclone.v3 = v3

    Set Clone = myclone
End Property

Rem -----------------------------------------------------------------------------------------------------------------------
Rem --- CLASS METHODS                                                                              ---
Rem -----------------------------------------------------------------------------------------------------------------------

Rem ---
Rem --- calculate distance between point p1 and p2
Rem ---
Private Function distance(p1 As Variant, p2 As Variant) As Double
    Dim a As Double
    Dim b As Double
    
    a = p2(0) - p1(0)
    b = p2(1) - p1(1)
    
    distance = Sqr((a * a) + (b * b))
    
End Function

Rem ---
Rem --- check if point is inside circle formed by the vertices of the triangle
Rem ---
Public Function isInCircle(ByRef Point As Variant) As Boolean
    Dim ret As Boolean
    Dim Cp(1) As Double
    
    Cp(0) = Cpx
    Cp(1) = Cpy
    
    ret = False
    
    If Point(0) < Xmax Then
        If distance(Point, Cp) <= r Then
            ret = True
        End If
    End If
    
    isInCircle = ret
End Function
