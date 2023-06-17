Attribute VB_Name = "modBrentTestFunctions"

'@Folder("Tests.helper stuff")

Option Explicit

Private Const defaultA As Double = 1
Private Const defaultX0 As Double = -2
Private Const defaultY0 As Double = -4


Public Function Parabola_Vertex( _
    ByVal x As Double, _
    Optional ByVal a As Variant, _
    Optional ByVal x0 As Variant, _
    Optional ByVal y0 As Variant _
        ) As Double
    
    If IsMissing(a) Or IsEmpty(a) Then
        a = defaultA
    End If
    If IsMissing(x0) Or IsEmpty(x0) Then
        x0 = defaultX0
    End If
    If IsMissing(y0) Or IsEmpty(y0) Then
        y0 = defaultY0
    End If
    
    Parabola_Vertex = a * (x - x0) ^ 2 + y0
    
End Function


Public Function Parabola_Vertex_LeftRoot( _
    Optional ByVal a As Variant, _
    Optional ByVal x0 As Variant, _
    Optional ByVal y0 As Variant _
        ) As Double
    
    Dim b As Double
    Dim c As Double
    
    
    If IsMissing(a) Or IsEmpty(a) Then
        a = defaultA
    End If
    If IsMissing(x0) Or IsEmpty(x0) Then
        x0 = defaultX0
    End If
    If IsMissing(y0) Or IsEmpty(y0) Then
        y0 = defaultY0
    End If
    
    b = -2 * a * x0
    c = y0 + b ^ 2 / 4 / a
    
    Parabola_Vertex_LeftRoot = (-b - Sqr(b ^ 2 - 4 * a * c)) / 2 / a
    
End Function


Public Function Parabola_Vertex_RightRoot( _
    Optional ByVal a As Variant, _
    Optional ByVal x0 As Variant, _
    Optional ByVal y0 As Variant _
        ) As Double
    
    Dim b As Double
    Dim c As Double
    
    
    If IsMissing(a) Or IsEmpty(a) Then
        a = defaultA
    End If
    If IsMissing(x0) Or IsEmpty(x0) Then
        x0 = defaultX0
    End If
    If IsMissing(y0) Or IsEmpty(y0) Then
        y0 = defaultY0
    End If
    
    b = -2 * a * x0
    c = y0 + b ^ 2 / 4 / a
    
    Parabola_Vertex_RightRoot = (-b + Sqr(b ^ 2 - 4 * a * c)) / 2 / a
    
End Function
