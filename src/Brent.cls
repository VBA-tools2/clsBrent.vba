VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Brent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'==============================================================================
'downloaded from
'<http://www.quantcode.com/modules/mydownloads/singlefile.php?cid=9&lid=232>
'==============================================================================

Option Explicit

Private maxIterations_ As Single
Private objectMode_ As Boolean
Private callbackObject_ As Object
Public Silent_ As Boolean

Public Sub AssignObject(callbackObject As Object)
  objectMode_ = True
  Set callbackObject_ = callbackObject
End Sub

Private Function RunFunction(FunctionName As String, x As Double) As Double
  If objectMode_ = False Then
    RunFunction = Application.Run(FunctionName, x)
  Else
    RunFunction = CallByName(callbackObject_, FunctionName, VbMethod, x)
  End If
End Function

Public Function Solve(FunctionName As String, Accuracy As Double, Guess As Double, xa As Double, xb As Double) As Double
  Dim Iter As Single
  Dim e, TempDbl, fxa, fxb, soln As Double

  Dim tempstr As String
  If Not xa < xb Then
    tempstr = "invalid range: xa (" & xa & ") >= xb (" & xb
    ShowMessageBox tempstr
    Exit Function
  End If

  
  fxa = RunFunction(FunctionName, xa)
  If (Abs(fxa) < Accuracy) Then
    Solve = xa
    Exit Function
  End If
  
  fxb = RunFunction(FunctionName, xb)
  If (Abs(fxb) < Accuracy) Then
    Solve = xb
    Exit Function
  End If
  
  Iter = 2
  
  If Not fxa * fxb < 0 Then
    tempstr = "root not bracketed: f[" & xa & "," & (xb) & "] -> [" & (fxa) & "," & (fxb) & "]"
    ShowMessageBox tempstr
    Exit Function
  End If
  If Not Guess > xa Then
    tempstr = " guess (" & Guess & ") < xa (" & xa & ")"
    ShowMessageBox tempstr
    Exit Function
  End If
  If Not Guess < xb Then
    tempstr = " guess (" & Guess & ") > xb (" & xb & ")"
    ShowMessageBox tempstr
    Exit Function
  End If
  
  soln = Guess
  Dim xAccuracy As Double
  xAccuracy = Accuracy
  Dim fsoln As Double
  Dim p As Double
  Dim q As Double
  Dim R As Double
  Dim s As Double
  Dim xTol1 As Double
  Dim xMid As Double
  Dim d As Double
  Dim Eps As Double
  Eps = 2 ^ -52
  
  d = 0
  e = 0
  
  
  soln = xb
  fsoln = fxb
  Do While (Iter <= maxIterations_)
    If ((fsoln > 0 And fxb > 0) Or (fsoln < 0 And fxb < 0)) Then
      xb = xa
      fxb = fxa
      d = soln - xa
      e = d
    End If
    If (Abs(fxb) < Abs(fsoln)) Then
      xa = soln
      soln = xb
      xb = xa
      fxa = fsoln
      fsoln = fxb
      fxb = fxa
    End If

    xTol1 = 2 * Eps * Abs(soln) + 0.5 * xAccuracy
    xMid = 0.5 * (xb - soln)
    'exit if function has converged
    If (Abs(xMid) <= xTol1 Or fsoln = 0) Then
      Solve = soln
      Exit Function
    End If
    If (Abs(e) >= xTol1 And Abs(fxa) > Abs(fsoln)) Then
    
      ' inverse quadratic interpolation
      s = fsoln / fxa
      If (xa = xb) Then
        p = 2 * xMid * s
        q = 1 - s
      Else
        q = fxa / fxb
        R = fsoln / fxb
        p = s * (2 * xMid * q * (q - R) - (soln - xa) * (R - 1))
        q = (q - 1) * (R - 1) * (s - 1)
      End If
      If (p > 0) Then
        q = -q
      End If
      p = Abs(p)
      
     
      If (p < 0.5 * Application.min(3 * xMid * q - Abs(xTol1 * q), Abs(e * q))) Then
        'secant mathod is fast so consider interpolation
        e = d
        d = p / q
      Else
        'bisection step is considered since the secant method is not progressing fast
        d = xMid
        e = d
      End If
    Else
      ' Bounds decreasing too slowly, use bisection
      d = xMid
      e = d
    End If
    xa = soln
    fxa = fsoln
    If (Abs(d) > xTol1) Then
      soln = soln + d
    Else
      soln = soln + sign(xTol1, xMid)
    End If
    fsoln = RunFunction(FunctionName, soln)
    Iter = Iter + 1
  Loop
  
End Function

Public Sub ShowMessageBox(Msg As String)
  If Silent_ = False Then
    MsgBox Msg
  End If
End Sub

Public Function sign(a As Double, b As Double) As Double
  If b >= 0 Then
    sign = Abs(a)
  Else
    sign = -Abs(a)
  End If
End Function

Private Sub Class_Initialize()
  maxIterations_ = 100
  objectMode_ = False
'  Silent_ = False
  Silent_ = True
End Sub
