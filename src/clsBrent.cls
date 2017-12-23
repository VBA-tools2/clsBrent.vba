VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBrent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'==============================================================================
'downloaded from
'<http://www.quantcode.com/modules/mydownloads/singlefile.php?cid=9&lid=232>
'    modified functions to be able to state more parameters to the
'    functions which are then called. Added an "X" to the function names
'    which are now named
'    - RunFunctionX
'    - SolveX
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

'==============================================================================

'modified function (originally 'RunFunction') which can handle an optional
'number of additional arguments, which can be stored in 'Arr'
Private Function RunFunctionX( _
    FunctionName As String, _
    x As Double, _
    Optional Arr As Variant _
        ) As Double
    
    If IsMissing(Arr) Then
        If objectMode_ = False Then
            RunFunctionX = Application.Run(FunctionName, x)
        Else
            RunFunctionX = CallByName(callbackObject_, FunctionName, VbMethod, x)
        End If
    Else
        Select Case UBound(Arr)
            Case 2
                If objectMode_ = False Then
                    RunFunctionX = Application.Run(FunctionName, _
                            x, Arr(0), Arr(1), Arr(2))
                Else
                    RunFunctionX = CallByName(callbackObject_, FunctionName, _
                        VbMethod, x, Arr(0), Arr(1), Arr(2))
                End If
            Case 5
                If objectMode_ = False Then
                    RunFunctionX = Application.Run(FunctionName, _
                            x, Arr(0), Arr(1), Arr(2), Arr(3), Arr(4), Arr(5))
                Else
                    RunFunctionX = CallByName(callbackObject_, FunctionName, _
                        VbMethod, x, Arr(0), Arr(1), Arr(2), Arr(3), Arr(4), Arr(5))
                End If
            Case 7
                If objectMode_ = False Then
                    RunFunctionX = Application.Run(FunctionName, _
                            x, Arr(0), Arr(1), Arr(2), Arr(3), Arr(4), Arr(5), _
                            Arr(6), Arr(7))
                Else
                    RunFunctionX = CallByName(callbackObject_, FunctionName, _
                        VbMethod, x, Arr(0), Arr(1), Arr(2), Arr(3), Arr(4), Arr(5), _
                        Arr(6), Arr(7))
                End If
        End Select
    End If
    
End Function

'modified function (originally 'Solve') which can handle an optional
'number of arbitrary arguments, which can be stored in 'Arr'
Public Function SolveX( _
    FunctionName As String, _
    Accuracy As Double, _
    Guess As Double, _
    xa As Double, _
    xb As Double, _
    Optional Arr As Variant _
        ) As Double
    
  Dim Iter As Single
  Dim e As Double
  Dim fxa As Double
  Dim fxb As Double
  Dim soln As Double
  
  Dim TempStr As String
  If Not xa < xb Then
    TempStr = "invalid range: xa (" & xa & ") >= xb (" & xb & ")"
    ShowMessageBox TempStr
    Exit Function
  End If

  
  fxa = RunFunctionX(FunctionName, xa, Arr)
  If (Abs(fxa) < Accuracy) Then
    SolveX = xa
    Exit Function
  End If
  
  fxb = RunFunctionX(FunctionName, xb, Arr)
  If (Abs(fxb) < Accuracy) Then
    SolveX = xb
    Exit Function
  End If
  
  Iter = 2
  
  If Not fxa * fxb < 0 Then
    TempStr = "root not bracketed: f[" & xa & "," & (xb) & "] -> [" & (fxa) & "," & (fxb) & "]"
    ShowMessageBox TempStr
    Exit Function
  End If
  If Not Guess > xa Then
    TempStr = " guess (" & Guess & ") < xa (" & xa & ")"
    ShowMessageBox TempStr
    Exit Function
  End If
  If Not Guess < xb Then
    TempStr = " guess (" & Guess & ") > xb (" & xb & ")"
    ShowMessageBox TempStr
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
      SolveX = soln
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
    fsoln = RunFunctionX(FunctionName, soln, Arr)
    Iter = Iter + 1
  Loop
  
End Function

'==============================================================================

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
