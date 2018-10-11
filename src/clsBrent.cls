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
'also with some additions that were inspired from
'<http://daft.engineer/hacks-and-kludges/numerical-methods-and-excel-vba/>
'==============================================================================

Option Explicit

Private pMaxIterations As Long
Private pObjectMode As Boolean
Private pCallbackObject As Object
Private pSilent As Boolean
Private pAccuracy As Double
Private pGuess As Double
Private pLowerBound As Double
Private pUpperBound As Double
Private pFunctionName As String
Private pArr As Variant

Private pbFunctionNameSet As Boolean
Private pbLowerBoundSet As Boolean
Private pbUpperBoundSet As Boolean
Private pbGuessSet As Boolean


Public Property Get MaxIterations() As Long
    MaxIterations = pMaxIterations
End Property

Public Property Let MaxIterations(Value As Long)
    pMaxIterations = Value
End Property


Public Property Get CallbackObject() As Object
    Set CallbackObject = pCallbackObject
End Property

Public Property Let CallbackObject(Value As Object)
    Set pCallbackObject = Value
End Property


Public Property Get Silent() As Boolean
    Silent = pSilent
End Property

Public Property Let Silent(Value As Boolean)
    pSilent = Value
End Property


Public Property Get Accuracy() As Double
    Accuracy = pAccuracy
End Property

Public Property Let Accuracy(Value As Double)
    pAccuracy = Value
End Property


Public Property Get Guess() As Double
    Guess = pGuess
End Property

Public Property Let Guess(Value As Double)
    pGuess = Value
    pbGuessSet = True
End Property


Public Property Get LowerBound() As Double
    LowerBound = pLowerBound
End Property

Public Property Let LowerBound(Value As Double)
    pLowerBound = Value
    pbLowerBoundSet = True
    Call CheckPrerequisites
End Property


Public Property Get UpperBound() As Double
    UpperBound = pUpperBound
End Property

Public Property Let UpperBound(Value As Double)
    pUpperBound = Value
    pbUpperBoundSet = True
    Call CheckPrerequisites
End Property


Public Property Get Arr() As Variant
    Arr = pArr
End Property

Public Property Let Arr(Value As Variant)
    pArr = Value
End Property


Public Property Get FunctionName() As String
    FunctionName = pFunctionName
End Property

Public Property Let FunctionName(Value As String)
    pFunctionName = Value
    pbFunctionNameSet = True
    Call CheckPrerequisites
End Property


Public Property Get Solve() As Variant
    'initialize with an error value
    Solve = CVErr(xlErrValue)
    
    'call function
    Solve = SolveX()
End Property


'------------------------------------------------------------------------------

Private Sub Class_Initialize()
    pMaxIterations = 100
    pObjectMode = False
    pSilent = True
    
    pAccuracy = 0.00001
End Sub


'==============================================================================

'modified function (originally 'Solve') which can handle an optional
'number of arbitrary arguments that can be stored in 'Arr'
'- xa = x_min
'- xb = x_max
Private Function SolveX() As Variant
    
    Dim xa As Double
    Dim xb As Double
    Dim Iter As Long
    Dim e As Double
    Dim fxa As Double
    Dim fxb As Double
    Dim soln As Double
    
    Dim TempStr As String
    
    
    ShowMessage pFunctionName & vbTab & pLowerBound & vbTab & pUpperBound
    
    'initialize with an error value
    SolveX = CVErr(xlErrValue)
    
    xa = pLowerBound
    xb = pUpperBound
    
    If xa >= xb Then
        TempStr = "  error: invalid range: xa (" & xa & ") >= xb (" & xb & ")"
        ShowMessage TempStr
        Exit Function
    End If
    
    fxa = RunFunctionX(xa)
    If Abs(fxa) < pAccuracy Then
        SolveX = xa
        ShowMessage "  SolveX = xa = " & xa
        Exit Function
    End If
    
    fxb = RunFunctionX(xb)
    If Abs(fxb) < pAccuracy Then
        SolveX = xb
        ShowMessage "  SolveX = xb = " & xb
        Exit Function
    End If
    
    If fxa * fxb >= 0 Then
        TempStr = "  error: root not bracketed: f[" & xa & ";" & xb & "] -> [" & fxa & ";" & fxb & "]"
        ShowMessage TempStr
        Exit Function
    End If
    If pGuess <= xa Then
        TempStr = "  error: guess (" & pGuess & ") < xa (" & xa & ")"
        ShowMessage TempStr
        Exit Function
    End If
    If pGuess >= xb Then
        TempStr = "  error: guess (" & pGuess & ") > xb (" & xb & ")"
        ShowMessage TempStr
        Exit Function
    End If
    
    soln = pGuess
    
    
    Dim fsoln As Double
    Dim p As Double
    Dim q As Double
    Dim r As Double
    Dim s As Double
    Dim xTol1 As Double
    Dim xMid As Double
    Dim d As Double
    
    '==========================================================================
    Const Eps As Double = 2 ^ -52
    '==========================================================================
    
    d = 0
    e = 0
    
    Iter = 2
    soln = xb
    fsoln = fxb
    
    Do
        If (fsoln > 0 And fxb > 0) Or (fsoln < 0 And fxb < 0) Then
            xb = xa
            fxb = fxa
            d = soln - xa
            e = d
        End If
        If Abs(fxb) < Abs(fsoln) Then
            xa = soln
            soln = xb
            xb = xa
            fxa = fsoln
            fsoln = fxb
            fxb = fxa
        End If
        
        xTol1 = 2 * Eps * Abs(soln) + 0.5 * pAccuracy
        xMid = 0.5 * (xb - soln)
        
        'exit if function has converged
        If Abs(xMid) <= xTol1 Or fsoln = 0 Then
            SolveX = soln
            ShowMessage "  SolveX(" & Iter & ") = " & soln _
                    & vbTab & "(" & fsoln & ")"
            Exit Function
        End If
        
        If Abs(e) >= xTol1 And Abs(fxa) > Abs(fsoln) Then
            'inverse quadratic interpolation
            s = fsoln / fxa
            If xa = xb Then
                p = 2 * xMid * s
                q = 1 - s
            Else
                q = fxa / fxb
                r = fsoln / fxb
                p = s * (2 * xMid * q * (q - r) - (soln - xa) * (r - 1))
                q = (q - 1) * (r - 1) * (s - 1)
            End If
            If p > 0 Then
                q = -q
            End If
            p = Abs(p)
            
            If p < 0.5 * Application.WorksheetFunction.Min( _
                    3 * xMid * q - Abs(xTol1 * q), _
                    Abs(e * q) _
            ) Then
                'secant method is fast so consider interpolation
                e = d
                d = p / q
            Else
                'bisection step is considered since the secant method is not progressing fast
                d = xMid
                e = d
            End If
        Else
            'bounds decreasing too slowly, use bisection
            d = xMid
            e = d
        End If
        
        xa = soln
        fxa = fsoln
        
        If Abs(d) > xTol1 Then
            soln = soln + d
        Else
            soln = soln + sign(xTol1, xMid)
        End If
        
        fsoln = RunFunctionX(soln)
        
        Iter = Iter + 1
    Loop While Iter <= pMaxIterations
    
End Function


'modified function (originally 'RunFunction') which can handle an optional
'number of additional arguments that can be stored in 'pArr'
Private Function RunFunctionX( _
    x As Double _
        ) As Double
    
    If IsEmpty(pArr) Then
        If Not pObjectMode Then
            RunFunctionX = Application.Run(pFunctionName, x)
        Else
            RunFunctionX = CallByName(pCallbackObject, pFunctionName, VbMethod, x)
        End If
    Else
        Select Case UBound(pArr)
            Case 2
                If Not pObjectMode Then
                    RunFunctionX = Application.Run(pFunctionName, _
                            x, pArr(0), pArr(1), pArr(2))
                Else
                    RunFunctionX = CallByName(pCallbackObject, pFunctionName, _
                            VbMethod, x, pArr(0), pArr(1), pArr(2))
                End If
            Case 5
                If Not pObjectMode Then
                    RunFunctionX = Application.Run(pFunctionName, _
                            x, pArr(0), pArr(1), pArr(2), pArr(3), pArr(4), pArr(5))
                Else
                    RunFunctionX = CallByName(pCallbackObject, pFunctionName, _
                            VbMethod, x, pArr(0), pArr(1), pArr(2), pArr(3), pArr(4), pArr(5))
                End If
            Case 7
                If Not pObjectMode Then
                    RunFunctionX = Application.Run(pFunctionName, _
                            x, pArr(0), pArr(1), pArr(2), pArr(3), pArr(4), pArr(5), _
                            pArr(6), pArr(7))
                Else
                    RunFunctionX = CallByName(pCallbackObject, pFunctionName, _
                            VbMethod, x, pArr(0), pArr(1), pArr(2), pArr(3), pArr(4), pArr(5), _
                            pArr(6), pArr(7))
                End If
        End Select
    End If
    
End Function

'==============================================================================

Private Sub ShowMessage(Msg As String)
    If Not pSilent Then
        Debug.Print Msg
    End If
End Sub


Public Sub AssignObject(CallbackObject As Object)
    pObjectMode = True
    Set pCallbackObject = CallbackObject
End Sub


Private Function sign(a As Double, b As Double) As Double
    If b >= 0 Then
        sign = Abs(a)
    Else
        sign = -Abs(a)
    End If
End Function
