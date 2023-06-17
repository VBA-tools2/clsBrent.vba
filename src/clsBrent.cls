VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBrent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'@Exposed

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

Private Type TBrent
    MaxIterations As Long
    ObjectMode As Boolean
    CallbackObject As Object
    Silent As Boolean
    Status As eBrentStatus
    Accuracy As Double
    Guess As Double
    LowerBound As Double
    UpperBound As Double
    FunctionName As String
    Arr As Variant
    
    bFunctionNameSet As Boolean
    bLowerBoundSet As Boolean
    bUpperBoundSet As Boolean
    bGuessSet As Boolean
    
    bScalarArr As Boolean
    bCustomArrBounds As Boolean
    LowerArrBound As Long
    UpperArrBound As Long
    
    MaxArgs As Long
End Type
Private This As TBrent


Public Enum eBrentStatus
    ePrerequisitesNotMet = -2
    ePrerequisitesMet
    eNoError
    eLowerBoundGreaterUpperBound
    eNotBracketed
    eGuessSmallerLowerBound
    eGuessGreaterUpperBound
    eMaxIterations
End Enum


Public Enum eBrentError
    [_First] = vbObjectError + 1
    ErrMoreThanMaxArgs
    [_Last] = ErrMoreThanMaxArgs
End Enum


Public Property Get MaxIterations() As Long
    MaxIterations = This.MaxIterations
End Property

Public Property Let MaxIterations(ByVal Value As Long)
    This.MaxIterations = Value
End Property


Public Property Get CallbackObject() As Object
    Set CallbackObject = This.CallbackObject
End Property

Public Property Let CallbackObject(ByVal Value As Object)
    Set This.CallbackObject = Value
End Property


Public Property Get Status() As eBrentStatus
    Status = This.Status
End Property


Public Property Get Silent() As Boolean
    Silent = This.Silent
End Property

Public Property Let Silent(ByVal Value As Boolean)
    This.Silent = Value
End Property


Public Property Get Accuracy() As Double
    Accuracy = This.Accuracy
End Property

Public Property Let Accuracy(ByVal Value As Double)
    This.Accuracy = Value
End Property


Public Property Get Guess() As Double
    Guess = This.Guess
End Property

Public Property Let Guess(ByVal Value As Double)
    With This
        .Guess = Value
        .bGuessSet = True
    End With
End Property


Public Property Get LowerBound() As Double
    LowerBound = This.LowerBound
End Property

Public Property Let LowerBound(ByVal Value As Double)
    With This
        .LowerBound = Value
        .bLowerBoundSet = True
    End With
    CheckPrerequisites
End Property


Public Property Get UpperBound() As Double
    UpperBound = This.UpperBound
End Property

Public Property Let UpperBound(ByVal Value As Double)
    With This
        .UpperBound = Value
        .bUpperBoundSet = True
    End With
    CheckPrerequisites
End Property


Public Property Get Arr() As Variant
    With This
        If .bScalarArr Then
            Dim ScalarArr As Variant
            ScalarArr = .Arr(0)
            
            Arr = ScalarArr
        ElseIf .bCustomArrBounds Then
            Dim ResultArr As Variant
            ReDim ResultArr(.LowerArrBound To .UpperArrBound)
            
            Dim i As Long
            For i = 0 To .UpperArrBound - .LowerArrBound
                ResultArr(.LowerArrBound + i) = .Arr(i)
            Next
            Arr = ResultArr
        Else
            Arr = .Arr
        End If
    End With
End Property

Public Property Let Arr(ByVal Value As Variant)
    With This
        If Not IsArray(Value) Then
            .bScalarArr = True
            ReDim .Arr(0 To 0)
            .Arr(0) = Value
        ElseIf LBound(Value) = 0 Then
            If UBound(Value) >= .MaxArgs Then RaiseErrorMoreThanMaxArgs
            
            .Arr = Value
        Else
            .bCustomArrBounds = True
            .LowerArrBound = LBound(Value)
            .UpperArrBound = UBound(Value)
            
            If .UpperArrBound - .LowerArrBound + 1 > .MaxArgs Then RaiseErrorMoreThanMaxArgs
            
            ReDim .Arr(0 To .UpperArrBound - .LowerArrBound)
            
            Dim i As Long
            For i = 0 To .UpperArrBound - .LowerArrBound
                .Arr(i) = Value(.LowerArrBound + i)
            Next
        End If
    End With
End Property


Public Property Get FunctionName() As String
    FunctionName = This.FunctionName
End Property

Public Property Let FunctionName(ByVal Value As String)
    With This
        .FunctionName = Value
        .bFunctionNameSet = True
    End With
    CheckPrerequisites
End Property


Public Property Get Solve() As Variant
    'initialize with an error value
    Solve = CVErr(xlErrValue)
    
    With This
        'check if all prerequisites are fulfilled
        If .Status < ePrerequisitesMet Then Exit Property
        If Not .bGuessSet Then _
                .Guess = Application.WorksheetFunction.Average( _
                        .LowerBound, .UpperBound _
                )
        .Status = eNoError
    End With
    
    'call function
    Solve = SolveX()
End Property


'------------------------------------------------------------------------------

Private Sub Class_Initialize()
    With This
        .MaxIterations = 100
        .ObjectMode = False
        .Silent = True
        .Status = ePrerequisitesNotMet
        
        .Accuracy = 0.00001
        
        '======================================================================
        'maximum allowed number of 'Arr' arguments
        '(keep in sync with number of '.Arr' arguments in 'RunFunctionX')
        .MaxArgs = 10
        '======================================================================
    End With
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
    
    
    With This
        ShowMessage .FunctionName & vbTab & .LowerBound & vbTab & .UpperBound
        
        'initialize with an error value
        SolveX = CVErr(xlErrValue)
        
        xa = .LowerBound
        xb = .UpperBound
        
        If xa >= xb Then
            .Status = eLowerBoundGreaterUpperBound
            TempStr = "  error: invalid range: xa (" & xa & ") >= xb (" & xb & ")"
            ShowMessage TempStr
            Exit Function
        End If
        
        fxa = RunFunctionX(xa)
        If Abs(fxa) < .Accuracy Then
            .Status = eNoError
            SolveX = xa
            ShowMessage "  SolveX = xa = " & xa
            Exit Function
        End If
        
        fxb = RunFunctionX(xb)
        If Abs(fxb) < .Accuracy Then
            .Status = eNoError
            SolveX = xb
            ShowMessage "  SolveX = xb = " & xb
            Exit Function
        End If
        
        If fxa * fxb >= 0 Then
            .Status = eNotBracketed
            TempStr = "  error: root not bracketed: f[" & xa & ";" & xb & "] -> [" & fxa & ";" & fxb & "]"
            ShowMessage TempStr
            Exit Function
        End If
        If .Guess < xa Then
            .Status = eGuessSmallerLowerBound
            TempStr = "  error: guess (" & .Guess & ") < xa (" & xa & ")"
            ShowMessage TempStr
            Exit Function
        End If
        If .Guess > xb Then
            .Status = eGuessGreaterUpperBound
            TempStr = "  error: guess (" & .Guess & ") > xb (" & xb & ")"
            ShowMessage TempStr
            Exit Function
        End If
        
        Dim fsoln As Double
        Dim p As Double
        Dim q As Double
        Dim r As Double
        Dim s As Double
        Dim xTol1 As Double
        Dim xMid As Double
        Dim d As Double
        
        '======================================================================
        Const Eps As Double = 2 ^ -52
        '======================================================================
        
        d = 0
        e = 0
        
        Iter = 2
        soln = .Guess
        fsoln = RunFunctionX(soln)
        
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
            
            xTol1 = 2 * Eps * Abs(soln) + 0.5 * .Accuracy
            xMid = 0.5 * (xb - soln)
            
            'exit if function has converged
            If Abs(xMid) <= xTol1 Or fsoln = 0 Then
                .Status = eNoError
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
        Loop While Iter <= .MaxIterations
        
        .Status = eMaxIterations
        TempStr = "  error: max iterations = " & .MaxIterations & " reached."
        ShowMessage TempStr
    End With
    
End Function


'modified function (originally 'RunFunction') which can handle an optional
'number of additional arguments that can be stored in 'Arr'
Private Function RunFunctionX( _
    ByVal x As Double _
        ) As Double
    
    With This
        '(otherwise it is hard to create/set '.Arr' in 'ResizeArr')
        If IsEmpty(.Arr) Then
            If Not .ObjectMode Then
                RunFunctionX = Application.Run(.FunctionName, x)
            Else
                RunFunctionX = CallByName(.CallbackObject, .FunctionName, VbMethod, x)
            End If
        Else
            ResizeArr
            
            If Not .ObjectMode Then
                RunFunctionX = Application.Run(.FunctionName, x, _
                        .Arr(0), .Arr(1), .Arr(2), .Arr(3), .Arr(4), _
                        .Arr(5), .Arr(6), .Arr(7), .Arr(8), .Arr(9))
            Else
                RunFunctionX = CallByName(.CallbackObject, .FunctionName, VbMethod, x, _
                        .Arr(0), .Arr(1), .Arr(2), .Arr(3), .Arr(4), _
                        .Arr(5), .Arr(6), .Arr(7), .Arr(8), .Arr(9))
            End If
        End If
    End With
    
End Function


'REF: <https://github.com/cristianbuse/Excel-VBA-ProgressBar/blob/c0f862c0f904ac0b9992eba6dee7734ccc7f66fe/src/ProgressBar.cls#L544-L552>
Private Sub ResizeArr(Optional ByVal Missing As Variant)
    
    With This
        If IsMissing(.Arr) Then Exit Sub
        
        Dim ArgCount As Long
        ArgCount = UBound(.Arr)
        
        ReDim Preserve .Arr(0 To .MaxArgs - 1)
        
        Dim i As Long
        For i = ArgCount + 1 To UBound(.Arr)
            .Arr(i) = Missing
        Next
    End With
    
End Sub


Private Sub CheckPrerequisites()
    With This
        If .bLowerBoundSet And _
                .bUpperBoundSet And _
                .bFunctionNameSet _
        Then
            .Status = ePrerequisitesMet
        End If
    End With
End Sub


'==============================================================================

Private Sub ShowMessage(ByVal Msg As String)
    If Not This.Silent Then
        Debug.Print Msg
    End If
End Sub


Public Sub AssignObject(ByVal CallbackObject As Object)
    With This
        .ObjectMode = True
        Set .CallbackObject = CallbackObject
    End With
End Sub


Private Function sign(ByVal a As Double, ByVal b As Double) As Double
    If b >= 0 Then
        sign = Abs(a)
    Else
        sign = -Abs(a)
    End If
End Function


'==============================================================================
Private Sub RaiseErrorMoreThanMaxArgs()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eBrentError.ErrMoreThanMaxArgs, _
            Description:= _
                    "Currently only " & This.MaxArgs & " additional " & _
                    "arguments are allowed/supported."
End Sub
