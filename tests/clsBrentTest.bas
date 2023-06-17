Attribute VB_Name = "clsBrentTest"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider

Private Actual As Variant
Private sThisWorkbook As String
Private Brent As clsBrent
Private Expected As Variant


'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    sThisWorkbook = "'" & ThisWorkbook.Name & "'!"
End Sub


'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub


'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub


'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'==============================================================================
'small helper function
'the analytical solution will most likely a bit different than the numerical solution
'(the given 'Accuracy' corresponds to the default value of 'clsBrent')
Private Function RealEqual( _
    ByVal x As Double, _
    ByVal y As Double, _
    Optional ByVal Accuracy As Double = 0.00001 _
        ) As Boolean
    RealEqual = Abs(x - y) <= Accuracy
End Function


'==============================================================================
'@TestMethod("SolveX")
Public Sub SolveX_MissingLowerBoundAndUpperBoundAndFunctionName_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.ePrerequisitesNotMet, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_MissingLowerBoundAndUpperBound_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.ePrerequisitesNotMet, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_MissingLowerBoundAndFunctionName_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .UpperBound = -1
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.ePrerequisitesNotMet, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_MissingUpperBoundAndFunctionName_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .LowerBound = -2
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.ePrerequisitesNotMet, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_MissingLowerBound_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .UpperBound = -1
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.ePrerequisitesNotMet, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_MissingUpperBound_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .LowerBound = -2
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.ePrerequisitesNotMet, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_MissingFunctionName_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .LowerBound = -2
        .UpperBound = -1
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.ePrerequisitesNotMet, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Private Sub SolveX_NotExistingFunctionName_ThrowsError()
    Const ExpectedError As Long = 1004
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    'Act:
    With Brent
        .LowerBound = 0
        .UpperBound = 1
        .FunctionName = "NotExistingFunctionName"
        Actual = .Solve
    End With
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_LowerBoundLargerThanUpperBound_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .LowerBound = -1
        .UpperBound = -2
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eLowerBoundGreaterUpperBound, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_RootNotBracketed_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .LowerBound = -2
        .UpperBound = -1
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eNotBracketed, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_GuessSmallerThanLowerBound_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .Guess = -3
        .LowerBound = -2
        .UpperBound = 1
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eGuessSmallerLowerBound, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_GuessGreaterThanUpperBound_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .Guess = 2
        .LowerBound = -2
        .UpperBound = 1
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eGuessGreaterUpperBound, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_MaxIterations_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .Guess = -10
        .LowerBound = -10
        .UpperBound = 1
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .MaxIterations = 10
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eMaxIterations, Brent.Status, "Test of Status"
    Assert.AreEqual Expected, Actual, "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_LowerBoundIsValidSolution_ReturnsLeftRoot()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    Expected = Parabola_Vertex_LeftRoot(a, x0, y0)
    With Brent
        .LowerBound = 0
        .UpperBound = 5
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eNoError, Brent.Status, "Test of Status"
    Assert.IsTrue RealEqual(Expected, Actual), "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_UpperBoundIsValidSolution_ReturnsLeftRoot()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    Expected = Parabola_Vertex_LeftRoot(a, x0, y0)
    With Brent
        .LowerBound = -5
        .UpperBound = 0
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eNoError, Brent.Status, "Test of Status"
    Assert.IsTrue RealEqual(Expected, Actual), "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_PropertiesToFindLeftRoot_ReturnsLeftRoot()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    Expected = Parabola_Vertex_LeftRoot(a, x0, y0)
    With Brent
        .Guess = -1.5
        .LowerBound = -2
        .UpperBound = 1
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eNoError, Brent.Status, "Test of Status"
    Assert.IsTrue RealEqual(Expected, Actual), "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_ArrWithLowerBoundGreaterZero_ReturnsLeftRoot()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    '==========================================================================
    
    Dim Arr(5 To 7) As Variant
    Arr(5) = a
    Arr(6) = x0
    Arr(7) = y0
    
    'Act:
    Expected = Parabola_Vertex_LeftRoot(a, x0, y0)
    With Brent
        .Guess = -1.5
        .LowerBound = -2
        .UpperBound = 1
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = Arr
        Actual = .Solve
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eNoError, Brent.Status, "Test of Status"
    Assert.IsTrue RealEqual(Expected, Actual), "Test of Result Value"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_ArrWithLowerBoundGreaterZero_ReturnsArr()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    '==========================================================================
    
    Dim Arr(5 To 7) As Variant
    Arr(5) = a
    Arr(6) = x0
    Arr(7) = y0
    
    'Act:
    With Brent
        .Arr = Arr
        
        Dim ReturnArr As Variant
        ReturnArr = .Arr
    End With
    
    'Assert:
    Assert.SequenceEquals Arr, ReturnArr
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_ArrWithLowerBoundQualZero_ReturnsArr()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    With Brent
        .Arr = Arr
        
        Dim ReturnArr As Variant
        ReturnArr = .Arr
    End With
    
    'Assert:
    Assert.SequenceEquals Arr, ReturnArr
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SolveX")
Public Sub SolveX_ScalarArrEntry_ReturnsLeftRoot()
    On Error GoTo TestFail
    
    'Arrange:
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    '==========================================================================
    
    'Act:
    Expected = Parabola_Vertex_LeftRoot(a)
    With Brent
        .Guess = -5
        .LowerBound = -6
        .UpperBound = -1
        .FunctionName = sThisWorkbook & "Parabola_Vertex"
        .Arr = a
        Actual = .Solve
        
        Dim ReturnArr As Double
        ReturnArr = .Arr
    End With
    
    'Assert:
    Assert.AreEqual eBrentStatus.eNoError, Brent.Status, "Test of Status"
    Assert.IsTrue RealEqual(Expected, Actual), "Test of Result Value"
    Assert.AreEqual a, ReturnArr, "Test of ReturnArr"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
