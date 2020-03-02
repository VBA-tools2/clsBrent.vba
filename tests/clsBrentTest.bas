Attribute VB_Name = "clsBrentTest"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
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
Public Sub SolveX_RootNotBracketed_ReturnsError()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Variant
    Dim sThisWorkbook As String
    
    Dim Brent As clsBrent
    Set Brent = New clsBrent
    
    '==========================================================================
    Const a As Double = 1
    Const x0 As Double = 1
    Const y0 As Double = -1
    
    Dim Expected As Variant
    Expected = CVErr(xlErrValue)
    '==========================================================================
    
    Dim Arr(0 To 2) As Variant
    
    
    Arr(0) = a
    Arr(1) = x0
    Arr(2) = y0
    
    'Act:
    sThisWorkbook = "'" & ThisWorkbook.Name & "'!"
    
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
Public Sub SolveX_PropertiesToFindLeftRoot_ReturnsLeftRoot()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Variant
    Dim Expected As Variant
    Dim sThisWorkbook As String
    
    Dim Brent As clsBrent
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
    sThisWorkbook = "'" & ThisWorkbook.Name & "'!"
    
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
