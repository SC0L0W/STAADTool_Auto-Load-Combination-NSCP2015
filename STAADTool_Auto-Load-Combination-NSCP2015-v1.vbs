Option Explicit

' Helper: Check if load case exists by title
Function LoadCaseExistsByTitle(Title As String, Func As Object) As Boolean
    Dim i As Long
    LoadCaseExistsByTitle = False
    On Error Resume Next
    For i = 1 To Func.Load.GetPrimaryLoadCaseCount
        If StrComp(Func.Load.GetLoadCaseTitle(i), Title, vbTextCompare) = 0 Then
            LoadCaseExistsByTitle = True
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function

' Helper: Get load case number by title
Function GetLoadCaseNumberByTitle(Title As String, Func As Object) As Long
    Dim i As Long
    GetLoadCaseNumberByTitle = 0
    On Error Resume Next
    For i = 1 To Func.Load.GetPrimaryLoadCaseCount
        If StrComp(Func.Load.GetLoadCaseTitle(i), Title, vbTextCompare) = 0 Then
            GetLoadCaseNumberByTitle = i
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function

' Helper: Add load case to combo only if exists
Sub AddLoadCaseToCombinationIfExists(Title As String, CombNum As Long, Factor As Double, Func As Object)
    If LoadCaseExistsByTitle(Title, Func) Then
        Dim LCNum As Long
        LCNum = GetLoadCaseNumberByTitle(Title, Func)
        If LCNum > 0 Then
            Dim res As Variant
            res = Func.Load.AddLoadAndFactorToCombination(CombNum, LCNum, Factor)
            If res <> 0 Then
                Debug.Print "Warning: Failed to add load case '" & Title & "' (LC " & LCNum & ")"
            End If
        End If
    Else
        Debug.Print "Warning: Load case with title '" & Title & "' not found. Skipping."
    End If
End Sub

' Helper: Create load combination (with skipping missing load cases)
Sub CreateCombinationSafe(Title As String, CombNum As Long, FactorsAndTitles As Variant, Func As Object)
    ' Create the combination
    Dim res As Variant
    On Error Resume Next
    Func.Load.DeleteLoadCombination CombNum
    res = Func.Load.CreateNewLoadCombination(Title, CombNum)
    If res = -1 Then
        Debug.Print "Error: Could not create combination " & CombNum
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Add each load case if exists
    Dim i As Integer
    For i = LBound(FactorsAndTitles) To UBound(FactorsAndTitles) Step 2
        Dim Factor As Double
        Dim TitleStr As String
        Factor = FactorsAndTitles(i)
        TitleStr = FactorsAndTitles(i + 1)
        AddLoadCaseToCombinationIfExists(TitleStr, CombNum, Factor, Func)
    Next
End Sub

' Main subroutine
Sub Main()
    Debug.Clear
    
    Dim Func As Object
    On Error GoTo ConnectionError
    Set Func = GetObject(, "StaadPro.OpenSTAAD")
    ' Test connection
    Dim TestVar As Long
    TestVar = Func.Geometry.GetMemberCount()
    On Error GoTo 0
    
    ' Define parameters
    Dim i As Integer
    Debug.Print "Starting NSCP 2015 Load Combination Generator..."
    Debug.Print "Multiple Combination Varieties"
    
    ' Get Load Cases
    Dim TotalLoadCase As Long
    TotalLoadCase = Func.Load.GetPrimaryLoadCaseCount()
    Debug.Print "Total Load Cases: "; TotalLoadCase
    
    If TotalLoadCase = 0 Then
        MsgBox "No load cases found! Please define load cases first.", vbExclamation, "No Load Cases"
        Exit Sub
    End If
    
    Dim LoadCaseNum() As Long
    ReDim LoadCaseNum(TotalLoadCase - 1)
    Func.Load.GetPrimaryLoadCaseNumbers(LoadCaseNum)
    
    Dim LoadCaseTitle() As String
    ReDim LoadCaseTitle(TotalLoadCase - 1)
    For i = 0 To TotalLoadCase - 1
        LoadCaseTitle(i) = Func.Load.GetLoadCaseTitle(LoadCaseNum(i))
        Debug.Print "Load Case "; LoadCaseNum(i); ": "; LoadCaseTitle(i)
    Next
    
    ' Classification based on title
    Dim DeadLoads1 As Long, DeadLoads2 As Long
    Dim LiveLoads As Long, LL1Loads As Long, LL2Loads As Long
    Dim RoofLoads As Long
    Dim EqX As Long, EqZ As Long
    Dim RSX As Long, RSZ As Long
    Dim WindX As Long, WindZ As Long
    
    DeadLoads1 = 0
    DeadLoads2 = 0
    LiveLoads = 0
    LL1Loads = 0
    LL2Loads = 0
    RoofLoads = 0
    EqX = 0
    EqZ = 0
    RSX = 0
    RSZ = 0
    WindX = 0
    WindZ = 0
    
    Dim LoadTitle As String
    Dim LoadCaseNum_i As Long
    
    For i = 0 To TotalLoadCase - 1
        LoadCaseNum_i = LoadCaseNum(i)
        LoadTitle = UCase(Trim(Func.Load.GetLoadCaseTitle(LoadCaseNum_i)))
        If InStr(LoadTitle, "DL1") > 0 Then
            DeadLoads1 = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "DL2") > 0 Then
            DeadLoads2 = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "LL1") > 0 Then
            LL1Loads = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "LL2") > 0 Then
            LL2Loads = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "LL") > 0 Then
            If InStr(LoadTitle, "LR") > 0 Or InStr(LoadTitle, "RL") > 0 Or InStr(LoadTitle, "RFL") > 0 Then
                RoofLoads = LoadCaseNum_i
            Else
                LiveLoads = LoadCaseNum_i
            End If
        ElseIf InStr(LoadTitle, "EX") > 0 Then
            EqX = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "EZ") > 0 Then
            EqZ = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "RSX") > 0 Then
            RSX = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "RSZ") > 0 Then
            RSZ = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "WX") > 0 Then
            WindX = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "WZ") > 0 Then
            WindZ = LoadCaseNum_i
        End If
    Next
    
    ' Design method selection
    Dim MethodName As String
    Dim EvValue As Double
    Dim LRFD_Start101 As Long, LRFD_Start201 As Long, LRFD_Start301 As Long, LRFD_Start401 As Long
    Dim ASD_Start501 As Long, ASD_Start601 As Long
    LRFD_Start101 = 101
    LRFD_Start201 = 201
    LRFD_Start301 = 301
    LRFD_Start401 = 401
    ASD_Start501 = 501
    ASD_Start601 = 601
    
    If MsgBox("Select Design Method:" & vbCrLf & vbCrLf & _
              "YES = LRFD (Strength Design)" & vbCrLf & _
              "         Generate 4 series (101, 201, 301, 401)" & vbCrLf & vbCrLf & _
              "NO = ASD (Allowable Stress Design)" & vbCrLf & _
              "         Generate 501+ and 601+ series", _
              vbYesNo + vbQuestion, "NSCP 2015 Design Method") = vbYes Then
        ' LRFD path
        MethodName = "LRFD"
        Dim EvInput As String
        EvInput = InputBox("Enter Ev factor for Series 301 & 401 (LRFD with Ev):" & vbCrLf & vbCrLf & _
                           "Typical values:" & vbCrLf & _
                           "Factor = 0.27 (when Ev = 0.5 * Ca * I * D, Ca=0.532, I=1.0)" & vbCrLf & _
                           "0.20 (for lower seismic zones)" & vbCrLf & vbCrLf & _
                           "Enter value:", "Ev Factor for Series 301", "0.20")
        If EvInput = "" Then
            MsgBox "Operation cancelled.", vbInformation, "Cancelled"
            Exit Sub
        End If
        If Not IsNumeric(EvInput) Then
            MsgBox "Invalid Ev value entered!", vbExclamation, "Invalid Input"
            Exit Sub
        End If
        EvValue = CDbl(EvInput)
        
        ' Generate LRFD series
        Call GenerateLRFD_Series101(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, LRFD_Start101)
        LRFD_Start101 = LRFD_Start101 + 50
        Call GenerateLRFD_Series201(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, LRFD_Start201)
        LRFD_Start201 = LRFD_Start201 + 50
        Call GenerateLRFD_Series301(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, LRFD_Start301, EvValue)
        LRFD_Start301 = LRFD_Start301 + 50
        Call GenerateLRFD_Series401(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, RSX, RSZ, WindX, WindZ, LRFD_Start401, EvValue)
        MsgBox "LRFD combinations generated successfully!", vbInformation
    Else
        ' ASD path
        MethodName = "ASD"
        Call GenerateASD_Basic(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, ASD_Start501)
        ASD_Start501 = ASD_Start501 + 50
        Call GenerateASD_Alternate(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, ASD_Start601)
        ASD_Start601 = ASD_Start601 + 50
        MsgBox "ASD combinations generated successfully!", vbInformation
    End If
    
    Exit Sub
    
ConnectionError:
    MsgBox "Cannot connect to STAAD.Pro!" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Connection Error"
End Sub

'=== LRFD SERIES 101: Basic LRFD (No orthogonal, No Ev) ===
Sub GenerateLRFD_Series101( _
    Func As Object, _
    DeadLoads1 As Long, _
    DeadLoads2 As Long, _
    LiveLoads As Long, _
    LL1Loads As Long, _
    LL2Loads As Long, _
    RoofLoads As Long, _
    EqX As Long, _
    EqZ As Long, _
    WindX As Long, _
    WindZ As Long, _
    ByRef CombStart As Long _
)
    Debug.Print vbCrLf & "--- Series 101: Basic LRFD ---"
    
    Dim CombNum As Long
    Dim CombTitle As String
    Dim CombString As String
    
    ' Helper: Create load combination
    ' Assumes CreateLoadCombination is defined elsewhere
    
    ' 1.4D
    CombNum = CombStart
    CombTitle = "LRFD-101: 1.4DL"
    CombString = " 1.4 " & DeadLoads1 & " 1.4 " & DeadLoads2
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6L + 1.6LL + 1.6LL1 + 1.6LL2 + 0.5LLR
    CombNum = CombStart
    CombTitle = "LRFD-102: 1.2DL + 1.6LL + 1.6LL1 + 1.6LL2 + 0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 1.6 " & LiveLoads & " 1.6 " & LL1Loads & " 1.6 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5LL + 0.5LL1 + LL2 + 1.6LLR
    CombNum = CombStart
    CombTitle = "LRFD-103: 1.2D + 0.5LL + 0.5LL1 + LL2 + 1.6Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-104: 1.2D + 0.5LL + 0.5LL1 + LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR + 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-105: 1.2D + 1.6LLR + Wx"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR -0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-106: 1.2D + 1.6LLR - 0.5Wx"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR +0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-107: 1.2D + 1.6LLR + 0.5Wz"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR -0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-108: 1.2D + 1.6LLR - 0.5Wz"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-109: 1.2D + Ex + 0.5L + 0.5LL1 + LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-110: 1.2D - Ex +0.5L + 0.5LL1 + LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-111: 1.2D + Ez + 0.5L + 0.5LL1 + LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-112: 1.2D - Ez + 0.5L + 0.5LL1 + LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-113: 0.9D + Ex + 0.5L + 0.5LL1 + LL2"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-114: 0.9D - Ex + 0.5L + 0.5LL1 + LL2"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-115: 0.9D + Ez + 0.5L + 0.5LL1 + LL2"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-116: 0.9D - Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-117: 1.2D + Wx + 0.5LL1 + LL2 + 0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-118: 1.2D - Wx - 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-119: 0.9D + Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-120: 0.9D - Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 1.2D + Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-121: 1.2D + Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-122: 1.2D - Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-123: 0.9D + Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-124: 0.9D - Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' Note: You can add further combinations as needed, possibly recursive calls or sequence continuation

    Debug.Print "Series 101 Complete: " & (CombStart - 1) & " combinations"
End Sub

'=== LRFD SERIES 201: LRFD with Orthogonal (No Ev) ===
Sub GenerateLRFD_Series201( _
    Func As Object, _
    DeadLoads1 As Long, _
    DeadLoads2 As Long, _
    LiveLoads As Long, _
    LL1Loads As Long, _
    LL2Loads As Long, _
    RoofLoads As Long, _
    EqX As Long, _
    EqZ As Long, _
    WindX As Long, _
    WindZ As Long, _
    ByRef CombStart As Long _
)
    Debug.Print vbCrLf & "--- Series 201: LRFD with Orthogonal ---"
    
    Dim CombNum As Long
    Dim CombTitle As String
    Dim CombString As String
    
    ' 1.4D
    CombNum = CombStart
    CombTitle = "LRFD-201: 1.4D"
    CombString = "1.4 " & DeadLoads1 & " 1.4 " & DeadLoads2
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6L + LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-202: 1.2D + 1.6L + 1.6LL1 + 1.6LL2 + 0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 1.6 " & LiveLoads & " 1.6 " & LL1Loads & " 1.6 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5LL + 0.5LL1 + LL2 + 1.6Roof
    CombNum = CombStart
    CombTitle = "LRFD-203: 1.2D + 0.5LL + 0.5LL1 + LL2 + 1.6Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-204: 1.2D + 0.5LL + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-205: 1.2D + 1.6LLR + 0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-206: 1.2D + 1.6LLR - 0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-207: 1.2D + 1.6LLR + 0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-208: 1.2D + 1.6LLR - 0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-209: 1.2D + Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " 0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-210: 1.2D + Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " -0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-211: 1.2D - Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " 0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-212: 1.2D - Ex - 0.3Ez + 0.5LL + 0.5LL1 +LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " -0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-213: 1.2D + Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " 0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-214: 1.2D + Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " -0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-215: 1.2D - Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " 0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Ez - 0.3Ex + 0.5LL + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-216: 1.2D - Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " -0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-217: 0.9D + Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " 0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-218: 0.9D + Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " -0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Ex + 0.3Ez + O.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-219: 0.9D - Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " 0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-220: 0.9D - Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " -0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-221: 0.9D + Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " 0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-222: 0.9D + Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " -0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-223: 0.9D - Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " 0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-224: 0.9D - Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " -0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-225: 1.2D + Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-226: 1.2D - Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-227: 1.2D + Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D - Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-228: 1.2D - Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.9D + Wx  + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-229: 0.9D + Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-230: 0.9D - Wx + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.9D + Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-231: 0.9D + Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D - Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-232: 0.9D - Wz + 0.5LL + 0.5LL1 + LL2 + 0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    
    Debug.Print "Series 201 Complete: " & (CombStart - 1) & " combinations"
End Sub

'=== LRFD SERIES 301: LRFD with Orthogonal + Ev ===
Sub GenerateLRFD_Series301( _
    Func As Object, _
    DeadLoads1 As Long, _
    DeadLoads2 As Long, _
    LiveLoads As Long, _
    LL1Loads As Long, _
    LL2Loads As Long, _
    RoofLoads As Long, _
    EqX As Long, _
    EqZ As Long, _
    WindX As Long, _
    WindZ As Long, _
    ByRef CombStart As Long, _
    EvValue As Double _
)
    Debug.Print vbCrLf & "--- Series 301: LRFD with Orthogonal + Ev ---"
    
    Dim CombNum As Long
    Dim CombTitle As String
    Dim CombString As String
    Dim D1 As Double, D2 As Double
    D1 = 1.2 + EvValue
    D2 = 1.2 - EvValue
    Dim D3 As Double, D4 As Double
    D3 = 0.9 + EvValue
    D4 = 0.9 - EvValue
    
    ' 1.4D
    CombNum = CombStart
    CombTitle = "LRFD-301: 1.4D"
    CombString = "1.4 " & DeadLoads1 & " 1.4 " & DeadLoads2
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6L + 1.6LL1 + 1.6LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-302: 1.2D + 1.6LL + 1.6LL1 + 1.6LL2 + 0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 1.6 " & LiveLoads & " 1.6 " & LL1Loads & " 1.6 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5LL + 0.5LL1 + LL2 + 1.6Roof
    CombNum = CombStart
    CombTitle = "LRFD-303: 1.2D + 0.5LL + 0.5LL1 + LL2 + 1.6Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-304: 1.2D + 0.5LLR + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-305: 1.2D + 1.6LLR + 0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-306: 1.2D + 1.6LLR - 0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-307: 1.2D + 1.6LLR + 0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-308: 1.2D + 1.6LLR - 0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' (1.2+Ev)D + Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-309: " & Format(D1, "0.00") & "D + Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D1, "0.00") & " " & DeadLoads1 & " " & _
        Format(D1, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqX & " " & _
        "0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2+Ev)D + Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-310: " & Format(D1, "0.00") & "D + Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D1, "0.00") & " " & DeadLoads1 & " " & _
        Format(D1, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqX & " " & _
        "-0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D - Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-311: " & Format(D2, "0.00") & "D - Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & EqX & " " & _
        "0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D - Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-312: " & Format(D2, "0.00") & "D - Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & EqX & " " & _
        "-0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2+Ev)D + Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-313: " & Format(D1, "0.00") & "D + Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D1, "0.00") & " " & DeadLoads1 & " " & _
        Format(D1, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqZ & " " & _
        "0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2+Ev)D + Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-314: " & Format(D1, "0.00") & "D + Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D1, "0.00") & " " & DeadLoads1 & " " & _
        Format(D1, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqZ & " " & _
        "-0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D - Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-315: " & Format(D2, "0.00") & "D - Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & EqZ & " " & _
        "0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D - Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-316: " & Format(D2, "0.00") & "D - Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & EqZ & " " & _
        "-0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-317: " & Format(D3, "0.00") & "D + Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D3, "0.00") & " " & DeadLoads1 & " " & _
        Format(D3, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqX & " " & _
        "0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-318: " & Format(D3, "0.00") & "D + Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D3, "0.00") & " " & DeadLoads1 & " " & _
        Format(D3, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqX & " " & _
        "-0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D - Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-319: " & Format(D4, "0.00") & "D - Ex + 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & EqX & " " & _
        "0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D - Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-320: " & Format(D4, "0.00") & "D - Ex - 0.3Ez + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & EqX & " " & _
        "-0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-321: " & Format(D3, "0.00") & "D + Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D3, "0.00") & " " & DeadLoads1 & " " & _
        Format(D3, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqZ & " " & _
        "0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-322: " & Format(D3, "0.00") & "D + Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D3, "0.00") & " " & DeadLoads1 & " " & _
        Format(D3, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqZ & " " & _
        "-0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D - Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-323: " & Format(D4, "0.00") & "D - Ez + 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & EqZ & " " & _
        "0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D - Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-324: " & Format(D4, "0.00") & " D - Ez - 0.3Ex + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & EqZ & " " & _
        "-0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    Debug.Print "Series 301 Complete: " & (CombStart - 1) & " combinations"
End Sub

'=== LRFD SERIES 401: LRFD with Orthogonal + Ev + RSA ===
Sub GenerateLRFD_Series401( _
    Func As Object, _
    DeadLoads1 As Long, _
    DeadLoads2 As Long, _
    LiveLoads As Long, _
    LL1Loads As Long, _
    LL2Loads As Long, _
    RoofLoads As Long, _
    RSX As Long, _
    RSZ As Long, _
    WindX As Long, _
    WindZ As Long, _
    ByRef CombStart As Long, _
    EvValue As Double _
)
    Debug.Print vbCrLf & "--- Series 301: LRFD with Orthogonal + Ev + RSA ---"
    
    Dim CombNum As Long
    Dim CombTitle As String
    Dim CombString As String
    Dim D1 As Double, D2 As Double
    D1 = 1.2 + EvValue
    D2 = 1.2 - EvValue
    Dim D3 As Double, D4 As Double
    D3 = 0.9 + EvValue
    D4 = 0.9 - EvValue
    
    ' 1.4D
    CombNum = CombStart
    CombTitle = "LRFD-401: 1.4D"
    CombString = "1.4 " & DeadLoads1 & " 1.4 " & DeadLoads2
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6L + 1.6LL1 + 1.6LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-402: 1.2D + 1.6LL + 1.6LL1 + 1.6LL2 + 0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 1.6 " & LiveLoads & " 1.6 " & LL1Loads & " 1.6 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5LL + 0.5LL1 + LL2 + 1.6Roof
    CombNum = CombStart
    CombTitle = "LRFD-403: 1.2D + 0.5LL + 0.5LL1 + LL2 + 1.6Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-404: 1.2D + 0.5LLR + 0.5LL1 + LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-405: 1.2D + 1.6LLR + 0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-406: 1.2D + 1.6LLR - 0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-407: 1.2D + 1.6LLR + 0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-408: 1.2D + 1.6LLR - 0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' (1.2+Ev)D + RSX + 0.3RSZ + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-409: " & Format(D1, "0.00") & "D + RSX + 0.3RSZ + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D1, "0.00") & " " & DeadLoads1 & " " & _
        Format(D1, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & RSX & " " & _
        "0.3 " & RSZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2+Ev)D + RSX - 0.3RSZ + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-410: " & Format(D1, "0.00") & "D + RSX - 0.3RSZ + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D1, "0.00") & " " & DeadLoads1 & " " & _
        Format(D1, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & RSX & " " & _
        "-0.3 " & RSZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D - RSX + 0.3RSZ + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-411: " & Format(D2, "0.00") & "D - RSX + 0.3RSZ + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & RSX & " " & _
        "0.3 " & RSZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D - RSX - 0.3RSZ + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-412: " & Format(D2, "0.00") & "D - RSX - 0.3RSZ + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & RSX & " " & _
        "-0.3 " & RSZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2+Ev)D + RSZ + 0.3RSX + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-413: " & Format(D1, "0.00") & "D + RSZ + 0.3RSX + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D1, "0.00") & " " & DeadLoads1 & " " & _
        Format(D1, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & RSZ & " " & _
        "0.3 " & RSX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2+Ev)D + RSZ - 0.3RSX + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-414: " & Format(D1, "0.00") & "D + RSZ - 0.3RSX + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D1, "0.00") & " " & DeadLoads1 & " " & _
        Format(D1, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & RSZ & " " & _
        "-0.3 " & RSX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D - RSZ + 0.3RSX + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-415: " & Format(D2, "0.00") & "D - RSZ + 0.3RSX + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & RSZ & " " & _
        "0.3 " & RSX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D - RSZ - 0.3RSX + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-416: " & Format(D2, "0.00") & "D - RSZ - 0.3RSX + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & RSZ & " " & _
        "-0.3 " & RSX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + RSX + 0.3RSZ + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-417: " & Format(D3, "0.00") & "D + RSX + 0.3RSZ + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D3, "0.00") & " " & DeadLoads1 & " " & _
        Format(D3, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & RSX & " " & _
        "0.3 " & RSZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + RSX - 0.3RSZ + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-418: " & Format(D3, "0.00") & "D + RSX - 0.3RSZ + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D3, "0.00") & " " & DeadLoads1 & " " & _
        Format(D3, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & RSX & " " & _
        "-0.3 " & RSZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D - RSX + 0.3RSZ + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-419: " & Format(D4, "0.00") & "D - RSX + 0.3RSZ + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & RSX & " " & _
        "0.3 " & RSZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D - RSX - 0.3RSZ + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-420: " & Format(D4, "0.00") & "D - RSX - 0.3RSZ + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & RSX & " " & _
        "-0.3 " & RSZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + RSZ + 0.3RSX + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-421: " & Format(D3, "0.00") & "D + RSZ + 0.3RSX + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D3, "0.00") & " " & DeadLoads1 & " " & _
        Format(D3, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & RSZ & " " & _
        "0.3 " & RSX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + RSZ - 0.3RSX + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-422: " & Format(D3, "0.00") & "D + RSZ - 0.3RSX + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D3, "0.00") & " " & DeadLoads1 & " " & _
        Format(D3, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & RSZ & " " & _
        "-0.3 " & RSX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D - RSZ + 0.3RSX + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-423: " & Format(D4, "0.00") & "D - RSZ + 0.3RSX + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & RSZ & " " & _
        "0.3 " & RSX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D - RSZ - 0.3RSX + 0.5LL + 0.5LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-424: " & Format(D4, "0.00") & " D - RSZ - 0.3RSX + 0.5LL + 0.5LL1 + LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "-1.0 " & RSZ & " " & _
        "-0.3 " & RSX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    Debug.Print "Series 401 Complete: " & (CombStart - 1) & " combinations"
End Sub


'=== ASD Basic: NSCP 203.4.1 ===
Sub GenerateASD_Basic(Func As Object, DeadLoads1 As Long, DeadLoads2 As Long, LiveLoads As Long, LL1Loads As Long, LL2Loads As Long, RoofLoads As Long, EqX As Long, EqZ As Long, WindX As Long, WindZ As Long, ByRef CombStart As Long)
    Debug.Print vbCrLf; "--- ASD Basic (NSCP 203.4.1) ---"
    Dim CombNum As Long
    Dim CombTitle As String
    Dim CombString As String
    
    ' D
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL1 + DL2"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + LL
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL1 + DL2 + LL + LL1 + LL2"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LLR
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL1 + DL2 + LLR"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + 0.75LL + 0.75LLR
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL1 + DL2 + LL + LL1 + LL2"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 0.75 " & LiveLoads & " 0.75 " & LL1Loads & " 0.75 " & LL2Loads & " 0.75 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + 0.715Ex + 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.715Ex + 0.215Ez"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 0.715 " & EqX & " 0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + 0.715Ex - 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.715Ex - 0.215Ez"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 0.715 " & EqX & " -0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D - 0.715Ex + 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.715Ex + 0.215Ez"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " -0.715 " & EqX & " 0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D - 0.715Ex - 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.715Ex - 0.215Ez"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " -0.715 " & EqX & " -0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + 0.715Ez + 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.715Ez + 0.215Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 0.715 " & EqZ & " 0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + 0.715Ez - 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.715Ez - 0.215Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 0.715 " & EqZ & " -0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D - 0.715Ez + 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D - 0.715Ez + 0.215Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " -0.715 " & EqZ & " 0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D - 0.715Ez - 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D - 0.715Ez - 0.215Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " -0.715 " & EqZ & " -0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + 0.6Wx
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.6Wx"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 0.6 " & WindX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D - 0.6Wx
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D - 0.6Wx"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " -0.6 " & WindX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + 0.6Wz
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.6Wz"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 0.6 " & WindZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D - 0.6Wz
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D - 0.6Wz"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " -0.6 " & WindZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    
    Debug.Print "ASD Basic Complete: " & (CombStart - 1) & " combinations"
End Sub

'=== ASD ALTERNATE: NSCP 203.4.2 ===
Sub GenerateASD_Alternate(Func As Object, DeadLoads1 As Long, DeadLoads2 As Long, LiveLoads As Long, LL1Loads As Long, LL2Loads As Long, RoofLoads As Long, EqX As Long, EqZ As Long, WindX As Long, WindZ As Long, ByRef CombStart As Long)
    Debug.Print vbCrLf; "--- ASD Alternate (NSCP 203.4.2) ---"
    Dim CombNum As Long
    Dim CombTitle As String
    Dim CombString As String
    
    ' DL + DL2
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL1 + DL2"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1




    ' D + LL + LL1 + LL2 + LLR (if exists)
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL + LL + LL1 + LL2 + LLR"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " 1 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 + 0.6Wx (if exists)
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL + LL + LL1 + LL2 + 0.6Wx"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " 0.6 " & WindX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 - 0.6Wx (if exists)
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL + LL + LL1 + LL2 - 0.6Wx"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " -0.6 " & WindX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 + 0.6Wz (if exists)
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL + LL + LL1 + LL2 + 0.6Wz"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " 0.6 " & WindZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 - 0.6Wz (if exists)
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": DL + LL + LL1 + LL2 - 0.6Wz"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " -0.6 " & WindZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 + 0.715Ex + 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + LL + LL1 + LL2 + 0.715Ex + 0.215Ez"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " 0.715 " & EqX & " 0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 + 0.715Ex - 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + LL + LL1 + LL2 + 0.715Ex - 0.215Ez"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " 0.715 " & EqX & " -0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 - 0.715Ex + 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + LL + LL1 + LL2 + 0.715Ex + 0.215Ez"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " -0.715 " & EqX & " 0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 - 0.715Ex - 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + LL + LL1 + LL2 + 0.715Ex - 0.215Ez"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " -0.715 " & EqX & " -0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 + 0.715Ez + 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + LL + LL1 + LL2 + 0.715Ez + 0.215Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " 0.715 " & EqZ & " 0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 + 0.715Ez - 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + LL + LL1 + LL2 + 0.715Ez - 0.215Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " 0.715 " & EqZ & " -0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 - 0.715Ez + 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + LL + LL1 + LL2 - 0.715Ez + 0.215Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " -0.715 " & EqZ & " 0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' D + LL + LL1 + LL2 - 0.715Ez - 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + LL + LL1 + LL2 - 0.715Ez - 0.215Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & DeadLoads2 & " 1.0 " & LiveLoads & " 1 " & LL1Loads & " 1 " & LL2Loads & " -0.715 " & EqZ & " -0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D + 0.6Wx
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D + 0.6Wx"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " 0.6 " & WindX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D - 0.6Wx
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D - 0.6Wx"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " -0.6 " & WindX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D + 0.6Wz
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D + 0.6Wz"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " 0.6 " & WindZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D - 0.6Wz
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D - 0.6Wz"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " -0.6 " & WindZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D + 0.715Ex + 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D + 0.715Ex + 0.215Ez"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " 0.715 " & EqX & " 0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D + 0.715Ex - 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D + 0.715Ex - 0.215Ez"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " 0.715 " & EqX & " -0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D - 0.715Ex + 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D + 0.715Ex + 0.215Ez"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " -0.715 " & EqX & " 0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D - 0.715Ex - 0.215Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D + 0.715Ex - 0.215Ez"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " -0.715 " & EqX & " -0.215 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D + 0.715Ez + 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.715Ez + 0.215Ex"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " 0.715 " & EqZ & " 0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D + 0.715Ez - 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D + 0.715Ez - 0.215Ex"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " 0.715 " & EqZ & " -0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D - 0.715Ez + 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D - 0.715Ez + 0.215Ex"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " -0.715 " & EqZ & " 0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 0.6D - 0.715Ez - 0.215Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D - 0.715Ez - 0.215Ex"
    CombString = "0.6 " & DeadLoads1 & " 0.6 " & DeadLoads2 & " -0.715 " & EqZ & " -0.215 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1


    
    Debug.Print "ASD Alternate Complete: " & (CombStart - 1) & " combinations"
End Sub

' Helper to append load case number
Function AppendLoad(CurrentLoads As String, LoadNum As Long) As String
    If CurrentLoads = "" Then
        AppendLoad = CStr(LoadNum)
    Else
        AppendLoad = CurrentLoads & CStr(LoadNum)
    End If
End Function

' Helper to create load combination
Sub CreateLoadCombination(Func As Object, CombNum As Long, CombTitle As String, CombString As String)
    On Error GoTo ErrorHandler
    Debug.Print "Creating: "; CombTitle; " (Comb "; CombNum; ")"
    Debug.Print "  String: "; CombString
    
    ' Delete if exists
    On Error Resume Next
    Func.Load.DeleteLoadCombination CombNum
    On Error GoTo ErrorHandler
    
    ' Create new
    Dim Result As Variant
    Result = Func.Load.CreateNewLoadCombination(CombTitle, CombNum)
    If Result = -1 Then
        Debug.Print "  ERROR: Failed to create combination at number "; CombNum
        Exit Sub
    End If
    
    ' Parse combination string
    Dim LoadParts() As String
    LoadParts = Split(Trim(CombString), " ")
    If (UBound(LoadParts) + 1) Mod 2 <> 0 Then
        Debug.Print "  ERROR: Invalid combination string format"
        Exit Sub
    End If
    
    Dim SuccessCount As Integer
    SuccessCount = 0
    Dim i As Integer
    For i = 0 To UBound(LoadParts) - 1 Step 2
        If IsNumeric(LoadParts(i)) And IsNumeric(LoadParts(i + 1)) Then
            Dim Factor As Double
            Dim LCNum As Long
            Factor = CDbl(LoadParts(i))
            LCNum = CLng(LoadParts(i + 1))
            ' Verify load case exists
            On Error Resume Next
            Dim LoadTitle As String
            LoadTitle = Func.Load.GetLoadCaseTitle(LCNum)
            If Err.Number <> 0 Then
                Debug.Print "  WARNING: Load case "; LCNum; " not found"
                Err.Clear
                GoTo NextLoad
            End If
            On Error GoTo ErrorHandler
            Dim AddRes As Variant
            AddRes = Func.Load.AddLoadAndFactorToCombination(CombNum, LCNum, Factor)
            If AddRes = 0 Then
                SuccessCount = SuccessCount + 1
            Else
                Debug.Print "  WARNING: Failed to add LC"; LCNum; " (code: "; AddRes; ")"
            End If
        End If
NextLoad:
    Next
    If SuccessCount > 0 Then
        Debug.Print "  Success: Added "; SuccessCount; " load cases"
    Else
        Debug.Print "  ERROR: No load cases added"
        On Error Resume Next
        Func.Load.DeleteLoadCombination(CombNum)
        On Error GoTo ErrorHandler
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "  ERROR: "; Err.Description; " ("; Err.Number; ")"
End Sub