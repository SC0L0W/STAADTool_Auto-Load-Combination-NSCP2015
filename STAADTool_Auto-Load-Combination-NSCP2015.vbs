Option Explicit

Sub Main()
    'Clear information
    Debug.Clear
    
    'Define OpenStaad Reference with error handling
    Dim Func As Object
    On Error GoTo ConnectionError
    
    Set Func = GetObject(, "StaadPro.OpenSTAAD")
    
    'Test connection
    Dim TestVar As Long
    TestVar = Func.Geometry.GetMemberCount()
    
    On Error GoTo 0
    
    'Define general parameters
    Dim i As Integer
    Dim Result As Variant
    
    Debug.Print "Starting NSCP 2015 Load Combination Generator..."
    Debug.Print "Multiple Combination Varieties"
    
    'Get Load Cases
    Dim TotalLoadCase As Long
    TotalLoadCase = Func.Load.GetPrimaryLoadCaseCount()
    Debug.Print "Total Load Cases: "; TotalLoadCase
    
    If TotalLoadCase = 0 Then
        MsgBox "No load cases found! Please define load cases first.", vbExclamation, "No Load Cases"
        Exit Sub
    End If
    
    'Get load case numbers and titles
    Dim LoadCaseNum() As Long
    ReDim LoadCaseNum(TotalLoadCase - 1)
    Func.Load.GetPrimaryLoadCaseNumbers(LoadCaseNum)
    
    Dim LoadCaseTitle() As String
    ReDim LoadCaseTitle(TotalLoadCase - 1)
    For i = 0 To TotalLoadCase - 1
        LoadCaseTitle(i) = Func.Load.GetLoadCaseTitle(LoadCaseNum(i))
        Debug.Print "Load Case "; LoadCaseNum(i); ": "; LoadCaseTitle(i)
    Next
    
    
    ' Classify load cases based on title patterns
    Dim DeadLoads1 As Long, DeadLoads2 As Long
    Dim LiveLoads As Long, LL1Loads As Long, LL2Loads As Long
    Dim RoofLoads As Long
    Dim EqX As Long, EqZ As Long
    Dim WindX As Long, WindZ As Long

    ' Initialize to 0
    DeadLoads1 = 0
    DeadLoads2 = 0
    LiveLoads = 0
    LL1Loads = 0
    LL2Loads = 0
    RoofLoads = 0
    EqX = 0
    EqZ = 0
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
            If InStr(LoadTitle, "LLR") > 0 Or InStr(LoadTitle, "RL") > 0 Or InStr(LoadTitle, "RFL") > 0 Then
                RoofLoads = LoadCaseNum_i
            Else
                LiveLoads = LoadCaseNum_i
            End If
        ElseIf InStr(LoadTitle, "EX") > 0 Then
            EqX = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "EZ") > 0 Then
            EqZ = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "WX") > 0 Then
            WindX = LoadCaseNum_i
        ElseIf InStr(LoadTitle, "WZ") > 0 Then
            WindZ = LoadCaseNum_i
        End If
    Next
    
    ' User prompt for design method
    Dim MethodName As String
    Dim EvValue As Double
    Dim LRFD_Start101 As Long, LRFD_Start201 As Long, LRFD_Start301 As Long
    Dim ASD_Start400 As Long, ASD_Start500 As Long
    
    LRFD_Start101 = 101
    LRFD_Start201 = 201
    LRFD_Start301 = 301
    ASD_Start400 = 400
    ASD_Start500 = 500
    
    If MsgBox("Select Design Method:" & vbCrLf & vbCrLf & _
              "YES = LRFD (Strength Design)" & vbCrLf & _
              "         Generate 3 series (101, 201, 301)" & vbCrLf & vbCrLf & _
              "NO = ASD (Allowable Stress Design)" & vbCrLf & _
              "         Generate 400+ and 500+ series", _
              vbYesNo + vbQuestion, "NSCP 2015 Design Method") = vbYes Then
        ' LRFD path
        MethodName = "LRFD"
        Dim EvInput As String
        EvInput = InputBox("Enter Ev factor for Series 301 (LRFD with Ev):" & vbCrLf & vbCrLf & _
                           "Typical values:" & vbCrLf & _
                           "• 0.27 (when Ev = 0.5 * Ca * I * D, Ca=0.532, I=1.0)" & vbCrLf & _
                           "• 0.20 (for lower seismic zones)" & vbCrLf & vbCrLf & _
                           "Enter value:", "Ev Factor for Series 301", "0.27")
        If EvInput = "" Then
            MsgBox "Operation cancelled.", vbInformation, "Cancelled"
            Exit Sub
        End If
        If Not IsNumeric(EvInput) Then
            MsgBox "Invalid Ev value entered!", vbExclamation, "Invalid Input"
            Exit Sub
        End If
        EvValue = CDbl(EvInput)
        
        ' Generate series
        Call GenerateLRFD_Series101(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, LRFD_Start101)
        LRFD_Start101 = LRFD_Start101 + 23
        Call GenerateLRFD_Series201(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, LRFD_Start201)
        LRFD_Start201 = LRFD_Start201 + 10
        Call GenerateLRFD_Series301(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, LRFD_Start301, EvValue)
        LRFD_Start301 = LRFD_Start301 + 10
        MsgBox "LRFD combinations generated successfully!", vbInformation
    Else
        ' ASD path
        MethodName = "ASD"
        Call GenerateASD_Basic(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, ASD_Start400)
        ASD_Start400 = ASD_Start400 + 10
        Call GenerateASD_Alternate(Func, DeadLoads1, DeadLoads2, LiveLoads, LL1Loads, LL2Loads, RoofLoads, EqX, EqZ, WindX, WindZ, ASD_Start500)
        ASD_Start500 = ASD_Start500 + 6
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
    CombTitle = "LRFD-101: 1.4D"
    CombString = " 1.4 " & DeadLoads1 & " 1.4 " & DeadLoads2
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6L + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-102: 1.2D+1.6L+1.6LL1+1.6LL2+0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 1.6 " & LiveLoads & " 1.6 " & LL1Loads & " 1.6 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + LL + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-103: 1.2D+LL+LL1+LL2+1.6Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-104: 1.2D+LLR+LL1+LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR + 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-105: 1.2D+LLR+Wx"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR -0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-106: 1.2D+LLR-0.5Wx"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR +0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-107: 1.2D+LLR+0.5Wz"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6LLR -0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-108: 1.2D+LLR-0.5Wz"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L + Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-109: 1.2D+0.5L+Ex+0.5LL1+LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L - Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-110: 1.2D+0.5L-Ex+0.5LL1+LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L + Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-111: 1.2D+0.5L+Ez+0.5LL1+LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L - Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-112: 1.2D+0.5L-Ez+0.5LL1+LL2"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L + Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-113: 0.9D+0.5L+Ex+0.5LL1+LL2"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L - Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-114: 0.9D+0.5L-Ex+0.5LL1+LL2"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L + Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-115: 0.9D+0.5L+Ez+0.5LL1+LL2"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L - Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-116: 0.9D+0.9L-Ez+0.5LL1+LL2"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " -1.0 " & LiveLoads & " -1.0 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L + Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-117: 1.2D+0.5L+Wx+0.5LL1+LL2+0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L - Wx + LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-118: 1.2D+0.5L-Wx+LL1+LL2+0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L + Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-119: 0.9D+0.5L+Wx+0.5LL1+LL2+0.5Roof"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L - Wx + LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-120: 0.9D+0.5L-Wx+LL1+LL2+0.5Roof"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' 1.2D + 0.5L + Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-121: 1.2D+0.5L+Wz+0.5LL1+LL2+0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L - Wz + LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-122: 1.2D+0.5L-Wz+LL1+LL2+0.5Roof"
    CombString = " 1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L + Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-123: 0.9D+0.5L+Wz+0.5LL1+LL2+0.5Roof"
    CombString = " 0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L - Wz + LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-124: 0.9D+0.5L-Wz+LL1+LL2+0.5Roof"
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
    CombTitle = "LRFD-202: 1.2D+1.6L+1.6LL1+1.6LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 1.6 " & LiveLoads & " 1.6 " & LL1Loads & " 1.6 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + LL + LL1 + LL2 + 1.6Roof
    CombNum = CombStart
    CombTitle = "LRFD-203: 1.2D+LL+LL1+LL2+1.6Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + LL + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-204: 1.2D+LLR+LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-205: 1.2D+LLR+0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-206: 1.2D+LLR-0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-207: 1.2D+LLR+0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-208: 1.2D+LLR-0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L + Ex + 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-209: 1.2D+0.5L+Ex+0.3Ez+0.5LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " 0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L + Ex - 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-210: 1.2D+0.5L+Ex-0.3Ez+0.5LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " -0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L - Ex + 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-211: 1.2D+0.5L-Ex+0.3Ez+0.5LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " 0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L - Ex - 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-212: 1.2D+0.5L-Ex-0.3Ez+0.5LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " -0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L + Ez + 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-213: 1.2D+0.5L+Ez+0.3Ex+0.5LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " 0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L + Ez - 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-214: 1.2D+0.5L+Ez-0.3Ex+0.5LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " -0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L - Ez + 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-215: 1.2D+0.5L-Ez+0.3Ex+0.5LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " 0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + L - Ez - 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-216: 1.2D+0.5L-Ez-0.3Ex+0.5LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " -0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L + Ex + 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-217: 0.9D+0.5L+Ex+0.3Ez+0.5LL1+LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " 0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L + Ex - 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-218: 0.9D+0.5L+Ex-0.3Ez+0.5LL1+LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqX & " -0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L - Ex + 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-219: 0.9D+0.5L-Ex+0.3Ez+0.5LL1+LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " 0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L - Ex - 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-220: 0.9D+0.5L-Ex-0.3Ez+0.5LL1+LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqX & " -0.3 " & EqZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L + Ez + 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-221: 0.9D+0.5L+Ez+0.3Ex+0.5LL1+LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " 0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L + Ez - 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-222: 0.9D+0.5L+Ez-0.3Ex+0.5LL1+LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & EqZ & " -0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L - Ez + 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-223: 0.9D+0.5L-Ez+0.3Ex+0.5LL1+LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " 0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + L - Ez - 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-224: 0.9D+0.5L-Ez-0.3Ex+0.5LL1+LL2"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & EqZ & " -0.3 " & EqX & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L + Wx + 0.3Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-225: 1.2D+0.5L+Wx+0.3Wz+0.5LL1+LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " 0.3 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L + Wx - 0.3Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-226: 1.2D+0.5L+Wx-0.3Wz+0.5LL1+LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " -0.3 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L - Wx + 0.3Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-227: 1.2D+0.5L-Wx+0.3Wz+0.5LL1+LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " 0.3 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L - Wx - 0.3Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-228: 1.2D+0.5L-Wx-0.3Wz+0.5LL1+LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " -0.3 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L + Wx + 0.3Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-229: 0.9D+0.5L+Wx+0.3Wz+0.5LL1+LL2+0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " 0.3 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L + Wx - 0.3Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-230: 0.9D+0.5L+Wx-0.3Wz+0.5LL1+LL2+0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindX & " -0.3 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L - Wx + 0.3Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-231: 0.9D+0.5L-Wx+0.3Wz+0.5LL1+LL2+0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " 0.3 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L - Wx - 0.3Wz + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-231: 0.9D+0.5L-Wx-0.3Wz+0.5LL1+LL2+0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindX & " -0.3 " & WindZ & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
   
    ' 1.2D + 0.5L + Wz + 0.3Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-232: 1.2D+0.5L+Wz+0.3Wx+0.5LL1+LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " 0.3 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L + Wz - 0.3Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-233: 1.2D+0.5L+Wz-0.3Wx+0.5LL1+LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " -0.3 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L - Wz + 0.3Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-234: 1.2D+0.5L-Wz+0.3Wx+0.5LL1+LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " 0.3 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 0.5L - Wz - 0.3Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-235: 1.2D+0.5L-Wz-0.3Wx+0.5LL1+LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " -0.3 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L + Wz + 0.3Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-236: 0.9D+0.5L+Wz+0.3Wx+0.5LL1+LL2+0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " 0.3 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L + Wz - 0.3Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-237: 0.9D+0.5L+Wz-0.3Wx+0.5LL1+LL2+0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " 1.0 " & WindZ & " -0.3 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L - Wz + 0.3Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-238: 0.9D+0.5L-Wz+0.3Wx+0.5LL1+LL2+0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " 0.3 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.9D + 0.5L - Wz - 0.3Wx + 0.5LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-239: 0.9D+0.5L-Wz-0.3Wx+0.5LL1+LL2+0.5Roof"
    CombString = "0.9 " & DeadLoads1 & " 0.9 " & DeadLoads2 & " 0.5 " & LiveLoads & " -1.0 " & WindZ & " -0.3 " & WindX & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 0.5 " & RoofLoads
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
    
    ' 1.2D + 1.6L + LL1 + LL2 + 0.5Roof
    CombNum = CombStart
    CombTitle = "LRFD-302: 1.2D+1.6L+1.6LL1+1.6LL2+0.5Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 1.6 " & LiveLoads & " 1.6 " & LL1Loads & " 1.6 " & LL2Loads & " 0.5 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + LL + LL1 + LL2 + 1.6Roof
    CombNum = CombStart
    CombTitle = "LRFD-303: 1.2D+LL+LL1+LL2+1.6Roof"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + LL + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-304: 1.2D+LLR+LL1+LL2"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & LiveLoads & " 0.5 " & LL1Loads & " 1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-305: 1.2D+LLR+0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wx
    CombNum = CombStart
    CombTitle = "LRFD-306: 1.2D+LLR-0.5Wx"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindX & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof + 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-307: 1.2D+LLR+0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " 0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 1.2D + 1.6Roof - 0.5Wz
    CombNum = CombStart
    CombTitle = "LRFD-308: 1.2D+LLR-0.5Wz"
    CombString = "1.2 " & DeadLoads1 & " 1.2 " & DeadLoads2 & " -0.5 " & WindZ & " 1.6 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' (1.2+Ev)D + L + Ex + 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-309: " & Format(D1, "0.00") & "D+0.5L+Ex+0.3Ez+0.5LL1+LL2"
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

    ' (1.2+Ev)D + L + Ex - 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-310: " & Format(D1, "0.00") & "D+0.5L+Ex-0.3Ez+0.5LL1+LL2"
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

    ' (1.2-Ev)D + L - Ex + 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-311: " & Format(D2, "0.00") & "D+0.5L-Ex+0.3Ez+0.5LL1+LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqX & " " & _
        "0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D + L - Ex - 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-312: " & Format(D2, "0.00") & "D+0.5L-Ex-0.3Ez+0.5LL1+LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqX & " " & _
        "-0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2+Ev)D + L + Ez + 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-313: " & Format(D1, "0.00") & "D+0.5L+Ez+0.3Ex+0.5LL1+LL2"
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

    ' (1.2+Ev)D + L + Ez - 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-314: " & Format(D1, "0.00") & "D+0.5L+Ez-0.3Ex+0.5LL1+LL2"
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

    ' (1.2-Ev)D + L - Ez + 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-315: " & Format(D2, "0.00") & "D+0.5L-Ez+0.3Ex+0.5LL1+LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqZ & " " & _
        "-0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (1.2-Ev)D + L - Ez - 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-316: " & Format(D2, "0.00") & "D+0.5L-Ez-0.3Ex+0.5LL1+LL2"
    CombString = _
        Format(D2, "0.00") & " " & DeadLoads1 & " " & _
        Format(D2, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqZ & " " & _
        "-0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + L + Ex + 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-317: " & Format(D3, "0.00") & "D+0.5L+Ex+0.3Ez+0.5LL1+LL2"
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

    ' (0.9+Ev)D + L + Ex - 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-318: " & Format(D3, "0.00") & "D+0.5L+Ex-0.3Ez+0.5LL1+LL2"
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

    ' (0.9-Ev)D + L - Ex + 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-319: " & Format(D4, "0.00") & "D+0.5L-Ex+0.3Ez+0.5LL1+LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqX & " " & _
        "0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D + L - Ex - 0.3Ez + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-320: " & Format(D4, "0.00") & "D+0.5L-Ex-0.3Ez+0.5LL1+LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqX & " " & _
        "-0.3 " & EqZ & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9+Ev)D + L + Ez + 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-321: " & Format(D3, "0.00") & "D+0.5L+Ez+0.3Ex+0.5LL1+LL2"
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

    ' (0.9+Ev)D + L + Ez - 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-322: " & Format(D3, "0.00") & "D+0.5L+Ez-0.3Ex+0.5LL1+LL2"
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

    ' (0.9-Ev)D + L - Ez + 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-323: " & Format(D4, "0.00") & "D+0.5L-Ez+0.3Ex+0.5LL1+LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqZ & " " & _
        "0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    ' (0.9-Ev)D + L - Ez - 0.3Ex + LL1 + LL2
    CombNum = CombStart
    CombTitle = "LRFD-324: " & Format(D4, "0.00") & "D+0.5L-Ez-0.3Ex+0.5LL1+LL2"
    CombString = _
        Format(D4, "0.00") & " " & DeadLoads1 & " " & _
        Format(D4, "0.00") & " " & DeadLoads2 & " " & _
        "0.5 " & LiveLoads & " " & _
        "1.0 " & EqZ & " " & _
        "-0.3 " & EqX & " " & _
        "0.5 " & LL1Loads & " " & _
        "1 " & LL2Loads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1

    Debug.Print "Series 301 Complete: " & (CombStart - 1) & " combinations"
End Sub

'=== ASD Basic: NSCP 203.4.1 ===
Sub GenerateASD_Basic(Func As Object, DeadLoads1 As Long, DeadLoads2 As Long, LiveLoads As Long, LL1Loads As Long, LL2Loads As Long, RoofLoads As Long, EqX As Long, EqZ As Long, WindX As Long, WindZ As Long, ByRef CombStart As Long)
    Debug.Print vbCrLf; "--- ASD Basic (NSCP 203.4.1) ---"
    Dim CombNum As Long
    Dim CombTitle As String
    Dim CombString As String
    
    ' D
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D"
    CombString = "1.0 " & DeadLoads1
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + L
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D+L"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & LiveLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + 0.75L + 0.525E (with orthogonal)
    ' D + 0.75L + 0.525Ex + 0.1575Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D+0.75L+0.525Ex+0.158Ez"
    CombString = "1.0 " & DeadLoads1 & " 0.75 " & LiveLoads & " 0.525 " & EqX & " 0.1575 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + 0.75L + 0.525Ex - 0.1575Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D+0.75L+0.525Ex-0.158Ez"
    CombString = "1.0 " & DeadLoads1 & " 0.75 " & LiveLoads & " 0.525 " & EqX & " -0.1575 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + 0.75L + 0.525Ez + 0.1575Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D+0.75L+0.525Ez+0.158Ex"
    CombString = "1.0 " & DeadLoads1 & " 0.75 " & LiveLoads & " 0.525 " & EqZ & " 0.1575 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + 0.75L + 0.525Ez - 0.1575Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D+0.75L+0.525Ez-0.158Ex"
    CombString = "1.0 " & DeadLoads1 & " 0.75 " & LiveLoads & " 0.525 " & EqZ & " -0.1575 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D + 0.7(Ex ± 0.3Ez)
    ' 0.6D + 0.7Ex + 0.21Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D+0.7Ex+0.21Ez"
    CombString = "0.6 " & DeadLoads1 & " 0.7 " & EqX & " 0.21 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D - 0.7Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D-0.7Ex"
    CombString = "0.6 " & DeadLoads1 & " -0.7 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D + 0.7Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D+0.7Ez"
    CombString = "0.6 " & DeadLoads1 & " 0.7 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D - 0.7Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D-0.7Ez"
    CombString = "0.6 " & DeadLoads1 & " -0.7 " & EqZ
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
    
    ' D
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D"
    CombString = "1.0 " & DeadLoads1
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + L + Lr (if exists)
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D+L+Lr"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & LiveLoads & " 1.0 " & RoofLoads
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + L + 0.7Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D+L+0.7Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & LiveLoads & " 0.7 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' D + L - 0.7Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": D+L-0.7Ex"
    CombString = "1.0 " & DeadLoads1 & " 1.0 " & LiveLoads & " -0.7 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D + 0.7Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D+0.7Ex"
    CombString = "0.6 " & DeadLoads1 & " 0.7 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D - 0.7Ex
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D-0.7Ex"
    CombString = "0.6 " & DeadLoads1 & " -0.7 " & EqX
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D + 0.7Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D+0.7Ez"
    CombString = "0.6 " & DeadLoads1 & " 0.7 " & EqZ
    Call CreateLoadCombination(Func, CombNum, CombTitle, CombString)
    CombStart = CombStart + 1
    
    ' 0.6D - 0.7Ez
    CombNum = CombStart
    CombTitle = "ASD-" & CombNum & ": 0.6D-0.7Ez"
    CombString = "0.6 " & DeadLoads1 & " -0.7 " & EqZ
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