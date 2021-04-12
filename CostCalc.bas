Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub Button2_Click()
    Dim i As Integer
    Dim k As Integer
    Dim Count As Integer
    Dim years As Integer
    Range("A9:A120").Clear
    Range("B9:B120").Clear
    Range("C9:C120").Clear
    Range("D9:D120").Clear
    If Range("B3").Value = "" Then MsgBox "Please Enter a value for number of periods"
    If Not Range("B3").Value = "" Then i = Range("B3").Value
    years = Int(i / 12)
    For Count = 1 To years
       For k = 1 To 12
           Cells(k + (12 * (Count - 1)) + 8, 1).Value = k + (12 * (Count - 1))
           Cells(k + (12 * (Count - 1)) + 8, 2).Value = Range("F4").Value * (1 + Range("B5").Value) ^ (Count - 1)
       Next k
    Next Count
    Range("C9") = Range("B4") * Range("F5") / 12
    Range("D9") = Range("B4") - Range("B9") + Range("C9")
    For Count = 1 To i - 1
        Cells(9 + Count, 3).Value = Cells(8 + Count, 4).Value * Range("F5").Value / 12
        Cells(9 + Count, 4).Value = Cells(8 + Count, 4).Value - Cells(9 + Count, 2).Value + Cells(9 + Count, 3).Value
    Next Count
    
    Dim TOL As Double
    Dim Increment As Double
    Dim Flag As Boolean
    Dim PreviousVal As Double
    TOL = 0.00001
    Increment = 0.01
    Flag = "False"
    PreviousVal = Cells(8 + i, 4)
    While Abs(Cells(8 + i, 4) - 0) > TOL And StopVal < 150
        Flag = "False"
        Range("F5") = Range("F5") + Increment
        If Range("F5") < 0 Then Range("F5") = Range("F5") * -1
        Range("C9") = Range("B4") * Range("B6") / 12
        Range("D9") = Range("B4") - Range("B9") + Range("C9")
        For Count = 1 To years
           For k = 1 To 12
               Cells(k + (12 * (Count - 1)) + 8, 1).Value = k + (12 * (Count - 1))
               Cells(k + (12 * (Count - 1)) + 8, 2).Value = Range("F4").Value * (1 + Range("B5").Value) ^ (Count - 1)
           Next k
        Next Count
        For Count = 1 To i - 1
            Cells(9 + Count, 3).Value = Cells(8 + Count, 4).Value * Range("B6").Value / 12
            Cells(9 + Count, 4).Value = Cells(8 + Count, 4).Value - Cells(9 + Count, 2).Value + Cells(9 + Count, 3).Value
        Next Count
        
        If Abs(PreviousVal) < Abs(Cells(8 + i, 4)) Then Increment = Increment * -1
        If PreviousVal > 0 And Cells(8 + i, 4) < 0 Then Flag = "True"
        If PreviousVal < 0 And Cells(8 + i, 4) > 0 Then Flag = "True"
        If Flag = "True" Then Increment = Increment / 2
        PreviousVal = Cells(8 + i, 4)
        StopVal = StopVal + 1
    Wend
    Range("B9:D120").Style = "Currency"
    
End Sub


