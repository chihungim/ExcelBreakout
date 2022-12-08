VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim isGameRunning As Boolean

Private Sub CommandButton1_Click()
    Call StartGameLoop
End Sub

Sub UpdateScoreBoard()
    For Each shp In Shapes
        If shp.Name Like "_*" Then shp.Delete
    Next
    
    Dim d100: Dim d10: Dim d1
    
    d100 = Int([A6] / 100)
    d10 = [A6] / 10 Mod 10
    d1 = Int([A6] Mod 10)
    Application.CutCopyMode = True
    
    Sheet3.Shapes("_" & d100).Copy
    Sheet4.Range("L2").Select
    Paste

    Sheet3.Shapes("_" & d10).Copy
    Sheet4.Range("N2").Select
    Paste
    
    Sheet3.Shapes("_" & d1).Copy
    Sheet4.Range("P2").Select
    Paste
    Application.CutCopyMode = False
    
    [L2].Select
    CommandButton1.Activate
End Sub

Sub StartGameLoop()
    isGameRunning = True
    Call Worksheet_Activate
    t = Timer
    
    Do While isGameRunning
        If Timer - t >= 0.0001 Then
            Call ProcessInput
            Call MoveBall
            DoEvents
        End If
    Loop
End Sub

Sub MoveBall()
    If [bally] <= -50 Or [bally] >= 50 Then
        [balldiry] = [balldiry] * -1
    End If
    
    If [ballx] <= -50 Or [ballx] >= 50 Then
        [balldirx] = [balldirx] * -1
    End If
    
    
    l = [k2]: t = [L3]: r = [K4]: b = [L2]
    
    If [ballx] >= l And [ballx] <= r And [bally] <= t And [bally] >= b Then [balldiry] = [balldiry] * -1
    
    j = 2
    k = 0
    For i = 1 To [P1]
        If Range("S" & j) Then
            [balldiry] = [balldiry] * -1
            Range("Q" & j & ":R" & j + 4) = ""
            [A6] = [A6] + 1 * (6 - Int(i / 10))
            Debug.Print i
            Call UpdateScoreBoard
            Exit For
        End If
        j = j + 6
    Next
    
    If [bally] <= [L2] Then
        isGameRunning = False
        Call ResetGame
    End If
    
    'If [ballx] <= [K4] And [bally] >= [L5] And [bally] <= [L4] Then [balldiry] = [balldiry] * -1
    
    [bally] = [bally] + [balldiry]
    [ballx] = [ballx] + [balldirx]
End Sub

Sub ProcessInput()
    If GetAsyncKeyState(vbKeyLeft) Then
        [player1x] = WorksheetFunction.Max([player1x] - 2, -50)
    End If
    
    If GetAsyncKeyState(vbKeyRight) Then
        [player1x] = WorksheetFunction.Min([player1x] + 2, 30)
    End If
End Sub

Sub ResetGame()
    [A6] = 0
    isGameRunning = False
    Call SpawnBreak
    [ballx] = 0: [bally] = 0
    [player1x] = -10: [player1y] = -40
    Call UpdateScoreBoard
End Sub

Sub SpawnBreak()
    k = 2
    y = 45
    
    Range("Q2:S2000").ClearContents
    Dim ch As Chart
    Dim chobj As ChartObject
    Dim sc As seriesCollection
    
    
    Set chobj = ChartObjects(1)
    Set ch = chobj.Chart
    Set sc = ch.seriesCollection
    
'    For Each Item In sc
'        Item.Delete
'    Next
    
'    With sc.NewSeries
'        .Name = "Ball"
'        .XValues = Range("N2")
'        .Values = Range("O2")
'        .InvertColor = RGB(255, 0, 0)
'    End With
'
    
'    With sc.NewSeries
'        .Name = "Paddle"
'        .XValues = Range("K2:K6")
'        .Values = Range("L2:L6")
'    End With
    
    For i = 1 To 6
        x = -50
        For j = 1 To 10
            Range("Q" & k - 1 & ":R" & k - 1) = Array("X values", "Y Values")
            Range("Q" & k & ":S" & k + 4).Formula2R1C1 = Range("U2:W6").Formula2R1C1
            Range("Q" & k & ":R" & k) = Array(x, y)
'            With sc.NewSeries
'                .Name = "Block" & Int(k / 6)
'                .XValues = Range("Q" & k & ":Q" & k + 4)
'                .Values = Range("R" & k & ":R" & k + 4)
'            End With
            x = x + 10
            k = k + 6
        Next
        y = y - 5
    Next
End Sub

Public Sub CommandButton2_Click()
    Call ResetGame
End Sub

Private Sub Worksheet_Activate()
    Application.OnKey "{PGDN}", ""
    Application.OnKey "{PGUP}", ""
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
End Sub

Private Sub Worksheet_Deactivate()
    Application.OnKey "{PGDN}"
    Application.OnKey "{PGUP}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
End Sub

