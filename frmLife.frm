VERSION 5.00
Begin VB.Form frmLife 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artificial Life"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   DrawWidth       =   2
   FillColor       =   &H008080FF&
   ForeColor       =   &H00000000&
   Icon            =   "frmLife.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleMode       =   0  'User
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEaterNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Ant Eater Number : 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label lblAProp 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Pop. Level :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label lblALife 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Average Life Span :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label lblASpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Speed :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Line lnBreak2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3733.333
      X2              =   3733.333
      Y1              =   6120
      Y2              =   7200
   End
   Begin VB.Label lblHProp 
      BackStyle       =   0  'Transparent
      Caption         =   "Lowest Prop. Level :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblLLife 
      BackStyle       =   0  'Transparent
      Caption         =   "Longest Life :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblFast 
      BackStyle       =   0  'Transparent
      Caption         =   "Fastest :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Line lnBreak1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1813.333
      X2              =   1813.333
      Y1              =   6120
      Y2              =   7200
   End
   Begin VB.Label lblTick 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticks : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label lblProp 
      BackStyle       =   0  'Transparent
      Caption         =   "Propegation Rate :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblDeath 
      BackStyle       =   0  'Transparent
      Caption         =   "Death Rate : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblGen 
      BackStyle       =   0  'Transparent
      Caption         =   "Generations : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "Number :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
End
Attribute VB_Name = "frmLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program was created simply because I wanted to see what
'happened when random choices were taken within ceratin rules.
'The averages are taken because I wanted to see how they changed
'overall after a certain period of time. No two runnings of this
'program are exactly the same. Each time you run this program, it will
'be slightly different.
'On a technical note, this program contains a nice 2D array searching
'technique, even if i do say so myself :)
'
'Eric
'DiskJunky@hotmail.com

Public Sub DrawBlank()
'this draws a blank square

Dim X As Integer
Dim Y As Integer

For X = 0 To GridSize - 1
    For Y = 0 To GridSize - 1
        Grid(X, Y) = HereEmpty
    Next Y
Next X
End Sub

Public Sub CreateFood(Optional Multiply As Long, Optional Start As Boolean = False)
'This randomly places food in different places in the grid

Dim Upperbound As Integer
Const Lowerbound = 0

Dim Counter As Integer
Dim X As Integer
Dim Y As Integer
Dim MyFoodRate As Long
Dim MyUpperbound As Integer
Dim MyLowerbound As Integer

Upperbound = GridSize - 1

If Multiply <> 0 Then
    MyFoodRate = FoodRate * Multiply
Else
    MyUpperbound = FoodRate + 3
    MyLowerbound = 0
    MyFoodRate = GetRndInt(MyLowerbound, MyUpperbound) '((MyUpperbound - MyLowerbound + 1) * Rnd + MyLowerbound)
End If

'if the amount of food in grid is greater than 3/4 then, don't create more
If (FoodAmount + MyFoodRate) > ((GridSize ^ 2) - ((GridSize ^ 2) / 4)) Then
    Exit Sub
Else
    FoodAmount = FoodAmount + MyFoodRate
End If

'create "FoodRate" bits of food
For Counter = 1 To MyFoodRate
    Do
        'randomize
        X = GetRndInt(Lowerbound, Upperbound) '((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
        Y = GetRndInt(Lowerbound, Upperbound) '((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    Loop Until (Grid(X, Y) <> WorldDetails.Ant) And (Grid(X, Y) <> WorldDetails.Eater)
    
    If Not Start Then
        'draw the dot if not creating food for the start.
        Call frmLife.DrawDot(X, Y, vbGreen)
    End If
    
    Grid(X, Y) = Food
Next Counter

If Start Then
    Call DrawFrame
End If
End Sub

Public Sub FlipVal(ByRef Val1 As Integer, ByRef Val2 As Integer)
'This function swaps the two value

Dim Temp As Integer

Temp = Val1
Val1 = Val2
Val2 = Temp
End Sub

Public Function RndBool() As Integer
'this function returns either 1 or -1

Dim RndVal As Integer

'randomize
'RndVal = Int(Rnd)

If Rnd < 0.5 Then
    RndVal = -1
Else
    RndVal = 1
End If

RndBool = RndVal
End Function

Public Function SearchForFood(XCo As Integer, YCo As Integer, Speed As Integer, MyDirection As Direction, Optional FoodType As Integer) As Boolean
'This will look for a piece of food within half the speed (distance)
'radius from the co-ordinates given starting from the point outwards,
'in a random direction.

Const X = 0
Const Y = 1

Dim CounterX As Integer
Dim CounterY As Integer
Dim Counter As Integer
Dim Low(2) As Integer
Dim High(2) As Integer
Dim Found(2) As Integer
Dim TheFood As WorldDetails
Dim Offset(2) As Integer
Dim Step(2) As Integer

If FoodType <> 0 Then
    'look for ants
    TheFood = WorldDetails.Ant
Else
    TheFood = WorldDetails.Food
End If

SearchForFood = False
If Speed < 1 Then
    Exit Function
End If

'set the range in which to look for food
High(X) = XCo + Speed
Low(X) = XCo - Speed
High(Y) = YCo + Speed
Low(Y) = YCo - Speed

'half the search range depending on the direction
'RndBool is used to start the search from a random direction in the
'given search area
MyDirection = MyDirection Mod 4
Select Case MyDirection
Case Up
    'don't search the lower half
    Low(Y) = YCo - 1
    
    Step(X) = RndBool
    Step(Y) = -1
Case Right
    'don't search the left half
    Low(X) = XCo - 1
    Step(X) = 1
    Step(Y) = RndBool
Case Down
    'don't search the top half
    High(Y) = YCo + 1
    Step(X) = -1
    Step(Y) = RndBool
Case Direction.Left
    'don't search the right half
    High(X) = XCo + 1
    Step(X) = RndBool
    Step(Y) = 1
End Select

'start searching the area directly ahead
Offset(X) = (High(Y) - Low(Y)) / 2
Offset(Y) = (High(X) - Low(X)) / 2

'randomize search direction
For Counter = 0 To 1
    If Step(Counter) = -1 Then
        Call FlipVal(Low(Counter), High(Counter))
    End If
Next Counter

For CounterX = Low(X) To High(X) Step Step(X)
    For CounterY = Low(Y) To High(Y) Step Step(Y)
        'check values to see if they go past the edge and change them
        'accordingly.
        Found(X) = CheckRange(0, GridSize, OffsetRange(High(X), Low(X), CounterX, Offset(X)))
        Found(Y) = CheckRange(0, GridSize, OffsetRange(High(Y), Low(Y), CounterY, Offset(Y)))
        
        'look for food (1 = AntFood, 2 = Anteaters food (ants))
        If Grid(Found(X), Found(Y)) = TheFood Then
            XCo = Found(X)
            YCo = Found(Y)
            SearchForFood = True
            Exit Function
        End If
    Next CounterY
Next CounterX
End Function

Public Function OffsetRange(Max As Integer, Min As Integer, Value As Integer, Offset As Integer) As Integer
'this function will add Offset to Value. Should the result exceed
'Max, then the result continues from Min.
'eg, assume the function was called with;
'OffsetRange(50,10,45,10)
'the Result = 15 because 45+10=55 but Max = 50. The excess 5 is
'added to Min (10) to give the result.

Dim Result As Integer

If Min > Max Then
    'swap values
    Call FlipVal(Max, Min)
End If

Result = (Value + Offset) Mod (Max + 1)
If Result < Value Then
    'add to min
    Result = Result + Min
End If

OffsetRange = Result
End Function

Public Sub DrawDot(ByVal X As Integer, ByVal Y As Integer, ByVal Colour As Long)
'this draws a dot at the specified coordinates in the given colour

Static LastWidth As Byte

'if the display is pause, then exit
If PauseResult > 0 Then
    Exit Sub
End If

Select Case Colour
Case vbGreen
    If LastWidth <> 1 Then
        frmLife.DrawWidth = 1
        LastWidth = 1
        frmLife.ForeColor = Colour
    End If
Case Else
    If LastWidth <> 2 Then
        frmLife.DrawWidth = 2
        LastWidth = 2
    End If
    frmLife.ForeColor = Colour
End Select

frmLife.PSet (X * 2 * Screen.TwipsPerPixelX, Y * 2 * Screen.TwipsPerPixelY)
DoEvents
End Sub

Private Sub Form_Activate()
Dim Counter As Integer

Static TickCount As Double
Static StartOnly As Boolean
Static LastEaterNum As Integer

'This is where all the work is done. The program really runs from here.
'I put this as a While loop instead of a timer because it's quicker :)
While True 'NumOfAnts > 0
    If Not StartOnly Then
        Call DrawBlank
        Call StartingAnt
        Call DrawDot(Ants(1).XPos, Ants(1).YPos, vbBlack)
        StartOnly = True
    End If
    
    Call frmLife.CreateFood
    
    'move each ant
    For Counter = 1 To NumOfAnts
        If Counter > NumOfAnts Then
            Exit For
        End If
        Call MoveAnt(Counter)
        DoEvents
    Next Counter
    
    'if there are no eaters and the number of ants is available, then
    'create the ant eaters
    If (NumOfEaters = 0) And (NumOfAnts > IntroduceEaterAt) Then
        Call StartingEater
    Else
        If ((NumOfAnts * 3) < NumOfEaters) And ((TickNum Mod 10) = 0) Then 'And (NumOfEaters > 1)
            'reduce the number of eaters if there are more eaters than
            'ants, every tehnth tick
            Call KillEater(1)
        End If
        
        'move the eaters
        For Counter = 1 To NumOfEaters
            If Counter > NumOfEaters Then
                Exit For
            End If
            Call MoveEater(Counter)
            DoEvents
        Next Counter
    End If
    
    'calculate averages
    If (NumOfAnts > 0) And (PauseResult = 0) Then
        lblASpeed.Caption = "Average Speed : " & Format((TotalSpeed / NumOfAnts), "0.0") '& "%"
        lblALife.Caption = "Average Life Span : " & Format((TotalLife / NumOfAnts), "0.0") '& "%"
        lblAProp.Caption = "Average Prop. Rate : " & Format((TotalProp / NumOfAnts), "0.0") '& "%"
    End If
    
    If NumOfAnts = 0 Then
        'start again
        StartOnly = False
    End If
    
    'update the eater number display
    If (LastEaterNum <> NumOfEaters) Then 'And (PauseResult = 0)
        'update the display
        LastEaterNum = NumOfEaters
        lblEaterNum.Caption = "Ant Eater Number : " & NumOfEaters
    End If
    
    'update display if able
    lblCount = "Population : " & NumOfAnts
    
    If PauseResult <> 0 Then
        'one less tick to wait
        PauseResult = PauseResult - 1
        If PauseResult = 0 Then
            'redraw the entire frame
            Call DrawFrame
        End If
    End If
    
    TickCount = TickCount + 1
    lblTick.Caption = "Ticks : " & Format(TickCount, "0")
Wend

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'if the key pressed was between 1 and 9, then pause exexution
'of the display for key * 100 ticks

If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    PauseResult = Val(Chr(KeyAscii)) * 100
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If NumOfAnts = 0 Then
        Call DrawBlank
        Call StartingAnt
        Call frmLife.DrawDot(Ants(1).XPos, Ants(1).YPos, vbBlack)
        StartOnly = True
    Else
        'create one third of food
        Call frmLife.CreateFood(((GridSize ^ 2) * 0.33) / FoodRate)
    End If
Else
    'vbRightButton - stop drawing results for a certain number of
    'ticks
    PauseResult = 100
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Public Sub DrawFrame()
'This will draw everything in the entire grid

Dim X As Integer
Dim Y As Integer
Dim Colour As Long

'draw a blank square
frmLife.Cls

For X = 1 To GridSize
    For Y = 1 To GridSize
        'draw the appropiate dot
        Select Case Grid(X, Y)
        Case WorldDetails.HereEmpty
            Colour = vbWhite
        Case WorldDetails.Food
            Colour = vbGreen
        Case WorldDetails.Ant
            Colour = vbBlack
        Case WorldDetails.Eater
            Colour = vbCyan
        End Select
        
        If Colour <> vbWhite Then
            Call DrawDot(X, Y, Colour)
        End If
    Next Y
Next X
End Sub

