Attribute VB_Name = "Ant"
'changed values increase/decrease by the mutation rate
Public Const MutationRate = 2

'the amount of food that appears per tick
Public Const FoodRate = 50

'the amount of ticks a Ant can survive (max)
Public Const KillAtTick = 800

Public Const MaxLifeSpan = 5000

'area to search for food
Public Const InitialSpeed = 1

'maximum area to search for food
Public Const MaxSpeed = 4

'the amount of food needed before reproduction
'(recommended to be 1/3 KillAtTick)
Public Const PropegateLevel = 100

'this prevents clutter and helps keep the program speed at a
'reasonable level
Public Const MinPropLevel = 5

Const MaxAnts = 200

Public Const vbWhite = &H0&
Public Const vbBlack = &HFFFFFF
Public Const vbGreen = &HC000&
Public Const vbCyan = &HFF00FF
Public Const vbLightRed = &H8080FF
Public Const vbBlue = &HFF0000
Public Const vbLightBlue = &HFFFF00
'Public Const vbRed = &HFF&

Public Enum Direction
    Up = 0
    Right = 1
    Down = 2
    Left = 3
End Enum

Public Enum WorldDetails
    HereEmpty = 0
    Food = 1
    Ant = 2
    Eater = 3
End Enum

'This defines the various values needed for each Ant
'Ants propigate by taking in a certain amount of food before
'splitting in half. Each half "evolves" into a new Ant, whos
'attributes are mutated by the Mutation Rate (constant).
'If the number of ticks reaches the KillLevel (can be mutated), then
'that Ant dies.
Public Type CreatureType
    XPos As Integer
    YPos As Integer
    Speed As Integer
    TickNum As Integer
    KillLevel As Integer
    FoodLevel As Integer
    FoodToSplit As Integer
    Direction As Direction
End Type

'stores the grid containing where the food is
'Please note that GridSize should be at least (MaxSpeed*2)+1
Public Const GridSize = 200
Public Grid(GridSize, GridSize) As WorldDetails

'store the data on individual Ants
Public Ants() As CreatureType
Public NumOfAnts As Integer

Public FoodAmount As Long

Dim Generations As Long

Public Fastest As Long
Public HighProp As Long
Public LifeSpan As Long

Public TotalDead As Long
Public TotalLife As Double
Public TotalSpeed As Double
Public TotalProp As Double

Public PauseResult As Long 'the number of ticks to stop updating the display for

Public Sub StartingAnt()
'This is only activated during load. It sets the default values of
'the StartingAnt Ant

NumOfAnts = 1
ReDim Preserve Ants(NumOfAnts)

Randomize

With Ants(NumOfAnts)
    .Direction = GetRndInt(Up, Left) '((Left - Up + 1) * Rnd + Up)
    .FoodLevel = 0
    .FoodToSplit = GetRndInt(MinPropLevel, 100) 'Int((100 - MinPropLevel + 1) * Rnd + MinPropLevel)
    .KillLevel = GetRndInt(50, MaxLifeSpan) 'Int((MaxLifeSpan - 50 + 1) * Rnd + 50)
    .Speed = InitialSpeed
    .TickNum = 0
    Do
        .XPos = GetRndInt(0, GridSize) '(Int((GridSize + 1) * Rnd) Mod (GridSize + 1))
        .YPos = GetRndInt(0, GridSize) '(Int((GridSize + 1) * Rnd) Mod (GridSize + 1))
    Loop Until Grid(.XPos, .YPos) = HereEmpty
    
    Fastest = .Speed
    LifeSpan = .KillLevel
    HighProp = .FoodToSplit
    TotalLife = LifeSpan
    TotalProp = HighProp
    TotalSpeed = Fastest
    
    frmLife.lblFast.Caption = "Fastest : " & Fastest
    frmLife.lblLLife.Caption = "Longest Life : " & LifeSpan
    frmLife.lblHProp.Caption = "Lowest Prop. Rate : " & HighProp
End With

Generations = 1

'cover 1/3 the total area with food
Call frmLife.CreateFood(((GridSize ^ 2) * 0.33) / FoodRate, True)
End Sub

Public Sub KillAnt(AntNum As Integer)
'this reduces the number of Ants by one.

Dim Counter As Integer

'a dead Ant becomes food /(empty)
Grid(Ants(AntNum).XPos, Ants(AntNum).YPos) = Food
Call frmLife.DrawDot(Ants(AntNum).XPos, Ants(AntNum).YPos, vbWhite)
Call frmLife.DrawDot(Ants(AntNum).XPos, Ants(AntNum).YPos, vbGreen)

're-calculate averages
TotalLife = TotalLife - Ants(AntNum).KillLevel
TotalSpeed = TotalSpeed - Ants(AntNum).Speed
TotalProp = TotalProp - Ants(AntNum).FoodToSplit

'kill Ant by moving the array elements down one
For Counter = AntNum To (NumOfAnts - 1)
    Ants(Counter) = Ants(Counter + 1)
Next Counter

'reduce the Ant number if able
If NumOfAnts >= 1 Then
    NumOfAnts = NumOfAnts - 1
Else
    'if no ants to kill, kill an ant eater to control population and
    'give the ants a change to thrive again
    Call KillAnt(1)
End If
ReDim Preserve Ants(NumOfAnts)

'show stats
TotalDead = TotalDead + 1
frmLife.lblDeath.Caption = "Death Rate : " & Format((TotalDead / Generations) * 100, "0.0") & "%"
End Sub

Private Sub MutateAnt(Num As Integer)
'This sub will MutateAnt all the values of a Ant withing the
'mutation rate.

Dim NewVal As Integer
Dim Upperbound As Integer
Dim Lowerbound As Integer
Dim X As Integer
Dim Y As Integer
Dim Being As CreatureType

If NumOfAnts >= MaxAnts Then
    'if population pressure is too high, increase life span
    Ants(Num).FoodLevel = 0
    Ants(Num).KillLevel = Ants(Num).KillLevel + FoodRate
    TotalLife = TotalLife + FoodRate
    Exit Sub
End If

Being = Ants(Num)

Upperbound = MutationRate
Lowerbound = -MutationRate

'evolve Ant with new settings and "age" the original ant
With Being
    .FoodLevel = 0
    .TickNum = 0
    .Direction = GetRndInt(Left, Up) '((Left - Up + 1) * Rnd + Up)
    Upperbound = Upperbound * (GridSize / 15)
    Lowerbound = Lowerbound * (GridSize / 15)
    .FoodToSplit = .FoodToSplit + GetRndInt(Lowerbound, Upperbound) '((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    If (.FoodToSplit < MinPropLevel) Then
        .FoodToSplit = MinPropLevel
    End If
    If (.FoodToSplit < HighProp) Then
        'new Lowest propegation level
        HighProp = .FoodToSplit
        frmLife.lblHProp.Caption = "Lowest Prop. Level : " & HighProp
    End If
    TotalProp = TotalProp + .FoodToSplit
    
    Upperbound = MutationRate * (GridSize / 4)
    Lowerbound = MutationRate * (GridSize / 4)
    .KillLevel = .KillLevel + GetRndInt(Lowerbound, Upperbound) '((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    If .KillLevel > MaxLifeSpan Then
        .KillLevel = MaxLifeSpan
    End If
    If .KillLevel > LifeSpan Then
        'new Lowest life span
        LifeSpan = .KillLevel
        frmLife.lblLLife.Caption = "Longest Life : " & LifeSpan
    End If
    TotalLife = TotalLife + .KillLevel
    
    Upperbound = MutationRate
    Lowerbound = -MutationRate
    .Speed = .Speed + GetRndInt(Lowerbound, Upperbound) '((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    If .Speed > MaxSpeed Then
        .Speed = MaxSpeed
    End If
    If (.Speed > Fastest) Then
        'new fastest Ant
        Fastest = .Speed
        frmLife.lblFast.Caption = "Fastest : " & Fastest
    End If
    
    'set averages
    TotalSpeed = TotalSpeed + .Speed
End With

'the old Ant produces slower and searches smaller, dies faster
Ants(Num).FoodLevel = 0
Ants(Num).FoodToSplit = Ants(Num).FoodToSplit + 1
TotalProp = TotalProp + 1
'If (Ants(Num).FoodToSplit > MinPropLevel) Then
'    Ants(Num).FoodToSplit = MinPropLevel
'End If
Ants(Num).Speed = Ants(Num).Speed - 1
TotalSpeed = TotalSpeed - 1
'If Ants(Num).Speed > MaxSpeed Then
'    Ants(Num).Speed = MaxSpeed
'End If
Ants(Num).KillLevel = Ants(Num).KillLevel - 1
TotalLife = TotalLife - 1
Ants(Num).Direction = ((Left - Up + 1) * Rnd + Up)

'New Ant details
X = Being.XPos
Y = Being.YPos
If Not GetEmptyPos(X, Y) Then
    'if there is no available space for a new Ant, then don't
    'create one
    Exit Sub
End If

'create a new Ant
NumOfAnts = NumOfAnts + 1
ReDim Preserve Ants(NumOfAnts)
Ants(NumOfAnts) = Being
Ants(NumOfAnts).XPos = X
Ants(NumOfAnts).YPos = Y
Grid(X, Y) = Ant
Call frmLife.DrawDot(X, Y, vbBlack)

'show stats
Generations = Generations + 1
frmLife.lblGen.Caption = "Generations : " & Generations
frmLife.lblProp.Caption = "Propegation Rate : " & Format((NumOfAnts / Generations) * 100, "0.0") & "%"
End Sub

Public Function GetEmptyPos(X As Integer, Y As Integer) As Boolean
'this function returnes the position of an empty space around a given
'set of co-ordinates if a space is available (True), else the function
'will return the same position and the value (False).

Dim XVal As Integer
Dim YVal As Integer
Dim NewX As Integer
Dim NewY As Integer
Dim Counter As Integer

XVal = X
YVal = Y

'check above, below and to the sides of the given postion to check for
'an empty square.

GetEmptyPos = False

'check to each side
If X > 0 Then
    'check to the left
    If Grid((X - 1), Y) = HereEmpty Then
        'return new position
        X = X - 1
        GetEmptyPos = True
        Exit Function
    End If
End If

If X < GridSize Then
    'check to the right
    If Grid((X + 1), Y) = HereEmpty Then
        'return new position
        X = X + 1
        GetEmptyPos = True
        Exit Function
    End If
End If

'check above and below the current position
If Y > 0 Then
    'check above
    If Grid(X, (Y - 1)) = HereEmpty Then
         'return new position
        Y = Y - 1
        GetEmptyPos = True
        Exit Function
   End If
End If

If Y < GridSize Then
    'check below
    If Grid(X, (Y + 1)) = HereEmpty Then
        'return new position
        Y = Y + 1
        GetEmptyPos = True
        Exit Function
    End If
End If
End Function

Public Sub MoveAnt(AntNum As Integer)
'move the Ant in the direction it is meant to (if it can) and
'change the direction.
'If the ant cannot find food, and it is past it possilbe cannible rate,
'then turn the ant cannible

Const X = 0
Const Y = 1

Dim Target(2) As Integer
Dim Got(2) As Integer
Dim Speed As Integer
Dim MyX As Integer
Dim MyY As Integer

'check to see if changes are needed to be made to the Ant
'and the Ant has been removed, then exit
If Not CheckAnt(AntNum) Then
    Exit Sub
End If

'randomize

Got(X) = Ants(AntNum).XPos
Got(Y) = Ants(AntNum).YPos
Speed = Ants(AntNum).Speed

MyX = Got(X)
MyY = Got(Y)

If Not frmLife.SearchForFood(MyX, MyY, Speed, Ants(AntNum).Direction) Then
    Target(X) = Got(X)
    Target(Y) = Got(Y)
    
    'project co-ordinates to new target (move the target in the given
    'direction X amount of blocks - X = Speed)
    Select Case Ants(AntNum).Direction
    Case Up
        Target(Y) = Got(Y) - Speed
        If Target(Y) < 0 Then
            Target(Y) = (GridSize + 1) + Target(Y)
        End If
    Case Right
        Target(X) = (Got(X) + Speed) Mod (GridSize + 1)
    Case Down
        Target(Y) = (Got(Y) + Speed) Mod (GridSize + 1)
    Case Left
        Target(X) = Got(X) - Speed
        If Target(X) < 0 Then
            Target(X) = (GridSize + 1) + Target(X)
        End If
    End Select
    
    'if it can't find food, double it's kill time this turn
    Ants(AntNum).TickNum = Ants(AntNum).TickNum + FoodRate
    
    'temperorly increase it's speed to twice what it was
    TotalSpeed = TotalSpeed - Ants(AntNum).Speed
    Ants(AntNum).Speed = Ants(AntNum).Speed + 2 '* 2
    If Ants(AntNum).Speed > (MaxSpeed * 2) Then
        Ants(AntNum).Speed = (MaxSpeed * 2)
    End If
    TotalSpeed = TotalSpeed + Ants(AntNum).Speed
    
    'change the direction if it can't find food in the current direction
    Ants(AntNum).Direction = GetRndInt(Up, Left) '((Left - Up + 1) * Rnd + Up)
Else
    'found food
    Target(X) = MyX
    Target(Y) = MyY
    
    'increase ant ant's speed
    TotalSpeed = TotalSpeed - Ants(AntNum).Speed
    Ants(AntNum).Speed = Ants(AntNum).Speed + 1
    If Ants(AntNum).Speed > MaxSpeed Then
        Ants(AntNum).Speed = MaxSpeed
    End If
    TotalSpeed = TotalSpeed + Ants(AntNum).Speed
End If

If Grid(Target(X), Target(Y)) <> Ant Then
    'move Ant to new position
    Grid(Ants(AntNum).XPos, Ants(AntNum).YPos) = HereEmpty
    
    'draw a blank dot on the old position
    Call frmLife.DrawDot(Got(X), Got(Y), vbWhite)
    
    'if the new position contains food, then eat it
    If Grid(Target(X), Target(Y)) = Food Then
        Ants(AntNum).FoodLevel = Ants(AntNum).FoodLevel + 1
        FoodAmount = FoodAmount - 1
    End If
    
    Ants(AntNum).XPos = Target(X)
    Ants(AntNum).YPos = Target(Y)
    Grid(Target(X), Target(Y)) = Ant
    Call frmLife.DrawDot(Target(X), Target(Y), vbBlack)
End If

Ants(AntNum).TickNum = Ants(AntNum).TickNum + 1
End Sub

Public Function CheckAnt(Num As Integer) As Boolean
'This will check the stats of the selected Ant and make the
'appropiate changes.
Static test As Integer

CheckAnt = True

If (Ants(Num).FoodLevel >= Ants(Num).FoodToSplit) Then 'And (Ants(Num).FoodToSplit >= MinPropLevel)
    'MutateAnt Ant
    Call MutateAnt(Num)
End If

If (Ants(Num).TickNum >= Ants(Num).KillLevel) Or (Ants(Num).Speed = 0) Then
    'delete the Ant and report that the Ant has been removed
    Call KillAnt(Num)
    CheckAnt = False
End If
End Function

Public Function CheckRange(ByVal Min As Integer, ByVal Max As Integer, ByVal Value As Integer) As Integer
'This function will check to see if Value is between Min and Max.
'If not then it will divide Value until it fits between Min and Max.

Dim Offset As Integer
Dim TempVal As Integer

'exit program is max is less than or equal to min
If Max <= Min Then
    CheckRange = Min
    Exit Function
End If

'make sure Value does not exceed bounds under any initial conditions
Offset = Max - Min
Value = (Value Mod (Offset + 1)) + Min
If Value < Min Then
    Value = Max - (Min - Value)
End If

CheckRange = Value
End Function

Public Function GetRndInt(ByVal Min As Integer, ByVal Max As Integer) As Integer
'This function will produce a random number between the specified
'values

Dim Temp As Integer

'if the two values are equal then
If Min = Max Then
    'return same number and exit
    GetRndInt = Min
    Exit Function
End If

'if Min is bigger than Max then swap values
If Min > Max Then
    Temp = Min
    Min = Max
    Max = Min
End If

'Randomize
GetRndInt = CheckRange(Min, Max, Int((Max - Min + 1) * Rnd + Min))
End Function

