Attribute VB_Name = "AntEater"
'the number of ants to introduce the eaters at
Public Const IntroduceEaterAt = 20

'changed values increase/decrease by the mutation rate
Public Const MutationRateEater = 3

'the amount of food to look for per tick
Public Const FoodRateEater = 1

'the amount of ticks a Ant can survive (max)
Public Const KillAtTickEater = 2000

Public Const MaxLifeSpanEater = 5000

'area to search for food
Public Const InitialSpeedEater = 10

'maximum area to search for food
Public Const MaxSpeedEater = 15

'the amount of food needed before reproduction
'(recommended to be 1/3 KillAtTickEater)
Public Const PropegateLevelEater = 50

'this prevend clutter and helps keep the program speed at a
'reasonable level
Public Const MinPropLevelEater = 50

'the maximum amount of eaters on the screen at any one time.
Const MaxEaters = 100

'store data on the ant eaters
Public Eaters() As CreatureType
Public NumOfEaters As Integer

Public Sub StartingEater()
'This is only activated during load. It sets the default values of
'the Startingeater eater

NumOfEaters = 1
ReDim Preserve Eaters(NumOfEaters)

With Eaters(NumOfEaters)
    .Direction = GetRndInt(Left, Up)
    .FoodLevel = 0
    .FoodToSplit = GetRndInt(MinPropLevel, PropegateLevelEater)
    .KillLevel = GetRndInt(KillAtTickEater, MaxLifeSpanEater)
    .Speed = GetRndInt(MaxSpeedEater, InitialSpeedEater)
    .TickNum = 0
    Do
        'keep searching for a free space until one is found
        .XPos = GetRndInt(0, GridSize)
        .YPos = GetRndInt(0, GridSize)
    Loop Until Grid(.XPos, .YPos) = HereEmpty
End With
End Sub

Public Sub KillEater(EaterNum As Integer)
'this reduces the number of Eaters by one.

Dim Counter As Integer

'a dead Eater becomes food /(empty)
Grid(Eaters(EaterNum).XPos, Eaters(EaterNum).YPos) = Food
Call frmLife.DrawDot(Eaters(EaterNum).XPos, Eaters(EaterNum).YPos, vbWhite)
Call frmLife.DrawDot(Eaters(EaterNum).XPos, Eaters(EaterNum).YPos, vbGreen)

're-calculate averages
TotalLife = TotalLife - Eaters(EaterNum).KillLevel
TotalSpeed = TotalSpeed - Eaters(EaterNum).Speed
TotalProp = TotalProp - Eaters(EaterNum).FoodToSplit

'kill Eater
For Counter = EaterNum To (NumOfEaters - 1)
    Eaters(Counter) = Eaters(Counter + 1)
Next Counter

'reduce the Eater number
NumOfEaters = NumOfEaters - 1
ReDim Preserve Eaters(NumOfEaters)

'show stats
TotalDead = TotalDead + 1
'frmLife.lblDeath.Caption = "Death Rate : " & Format((TotalDead / Generations) * 100, "0.0") & "%"
End Sub

Public Sub MoveEater(EaterNum As Integer)
'move the Eater in the direction it is meEater to (if it can) and
'change the direction.

Const X = 0
Const Y = 1

Dim Target(2) As Integer
Dim Got(2) As Integer
Dim Speed As Integer
Dim MyX As Integer
Dim MyY As Integer
Dim GotAntNum As Integer
Dim Turn As Integer

'check to see if changes are needed to be made to the Eater
'and the Eater has been removed, then exit
If Not CheckEater(EaterNum) Then
    Exit Sub
End If

'randomize

Got(X) = Eaters(EaterNum).XPos
Got(Y) = Eaters(EaterNum).YPos
Speed = Eaters(EaterNum).Speed

MyX = Got(X)
MyY = Got(Y)

'look for several ants per turn (governed by foodrate)
If frmLife.SearchForFood(MyX, MyY, Speed, Eaters(EaterNum).Direction, 1) Then
    'found food
    Target(X) = MyX
    Target(Y) = MyY
    
    'move Eater to new position
    Grid(Eaters(EaterNum).XPos, Eaters(EaterNum).YPos) = HereEmpty
    
    'draw a blank dot on the old position
    Call frmLife.DrawDot(Got(X), Got(Y), vbWhite)
    
    'if the new position contains food, then eat it
    If Grid(Target(X), Target(Y)) = WorldDetails.Ant Then
        Eaters(EaterNum).FoodLevel = Eaters(EaterNum).FoodLevel + 1
        
        'reduce the ant population
        Call KillAnt(GetAntNumber(MyX, MyY))
    End If
    
    'move the ant eater to where the ant was
    Eaters(EaterNum).XPos = MyX
    Eaters(EaterNum).YPos = MyY
    
    Grid(MyX, MyY) = Eater
    Call frmLife.DrawDot(MyX, MyY, vbCyan)
    
    'reduce the eaters speed
    If Eaters(EaterNum).Speed > 1 Then
        Eaters(EaterNum).Speed = Eaters(EaterNum).Speed - 1 'CheckRange(1, MaxSpeedEater, Eaters(EaterNum).Speed - 1)
    End If
Else
    'just move the eater to a new position
    MyX = Got(X)
    MyY = Got(Y)
    
    'project co-ordinates to new target (in the given direction)
    Select Case Eaters(EaterNum).Direction
    Case Up
        MyY = Got(Y) - Speed
        If MyY < 0 Then
            MyY = (GridSize + 1) + MyY
        End If
    Case Right
        MyX = (Got(X) + Speed) Mod (GridSize + 1)
    Case Down
        MyY = (Got(Y) + Speed) Mod (GridSize + 1)
    Case Left
        MyX = Got(X) - Speed
        If MyX < 0 Then
            MyX = (GridSize + 1) + MyX
        End If
    End Select
    
    'copy over the old position
    Call frmLife.DrawDot(Got(X), Got(Y), vbWhite)
    Grid(Got(X), Got(Y)) = HereEmpty
    
    'increase the eaters speed
    If Eaters(EaterNum).Speed < MaxSpeedEater Then
        Eaters(EaterNum).Speed = Eaters(EaterNum).Speed + 1 'CheckRange(1, MaxSpeedEater, Eaters(EaterNum).Speed + 1)
    End If
    
    'move to the mew position
    Grid(MyX, MyY) = Eater
    Eaters(EaterNum).XPos = MyX
    Eaters(EaterNum).YPos = MyY
    Call frmLife.DrawDot(MyX, MyY, vbCyan)

    'change the direction
    Eaters(EaterNum).Direction = GetRndInt(Up, Left) '((Left - Up + 1) * Rnd + Up)
End If
    
DoEvents
Eaters(EaterNum).TickNum = Eaters(EaterNum).TickNum + 1
End Sub

Public Function CheckEater(Num As Integer) As Boolean
'This will check the stats of the selected Eater and make the
'appropiate changes.
Static test As Integer

CheckEater = True

If (Eaters(Num).FoodLevel >= Eaters(Num).FoodToSplit) Then 'And (Eaters(Num).FoodToSplit >= MinPropLevelEater)
    'MutateEater Eater
    Call MutateEater(Num)
End If

If (Eaters(Num).TickNum >= Eaters(Num).KillLevel) Or (Eaters(Num).Speed = 0) Then
    'delete the Eater and report that the Eater has been removed
    Call KillEater(Num)
    CheckEater = False
End If
End Function

Private Sub MutateEater(Num As Integer)
'This sub will MutateEater all the values of a Eater withing the
'mutation rate.

Dim NewVal As Integer
Dim Upperbound As Integer
Dim Lowerbound As Integer
Dim X As Integer
Dim Y As Integer
Dim Being As CreatureType

If (NumOfEaters >= MaxEaters) Then
    'if population pressure is too high, increase life span
    Eaters(Num).FoodLevel = 0
    Eaters(Num).KillLevel = Eaters(Num).KillLevel + FoodRateEater
    Exit Sub
End If

Being = Eaters(Num)

Upperbound = MutationRateEater
Lowerbound = -MutationRateEater

'evolve Eater with new settings and "age" the original Eater
With Being
    .FoodLevel = 0
    .TickNum = 0
    .Direction = ((Left - Up + 1) * Rnd + Up)
    Upperbound = Upperbound * (GridSize / 15)
    Lowerbound = Lowerbound * (GridSize / 15)
    .FoodToSplit = .FoodToSplit + ((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    If (.FoodToSplit < MinPropLevelEater) Then
        .FoodToSplit = MinPropLevelEater
    End If
    
    Upperbound = MutationRateEater * (GridSize / 4)
    Lowerbound = MutationRateEater * (GridSize / 4)
    .KillLevel = .KillLevel + ((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    If .KillLevel > MaxLifeSpanEater Then
        .KillLevel = MaxLifeSpanEater
    End If
    
    Upperbound = MutationRateEater
    Lowerbound = -MutationRateEater
    .Speed = .Speed + ((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    If .Speed > MaxSpeedEater Then
        .Speed = MaxSpeedEater
    End If
End With

'the old Eater produces slower and searches smaller, dies faster
Eaters(Num).FoodLevel = 0
Eaters(Num).FoodToSplit = Eaters(Num).FoodToSplit + 1
Eaters(Num).Speed = Eaters(Num).Speed - 1
Eaters(Num).KillLevel = Eaters(Num).KillLevel - 1
Eaters(Num).Direction = ((Left - Up + 1) * Rnd + Up)

'New Eater details
X = Eaters(EaterNum).XPos
Y = Eaters(EaterNum).YPos
If Not GetEmptyPos(X, Y) Then
    'if there is no available space for a new Eater, then don't
    'create one
    Eaters(Num).FoodLevel = 0
    Eaters(Num).KillLevel = Eaters(Num).KillLevel + FoodRateEater
    Exit Sub
End If

'create a new Eater
NumOfEaters = NumOfEaters + 1
ReDim Preserve Eaters(NumOfEaters)
Eaters(NumOfEaters) = Being
Eaters(NumOfEaters).XPos = X
Eaters(NumOfEaters).YPos = Y
Grid(X, Y) = Eater
Call frmLife.DrawDot(X, Y, vbBlack)
End Sub

Public Function GetAntNumber(XCo As Integer, YCo As Integer) As Integer
'this returns the ant's number in a given set of co-ordinates

Dim GotNum As Integer
Dim Counter As Integer

For Counter = 1 To NumOfAnts
    If (Ants(Counter).XPos = XCo) And (Ants(Counter).YPos = YCo) Then
        GotNum = Counter
        Exit For
    End If
Next Counter

GetAntNumber = GotNum
End Function
