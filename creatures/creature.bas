Attribute VB_Name = "Creatures"
Option Explicit

' Constants for 'laws of physics'

' Curve fitting: y = a(x^3) + b(x^2) + cx + d

Type CurveFit
    A As Single         'Cubed coefficient
    B As Single         'Squared
    c As Single         'Multiplier
    D As Single         'Offset
End Type

' Fertility constants
Public FertilityFactor As Single        ' Biomass production per unit of fertility
Public FertilityMax As Single           ' Highest fertility in the world
Public FertilityMin As Single           ' Lowest fertility in the world
Public BiomassMaxFactor As Single       ' Max biomass as multiple of fertility

' Mutation and genetics
Public Mut As Single                    ' Mutation factor
Public MetAgeFactor As Single           ' Dividend in metabolism to max age calculation
Public SpeedFactor As Single            ' Multiple of metabolism that gives us max speed
Public HungerFactor As Single           ' Multiple of metabolism that gives us food consumption

Type Terrain
    Fertility As Single
    Biomass As Single
    Hostility As Single ' How rough is the terrain?
End Type

Public WorldSize As Integer

Public World() As Terrain   ' Our world. Will be ReDim at run time

Type Creature
    Species As Integer      ' What species are we?
    MaxSize As Single       ' How big can it get?
    Size As Single          ' How big is it now?
    MaxAge As Integer       ' How old can it get?
    Age As Integer          ' How old is it now?
    EatFactor As Single     ' How hungry before we eat?
    GrazeFactor As Single   ' How hungry before look elsewhere for food?
    SizeToBreed As Single   ' How heavy to breed?
    MaxFood As Single       ' How much food can it hold?
    Food As Single          ' How full is it now?
    Injury As Single        ' How injured is it?
    Rested As Single        ' How rested is it?
    Metabolism As Single    ' How fast is our metabolism?
    Health As Single        ' Overall health - 0 is dead.
    MaxSpeed As Single      ' How fast can it go?
    XSpeed As Single        ' How fast is it moving now?
    YSpeed As Single        ' How fast is it moving now?
    XGoal As Single         ' Where are we headed?
    YGoal As Single
    x As Single         ' X coordinate
    y As Single         ' Y coordinate
End Type

Public LifeSize As Integer

Public Life() As Creature    ' Our inhabitants

Public Ticks As Long
Public LiveCount As Integer

Sub MakePhysics()

' Fertility
    ' multiply fertility by F1. Fertility factor of 5 gives us .05 kg per tick
    FertilityFactor = 0.01
    ' World fertility will range from 2 to 6
    FertilityMin = 2
    FertilityMax = 6
    ' Cells can grow up to 5 * fertility units of biomass
    BiomassMaxFactor = 5

' Mutation
    ' Mutation factor of +/- 5%
    Mut = 0.1
    ' Max speed is 10 * metabolism
    SpeedFactor = 10
    MetAgeFactor = 500
    HungerFactor = 0.3
    
End Sub
' Set the initial world values
'
Sub MakeWorld()

    Dim i As Integer
    Dim j As Integer
    
    ' Just make everything flat and boring for now.
    
    ReDim World(WorldSize, WorldSize)
    
    For i = 0 To UBound(World, 1)
        ' Make fertility vary linearly from min to max by latitude
        For j = 0 To UBound(World, 2)
            World(i, j).Fertility = (FertilityMax - FertilityMin) * (UBound(World, 2) - j) / UBound(World, 2) + FertilityMin
            World(i, j).Biomass = 5 + Rnd() * 10
            World(i, j).Hostility = 5
        Next j
    Next i
    
End Sub

' Populate the world
'
Sub MakeCreatures()

    Dim i As Integer

    ReDim Life(LifeSize)
    
    For i = 0 To UBound(Life)
        Life(i).Species = 1     ' all the same - we don't care
        Life(i).MaxSize = 100   ' 100 kgs
        Life(i).EatFactor = 0.4 + Rnd * 0.2 ' How hungry before we eat?
        Life(i).GrazeFactor = 0.65 + Rnd * 0.2 ' How hungry before we eat?
        Life(i).SizeToBreed = 0.75 * Life(i).MaxSize      ' How heavy to breed?
        Life(i).MaxFood = 10    ' Can hold 10 kgs
        Life(i).Food = 2 + Rnd * 8       ' a little hungry
        Life(i).Injury = 0      ' not injured
        Life(i).Rested = 10     ' rested
        Life(i).Metabolism = 0.8 + Rnd * 0.4 ' sloth < us < cheetah
        Life(i).Health = 10     ' healthy
        Life(i).XSpeed = 0       ' Not moving
        Life(i).YSpeed = 0       ' Not moving
        Life(i).x = Rnd() * UBound(World, 1) ' random spot
        Life(i).y = Rnd() * UBound(World, 2)
        Life(i).XGoal = Life(i).x
        Life(i).YGoal = Life(i).y
        
        Call Constrain(i)       ' Calculate constrained characteristics: maxspeed, maxage
        
        Life(i).Age = Life(i).MaxAge * Rnd
        Life(i).Size = (Life(i).Age / Life(i).MaxAge) * Life(i).MaxSize * 2
        
    Next i
    LiveCount = UBound(Life) + 1

End Sub
' Determine genetic constraints - calculate dependent characteristics.
' At this point, just max age and max speed
'
Public Sub Constrain(c As Integer)

    Life(c).MaxSpeed = SpeedFactor * Life(c).Metabolism   ' meters per tick
    Life(c).MaxAge = (MetAgeFactor / Life(c).Metabolism) - 100

End Sub
' Return a luminosity-linearized shade (of green for now) for value v
' relative to lower and upper limits l an u.
'
Public Function shade(v As Single, l As Single, u As Single) As Integer

    Dim m As Single     ' Scale multiplier (slope)
    Dim v2 As Single    ' normalized v
    Dim s As Double
    
    m = 255 / (u - l)   ' Scale v by m
    v2 = (v - l) * m
    
    ' Boost lower end of scale
    s = (0.00002773 * v2 ^ 3) + (-0.01543121 * v2 ^ 2) + (3.10358599 * v2) + 7.27102171
    shade = Int(s)
    'shade = Int(v2)
End Function
