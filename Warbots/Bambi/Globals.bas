Attribute VB_Name = "Globals"
Option Explicit
' There is one required global object: MyBot.
' This is the robot object which provides the interface to
' the robot server. In VB5, you can use the Object Browser
' (F2 Key) to view the methods available.

Global MyBot As RobotLink

' User defined globals:
' These are 'global' to this form. Use these or add your own.
' They are not required except as used by your application

Global speed As Integer
Global dir As Integer
Global scanres As Single
Global scandir As Single
Global flight As Long
Global reverse As Long
Global ccw As Integer
Global range As Single

' Status Tracking Variables - we like to do things in chunks
Global reloading As Single  ' Can't shoot while reloading
Global attacking As Integer ' We're attacking this robot
Global hunting As Boolean

' Information about a sighting of an enemy 'bot
Private Type sighting
    x As Single             ' most probable x,y
    y As Single
    x1 As Single            ' alternate x,y based on barrel heat
    y1 As Single
    x2 As Single            ' alternate x,y based on barrel heat
    y2 As Single
    t As Single
End Type

Private Type history
    depth As Integer        ' Number of sightings in array
    alive As Boolean
    lastdir As Single       ' Last non-qualified bearing
    lastrange As Single     ' Last non-qualified distance
    lastseen As Single      ' Time last seen (or pinged by)
    lastx As Single
    lasty As Single
    verified As Boolean     ' Verified on last attempt?
    scanres As Single       ' last resolution for this one
    vx As Single
    vy As Single
    s(3) As sighting
End Type

Global enemies(4) As history
Global closest As Integer          ' Enemy who is closest to us
Global standoff As Single          ' Distance to nearest enemy
Global btemp As Single             ' barrel temp
Global lastbtime As Single         ' time of last barrel temp calculation
Global nextcourse As Single        ' Next time to look at course
Global shells As Integer           ' remaining shells in this clip

