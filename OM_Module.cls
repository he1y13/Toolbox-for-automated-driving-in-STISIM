VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OM_Module"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright <2017> <Alexander Eriksson, Joost De Winter, Neville A Stanton, Transportation Research Group, University of Southampton, Uk>
'
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
'to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
'and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
'WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
'Acknowledgement: The authors conducted this work within the Marie Curie Initial Training Network (ITN) HF Auto - Human Factors of Automated Driving (PITN-GA-2013-605817).
'
'Note: the OM_Module shell class is accredited to Theodore J. Rosenthal and Jeff P. Chrstos (2013) and is provided with the STISIM Open Module feature.


'***************************************************************************'
'                                                                           '
'                                                                           '
'   CLASS:     OM_Module                                                    '
'                                                                           '
'   AUTHOR(S): Theodore J. Rosenthal                                        '
'              Jeff P. Chrstos                                              '
'                                                                           '
'   DATE:      January 23, 2013                                             '
'                                                                           '
'   COPYRIGHT:                                                              '
'                                                                           '
'   Copyright (c) by Systems Technology Incorporated:                       '
'                                                                           '
'   ALL RIGHTS RESERVED. The data and methods contained in this document    '
'   are proprietary to Systems Technology Incorporated.                     '
'                                                                           '
'   DESCRIPTION:                                                            '
'                                                                           '
'   Class containing all of the methods that the STISIM Drive Open Module   '
'   capability supports.                                                    '
'                                                                           '
'   Methods:                                                                '
'                                                                           '
'              AddNew - Uses the Open Module event in the SDL to initiate   '
'                       a new process                                       '
'       ControlInputs - Allows the user to modify the driver's control      '
'                       inputs                                              '
'            Dynamics - Allows the user to modify the vehicle dynamics      '
'         HandleCrash - Occurs anytime there is a crash                     '
'          Initialize - Handles any initialization that must occur before   '
'                       the simulation starts                               '
'   InitializeTerrain - Initialize the roadway terrain parameters before    '
'                       the simulation starts                               '
'             PostRun - Handle any final issues before the software exits   '
'    SavePlaybackData - Saves data into the playback file that will be      '
'                       needed by your Open Module during a playback        '
'             StartUp - Handle any processes immediately after the software '
'                       starts                                              '
'            Shutdown - Handles any shutdown processes immediately after a  '
'                       simulation run                                      '
'              Update - Allows the process to be updated each time through  '
'                       the simulation loop                                 '
'                                                                           '
'                                                                           '
'***************************************************************************'
'
'
Option Explicit
'
'
'***************************************************************************'
'                                                                           '
'                                                                           '
'   This section contains the definition of all variables that are          '
'   both global and local to this class                                     '
'                                                                           '
'                                                                           '
'***************************************************************************'
'
'
'   Create an instance of the windows tools class
'
'
Dim Tools As New TJRWinToolsCls
'
'
'   Create an instance of the graphics object
'
'
Dim Graphics As New TJR3DGraphics
'
'
'   Create an instance of the terrain object
'
'
Dim Terrain As New STI_3D_Terrain
'
'
'   Define the generic DirectSound sound object
'
'
Dim Ds As DirectSound8
Dim Dx As New DirectX8
'
'
'   Dimension all variables that are public to this class and the calling routine
'
'
Public BSAVData As String
Public ErrorMessage As String
Public LogFileHandle As Long
Public NewForm As Form
Public SaveControls As Integer
Public TextMessage As String
Public WillHandleCrash As Long
'
'
'   Create a type for the driver input information
'
'
Type DriverControlInputs
    Steer As Single                         ' Steering angle count from the controller card
    throttleIn As Single                      ' Throttle control count from the controller card
    throttleOut As Single
    BreakOut As Single                         ' Brake control count from the controller card
    BreakIn As Single
    Gear As Integer                         ' Current transmission gear
    Buttons As Long                         ' Current state of the driver's input buttons
    DInput As Long
    SteerOut As Single
    Clutch As Single
End Type
Dim Driver As DriverControlInputs
'
'
'   Define all vairables that will be global to this class
'
'
Dim Bool As Boolean                     ' Temporary boolean variable for use anytime a throw away boolean is needed
Dim DynVars As OMDynamicVariables       ' UDT containing STISIM Drive variables that change as the run progresses
Dim PlaybackMode As Integer             ' Flag specifying if it is real mode (0) or playback (1)
Dim St As String                        ' Temporary string variable for use anytime a throw away string is needed
Dim StaticVars As OMStaticVariables     ' UDT containing STISIM Drive variables that are fixed by the simulator
Dim WorldIndex() As Long                ' Index for the main graphics world
'
'
Type vehicle
    Acceleration As Single                  ' Vehicle acceleration
    BrakeModel As Long                      ' Brake light model handle
    index As Long                           ' Model index
    InitialHeading As Single                ' Vehicle initial heading when it is activated
    lat As Double                           ' Lane postion with respect to the roadway centerline
    Lon As Double                           ' Distance into the run (from the beginning)
    SixDOF As SixDOFPosition                ' Inertial orientation of the vehicle
    Speed As Single                         ' Vehicle speed
    SpinDuration As Single                  ' Time it takes to complete on wheel rotation
    SpinModel As Long                       ' Spin animation model handle
    SpinSpeed As Single                     ' Vehicle tire rotation speed
    VisFlag As Long                         ' Visibility flag
End Type
Dim v() As vehicle



Dim VAhead As SurrVehicle
Dim VNextAhead As SurrVehicle
Dim lahead As SurrVehicle
Dim rahead As SurrVehicle
Dim Lbehind As SurrVehicle
Dim Rbehind As SurrVehicle
Dim Traffic As TRFQuery

Dim cruise As controller
Dim follow As controller
Dim lat As controller
Dim linearplanner As Bezier
Dim lateral As lateral
Dim simulationtime As Long
Dim Freq As Integer
Dim thw As Double
Dim THWCycle As Integer
Dim vTarget As Single
Dim Vset As Single
Dim vCurr As Single
Dim mode As Integer 'mode of the longitudinal controller
Dim DriverInputPrev As Long 'previous button value
Dim ControlMode As Integer 'automation level
Dim LaneTarget As Integer
Dim currentLane As Integer
Dim calcThw As Double
Dim TTC As Double
Dim running As Boolean
'Allocate button-values for driver interaction
Dim cycleTHWBtn As Long
Dim ACCup As Long
Dim ACCDown As Long
Dim ActivateAcc As Long
Dim ActivateHAD As Long
Dim CancelAutomation As Long
Dim LeftLaneChange As Long
Dim RightLaneChange As Long
Dim buttonArray(0 To 8) As Long
'Array containing shutdown events
Dim eventArray(0 To 1) As String
Dim TOR As Integer
Dim TORCD As Double
Dim TORLim As Double

'Scaling of controller valuess
Dim ThrottleSF As Double
Dim BrakeSF As Double
Dim SteeringSF As Double
Dim ThrottleMin As Long
Dim BrakeMin As Long
Dim setup As Boolean

Dim socket As clientClass
Dim timecounter As Integer
Dim lahead_ As String
Dim rahead_ As String
Dim rbehind_ As String
Dim lbehind_ As String
Dim ahead As String
Dim msg As String



Public Function Update(dv As OMDynamicVariables, vehicle As DYNAMICSParams, Events As SimEvents, NumEvents As Integer, EDist() As Single, EDes() As Integer, EIndex() As Integer, SimSounds() As SoundEffects) As Boolean



Call Traffic.QueryTraffic(vehicle, dv, NumEvents, EDes(), EDist(), EIndex(), Events, 1000)

If Traffic.getFirstAhead(VAhead) Then

    If VAhead.range < 328 Then
        ahead = CStr(CInt(VAhead.range))
    Else
        ahead = "-999"
    End If
Else
    ahead = "-999"
End If
If Traffic.getRightAhead(rahead) Then
    If rahead.range < 328 And rahead.range > 0 Then
        rahead_ = CStr(CInt(rahead.range))
    Else
        rahead_ = "-999"
    End If
Else
    rahead_ = "-999"
End If
If Traffic.getLeftAhead(lahead) Then
    If lahead.range < 328 And lahead.range > 0 Then
        lahead_ = CStr(CInt(lahead.range))
    Else
        lahead_ = "-999"
    End If
Else
    lahead_ = "-999"
End If

If Traffic.getRightBehind(Rbehind) Then
    dv.DisplayStrings(5) = "HostDist: " + CStr(dv.Distance) + " LD: " + CStr(Rbehind.Xpos)
    If CDbl(Rbehind.Xpos) > -328# And CDbl(Rbehind.Xpos) < 0# Then
        rbehind_ = CStr(CInt(Rbehind.Xpos))
    Else
        rbehind_ = "-999"
    End If
Else
    rbehind_ = "-999"
End If
If Traffic.getLeftBehind(Lbehind) Then
    If Lbehind.Xpos > -328# And Lbehind.Xpos < 0# Then
        lbehind_ = CStr(CInt(Lbehind.Xpos))
    Else
        lbehind_ = "-999"
    End If
Else
    lbehind_ = "-999"
End If

'set the starting lane of the automation to the current lane before any control actions are carried out
currentLane = lateral.getCurrLane(CDbl(vehicle.YLanePos), dv)
If Not running Then
LaneTarget = currentLane
running = True
End If

'this line sets a message to be sent over TCP, not currently used but here for reference.
msg = CStr(vehicle.YLanePos) + "," + CStr(lateral.getCurrLane(CDbl(vehicle.YLanePos), dv)) + "," + CStr(thw) + "," + ahead + "," + rbehind_ + "," + rahead_ + "," + lbehind_ + "," + lahead_ + "," + CStr(lateral.noOfLanes(dv)) + "," + CStr(CInt(vehicle.U)) _
+ "," + CStr(vehicle.Omega) + "," + CStr(ControlMode) + "," + CStr(TORCD) + "," + CStr(CLng(dv.Distance)) + ":"
'making sure the interval in which data is sent does not cause issues with STISIM
'If timecounter > 4 Then
        'socket.Send msg
        'timecounter = 0
'End If

simulationtime = simulationtime + 1
timecounter = simulationtime

'assesses button presses from the driver
Call LaunchFunctionOnBtnPress(DriverInputPrev, CLng(Driver.Buttons))
DriverInputPrev = Driver.Buttons
'checks whether the automation shutdown event is to be triggered.
Call AutomationShutdown(CDbl(dv.Distance), simulationtime, ControlMode, Freq)


''-------
'' This section's code may be uncommented to force the automation to abide the speed limit, this will hinder the driver from overriding
'' the acc set speed parameters to velocities above the speed limit
'If vTarget > CSng(dv.SpeedLimit) Then
'    vTarget = CSng(dv.SpeedLimit)
'End If
''--------


'This calls the traffic querying update function that queries the traffic and updates the position of the traffic objects.
' it takes the same arguments as the Update function of the OM_Module with an additional 'Max Range' parameter that determinesss
' how far ahead it considers traffic (currently set to 500ft)

Dim x As Long
vCurr = vehicle.U
    If Traffic.getFirstAhead(VAhead) Then
        If VAhead.range < 328 And VAhead.range > 0 Then 'approx 100m
            If CSng(VAhead.Speed) < vTarget And vehicle.U > 0 Then
                dv.DisplayStrings(2) = "THW = " + CStr(CSng(VAhead.range / vehicle.U)) + " Desired HW: " + CStr(thw)
                calcThw = VAhead.range / vehicle.U
                TTC = VAhead.range / Abs(vehicle.U - VAhead.Speed)
                If (VAhead.range / vehicle.U) < (thw * 1.15) Then
                    mode = 2
                    dv.DisplayStrings(1) = "vehicle detected - in range - following"
                ElseIf Not Abs(vTarget - vehicle.U) > 3.28 * 3.5 Then
                    mode = 1
                    dv.DisplayStrings(1) = "vehicle detected - in range -  approaching"
                Else
                    mode = 3
                    dv.DisplayStrings(1) = "vehicle detected - in range - Adapting approach speed"
                End If
            Else
                If Abs(vTarget - vehicle.U) > 3.28 * 3.5 Then
                    mode = 3
                    dv.DisplayStrings(1) = "vehicle detected - too fast - Adapting speed"
                Else
                    mode = 1
                    dv.DisplayStrings(1) = "vehicle detected - too fast"
                End If
            End If
        Else
            If Abs(vTarget - vehicle.U) > 3.28 * 3.5 Then
                mode = 3
                dv.DisplayStrings(1) = "vehicle detected - out of range- Adapting speed"
            Else
                mode = 1
                dv.DisplayStrings(1) = "vehicle detected - out of range - cruising"
            End If
        End If
    Else
        If Abs(vTarget - vehicle.U) > 3.28 * 3.5 Then
            mode = 3
            dv.DisplayStrings(1) = "No vehicle detected - Adapting speed"
        Else
            mode = 1
            dv.DisplayStrings(1) = "No vehicle detected - cruising"
        End If
        calcThw = -999
        TTC = -999
    End If
 
 
'determines controller actions and assigning tasks to controllers based on the automation mode set by the driver
If ControlMode >= 1 Then
Select Case (mode)
    Case (0)
        Driver.throttleOut = Driver.throttleIn
        Driver.BreakOut = Driver.BreakIn
    Case (1)
        x = cruise.Control(vTarget, vehicle.U, Freq)
    Case (2)
    If VAhead.range > 0 And Traffic.getFirstAhead(VAhead) Then
        x = follow.Control((VAhead.range / vehicle.U) * 100, thw * 100, Freq)
    End If
    Case (3)
        If Not linearplanner.interpolating Then
            Call linearplanner.PlanTrajectory(vehicle.U, vTarget, simulationtime, 4.5, Freq)
            x = cruise.Control(linearplanner.UpdateLinear(simulationtime), vehicle.U, Freq)
        Else
            x = cruise.Control(linearplanner.UpdateLinear(simulationtime), vehicle.U, Freq)
        End If
End Select
    If (x > 0) Then
        Driver.throttleOut = (x * ThrottleSF) + ThrottleMin
        Driver.BreakOut = 0
    ElseIf (x < 0) Then
        Driver.throttleOut = 0
        Driver.BreakOut = (x * BrakeSF) + BrakeMin
    End If
End If
If ControlMode = 2 Then
    
    lat.SetKp (700 + (0.065 * ((vehicle.U * 0.68) ^ 1.7)))
    dv.DisplayStrings(3) = "PO: " + CStr(lateral.CalculateOffset(CDbl(vehicle.YLanePos), dv, LaneTarget)) + " CP: " + CStr(CSng(vehicle.YLanePos)) + " CL: " + CStr(lateral.getCurrLane(CDbl(vehicle.YLanePos), dv)) + "DL: " + CStr(LaneTarget)
    Call lateral.checkLaneStatus(dv, LaneTarget, lateral.getCurrLane(CDbl(vehicle.YLanePos), dv))
    Driver.SteerOut = SteeringSF * (lat.Control(lateral.CalculateOffset(CDbl(vehicle.YLanePos), dv, LaneTarget), CSng(vehicle.YLanePos), Freq))

End If



BSAVData = CStr(dv.Distance) + "," + CStr(ControlMode) + "," + CStr(mode) + "," + CStr(thw) + "," + CStr(calcThw) + "," + CStr(vTarget) + "," + CStr(vehicle.U) + "," + CStr(TTC) + "," + CStr(lateral.CalculateOffset(CDbl(vehicle.YLanePos), dv, LaneTarget)) + "," + CStr(vehicle.YLanePos)
'


On Error GoTo ErrorOccurred
'
'
'   Make the dynamic variables available to all other methods
'
'
DynVars = dv
'
'
'   Setup the return from function
'
'
Update = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("Update")
Update = False
'
'
End Function
'manages the externally paced shutdown event set in "shutdownevents.txt"
Public Sub AutomationShutdown(position As Double, simulationtime As Long, CM As Integer, Freq As Integer)
    Dim arr() As String
    arr = Split(eventArray(0), ";")
    If CDbl(arr(0)) > 0 Then
        If CDbl(arr(0)) >= position And CDbl(arr(0)) < position + 50 And TOR = 0 Then
            TOR = 1
            TORCD = CDbl(simulationtime) / CDbl(Freq)
        ElseIf (CDbl(simulationtime) / CDbl(Freq)) > (TORCD + CDbl(arr(1))) And (CDbl(simulationtime) / CDbl(Freq)) < (TORCD + CDbl(arr(2))) Then
            'vTarget = 0
            'If vCurr = vTarget Then
            '    Call cancelAD
            'End If
            Call cancelAD 'comment this away when uncommenting the above lines for a safe stop before disengaging AD
            TOR = 0
        End If
        
    End If
End Sub
'manages button presses
Public Sub LaunchFunctionOnBtnPress(InputPrev As Long, button As Long)
    If BtnTriggered(InputPrev, button) Then
            Select Case button
            Case (ACCup)
                Call increaseV
            Case (ACCDown)
                Call DecreaseV
            Case (RightLaneChange)
                LaneTarget = LaneTarget + 1
            Case (LeftLaneChange)
                 LaneTarget = LaneTarget - 1
            Case (cycleTHWBtn)
                Call cycleTHW
            Case (ActivateAcc)
                Call toggleACC
            Case (ActivateHAD)
                Call toggleHAD
            Case (CancelAutomation)
                Call cancelAD
          End Select
    End If
End Sub
'cycles between different time-headways
Public Sub cycleTHW()
    THWCycle = THWCycle + 1
    If THWCycle = 4 Then
        THWCycle = 1
    End If
    If THWCycle = 1 Then
        thw = 1
    ElseIf THWCycle = 2 Then
        thw = 1.5
    ElseIf THWCycle = 3 Then
        thw = 2
    End If
End Sub
'toggles adaptive cruise control
Public Sub toggleACC()
    If ControlMode = 1 Then
        ControlMode = 0
    Else
        ControlMode = 1
        vTarget = CSng(vCurr)
End Sub
Public Sub toggleHAD()
    If ControlMode = 2 Then
        ControlMode = 0
    Else
        ControlMode = 2
        vTarget = CSng(vCurr)
        LaneTarget = currentLane
    End If
End Sub
Public Sub cancelAD()
    ControlMode = 0
End Sub
Public Function BtnTriggered(InputPrev As Long, button As Long) As Boolean
    'On release of the button return True:
        If InputPrev <> button Then
            BtnTriggered = True
        Else: BtnTriggered = False
        End If
    End Function

Public Sub increaseV()
    Dim temp As Single
    temp = vTarget / CSng(1.09728) '1.09728
        If temp >= 0 And vTarget <= 130 Then
            temp = ((Fix(temp / CSng(5))) * CSng(5))
            vTarget = (temp + CSng(5)) * CSng(1.09728)
        End If
End Sub
Public Sub DecreaseV()
    Dim temp As Single
    temp = vTarget / CSng(1.09728) '1.09728
    If temp >= 0 Then
        temp = ((Fix(temp / CSng(5))) * CSng(5))
        vTarget = (temp - CSng(5)) * CSng(1.09728)
    End If
End Sub

'***************************************************************************'
'                                                                           '
'                                                                           '
'   This section contains the public methods that the main STISIM Drive     '
'   modules can access during the course of a simulation run                '
'                                                                           '
'                                                                           '
'***************************************************************************'
'
'
Public Function AddNew(OMVars As OMParameters) As Boolean
'
'
'   Function for adding a new interactive Open Module event
'
'
'   Parameters:
'
'   OMVars - User defined type containing the parameters for the given Open Module being acted on
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Setup the return from function
'
'
AddNew = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("AddNew")
AddNew = False
'
'
End Function

Public Function ControlInputs(Dyn As DYNAMICSParams, Steering As Single, Throttle As Single, Brake As Single, Gear As Integer, DInput As Long) As Boolean



With Driver
    .BreakIn = Brake
    .throttleIn = Throttle
    .Steer = Steering
    .Gear = Gear
    .Buttons = DInput
End With

If ControlMode = 2 Then
    Steering = Driver.SteerOut
    Throttle = Driver.throttleOut
    Brake = Driver.BreakOut
ElseIf ControlMode = 1 Then
    Throttle = Driver.throttleOut
    Brake = Driver.BreakOut '
ElseIf ControlMode = 0 Then
    Steering = Driver.Steer
    Throttle = Driver.throttleIn
    Brake = Driver.BreakIn

End If

On Error GoTo ErrorOccurred


'
ControlInputs = True
Exit Function

ErrorOccurred:
'
'
ErrorMessage = ProcessError("ControlInputs")
ControlInputs = False
'
'
End Function

Public Function Dynamics(Dyn As DYNAMICSParams) As Boolean
On Error GoTo ErrorOccurred
Dynamics = True
Exit Function
ErrorOccurred:
'
'
ErrorMessage = ProcessError("Dynamics")
Dynamics = False
'
'
End Function

Public Function HandleCrash(Override As Integer, Events As SimEvents, CrashEvent As Integer, EventIndex As Integer) As Boolean
'
'
'   Function for handling all Open Module action if there is a crash during the simulation run
'
'
'   Parameters:
'
'      Override - Parameter defining how STISIM Drive will handle the crash when this
'                 method returns control to it
'        Events - UDT that contains the parameter settings for each supported event
'    CrashEvent - Event designator for the event that caused the crash
'    EventIndex - Index specifying which instance of the crash event caused the crash
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
running = False 'resetting start lane for crashes.
On Error GoTo ErrorOccurred
'
'
'   Setup the return from function
'
'
Override = 0
HandleCrash = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("HandleCrash")
HandleCrash = False
'
'
End Function

Public Function Initialize(SV As OMStaticVariables, ParamFile As String, WorldIndexIn() As Long, GraphicsIn As TJR3DGraphics) As Boolean
'
'
'client stuff
'Set socket = New clientClass
setup = False
'socket.Connect "192.168.1.5", "5004"
'
Dim FileNum As Integer
Dim i As Integer
Dim NumVerts As Integer


Set Traffic = New TRFQuery
Set follow = New controller
Set cruise = New controller
Set lat = New controller
Set linearplanner = New Bezier
Set lateral = New lateral
cruise.SetKd (100)
cruise.SetKi (200)
cruise.SetKp (5000)
follow.SetKd (5)
follow.SetKi (20)
follow.SetKp (3000)
lat.SetKp (700)
lat.SetKi (150)
lat.SetKd (400)
Freq = 20
thw = 1.5
vTarget = 60
ControlMode = 2
LaneTarget = 2
running = False
Call readBtnAss
Call readShutdownEvents
cycleTHWBtn = CLng(buttonArray(0))
ACCDown = CLng(buttonArray(1))
ACCup = CLng(buttonArray(2))
ActivateAcc = CLng(buttonArray(3))
ActivateHAD = CLng(buttonArray(4))
CancelAutomation = CLng(buttonArray(5))
LeftLaneChange = CLng(buttonArray(6))
RightLaneChange = CLng(buttonArray(7))

'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Make the static variables available to all other methods
'
'
StaticVars = SV
'
'
'   Assign references to the main simulator's graphics object and the world indices so that they can be used throughout the OM
'
'
Set Graphics = GraphicsIn
NumVerts = UBound(WorldIndexIn)
ReDim WorldIndex(NumVerts)
For i = 0 To NumVerts
    WorldIndex(i) = WorldIndexIn(i)
Next
'
'
'   Only act if it is the center display system
'
'
If StaticVars.DisplaySystem = "CenterDisplay" Then
    '
    '
    '   Setup any labels that will be used to display data in the STISIM Drive runtime window display
    '
    '
    With SV
        .DisplayStrings(1) = "Lead Vehicle Distance"
        .DisplayStrings(2) = "Left Side Lead vehicle"
        .DisplayStrings(3) = "Right Side Lead Vehicle"
        .DisplayStrings(4) = "Rear Left"
        .DisplayStrings(5) = "Rear Right"
    End With
    '
    '
    '   If there is an initialization file specified then do the initializing
    '
    '
    If Len(ParamFile) Then
        '
        '
        '   Handle the initial parameter file
        '
        '
        FileNum = FreeFile(0)
        Open ParamFile For Input As #FileNum
        Close FileNum
    End If
End If
'
'
'   Setup the return from function
'
'
Initialize = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("Initialize")
Initialize = False
'
'
End Function

Public Function InitializeTerrain(TerrainIn As STI_3D_Terrain) As Boolean
'
'
'   Function for setting the terrain parameters
'
'
'   Parameters:
'
'   TerrainIn - Reference to the terrain object that is used by the main simulation loop
'
'   Returns:
'
'   True if the terrain data was setup correctly, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Initialize the terrain object
'
'
Set Terrain = TerrainIn
'
'
'   Setup the return from function
'
'
InitializeTerrain = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("InitializeTerrain")
InitializeTerrain = False
'
'
End Function

Public Function PostRun(Comments As String, DriverName As String, RunNumber As String, DriverID As String) As Boolean
'
'
'   Function for handling anything before the software exists
'
'
'   Parameters:
'
'     Comments - Comments entered in the subject information form
'   DriverName - Name of the driver from the subject information form
'    RunNumber - Run number entered in the subject information form
'     DriverID - ID entered from the subject information form
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Release the objects that were created
'
'
Set Ds = Nothing
Set Dx = Nothing
Set Tools = Nothing
'
'
'   Setup the return from function
'
'
PostRun = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("PostRun")
PostRun = False
'
'
End Function

Public Function SavePlaybackData(PlaybackData() As Single, PlaybackString As String) As Boolean
'
'
'   Function for specifying any OM data that will be stored as part of a playback file
'
'
'   Parameters:
'
'     PlaybackData - Array containing the data that will be saved
'   PlaybackString - String containing string data that will be saved
'
'   Returns:
'
'   True if everything worked fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Setup the return from function
'
'
SavePlaybackData = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("SavePlaybackData")
SavePlaybackData = False
'
'
End Function

Public Function Shutdown(RunCompleted As Integer) As Boolean
'
'
'   Function for handling Open Module processes immediately after a simulation run has ended
'
'
'   Parameters:
'
'   RunCompleted - Flag specifying if the run completed successfully or not
'
'                    0 - Aborted before start of run
'                    1 - Run completed successfully
'                  > 1 - Aborted during the run
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
'socket.Send "!EOC!"
On Error GoTo ErrorOccurred
'
'
'   Release the objects that were created
'
'
Set Graphics = Nothing
Set Terrain = Nothing
'
'
'   Setup the return from function
'
'
Shutdown = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("Shutdown")
Shutdown = False
'
'
End Function
'this is the cache read function
Public Sub readBtnAss()
    Dim f As Long
    
    On Error GoTo OpenFileError                 'Catch errors

    Dim testString As String

       
    f = FreeFile()
    Open "C:\STISIM\ButtonAssignment.txt" For Input As #f  'Open the file
    
    Dim counter As Integer                      'Create a counter
    
    For counter = 0 To 7                       'Cycle through the cache
        Line Input #f, testString               'Read each line of the file into temporary string
        buttonArray(counter) = testString        'Save each line of the file into cacheArray
    Next
    
    Close #f
    Exit Sub
OpenFileError:                                  'If there's an error
    Close #f                                    'Close the file
    readBtnAss                              'Try again
     
End Sub
Public Sub readShutdownEvents()
    Dim f As Long
    
    On Error GoTo OpenFileError                 'Catch errors

    Dim testString As String

       
    f = FreeFile()
    Open "C:\STISIM\shutdownevents.txt" For Input As #f  'Open the file
    
    Dim counter As Integer                      'Create a counter
    
    For counter = 0 To 1                       'Cycle through the cache
        Line Input #f, testString               'Read each line of the file into temporary string
        eventArray(counter) = testString        'Save each line of the file into cacheArray
    Next
    
    Close #f
    Exit Sub
OpenFileError:                                  'If there's an error
    Close #f                                    'Close the file
    readBtnAss                              'Try again
     
End Sub

Public Function StartUp(Config As GAINSParams, BackForm As Object, SV As OMStaticVariables, UseNew As Boolean, PlaybackData() As Single, PlaybackString As String) As Boolean
'
'
'   Function for handling Open Module processes immediately after the software is started
'
'
'   Parameters:
'
'           Config - Configuration file parameters
'         BackForm - Current STISIM Drive back ground form
'               SV - User defined type containing simulation static variables
'           UseNew - Flag specifying if a new background form will be used (True) or not (False)
'   PlaybackData() - Array containing any data that is being transfered from the playback file back into your module
'   PlaybackString - String containing any string data that is being transfered from the playback file back into your module
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling



ThrottleSF = (Config.IThrottleMax - Config.IThrottleMin) / 65000
BrakeSF = (Config.IBrakeMax - Config.IBrakeMin) / 65000
SteeringSF = Config.FKSW / 0.014
ThrottleMin = Config.IThrottleMin
BrakeMin = Config.IBrakeMin

On Error GoTo ErrorOccurred
'
'
'   Setup the sound object
'
'
If SV.SoundOn = True Then
    Set Ds = Dx.DirectSoundCreate(vbNullString)
    Call Ds.SetCooperativeLevel(SV.ActiveHandle, DSSCL_NORMAL)
End If
'
'


If Config.ITime2Col = 0 Then
    GoTo ErrorOccurred
End If
If Not Config.IControlFlag = 1 Then
    GoTo ErrorOccurred
End If
    
    


'
UseNew = False
StartUp = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'

If Config.ITime2Col = 0 Then
    ErrorMessage = TTCData(" TTCDataCollectionError ", " Enable collect time to collision data collection in the STISIM configuration file ")
    StartUp = False
ElseIf Not Config.IControlFlag = 1 Then
ErrorMessage = TTCData(" No Joystick found ", " This version of the automation will only run with a joystick/steering-wheel plugged in ")
    StartUp = False
Else

ErrorMessage = ProcessError("StartUp")
StartUp = False

End If
'
'
End Function


'
'
'**************************************************************************'
'                                                                          '
'                                                                          '
'   This section contains private routines that are need by and only       '
'   available to the methods in this module                                '
'                                                                          '
'                                                                          '
'**************************************************************************'
Private Function TTCData(ErrorType As String, Description As String) As String
St = "Simulation run aborted! An error has occurred in Open Module" & vbCrLf & vbCrLf
St = St & "Error: " & vbTab & ErrorType & vbCrLf & vbCrLf
St = St & "Description:" & vbTab & Description & vbCrLf & vbCrLf
TTCData = St
Bool = Tools.WriteToTJRFile(StaticVars.LogFileHandle, TTCData)
End Function
Private Function ProcessError(ModuleName As String) As String
'
'
'   Routine for adding information to the error message that will be returned when the
'   Open Module encounters problems
'
'   Parameters:
'
'   ModuleName - Name of the method where the error occured
'
'   Returns:
'
'   New error message including number and description and other information
'
'
'
'
'   Build the error message
'
'
St = "Simulation run aborted! An error has occurred in Open Module " & ModuleName & ":" & vbCrLf & vbCrLf
St = St & "Error number:" & vbTab & Trim(Str(Err.Number)) & vbCrLf
St = St & "Description:" & vbTab & Err.Description & vbCrLf
St = St & "Error source:" & vbTab & Err.Source & vbCrLf
ProcessError = St & "Last DLL Error:" & vbTab & Err.LastDllError & vbCrLf
Bool = Tools.WriteToTJRFile(StaticVars.LogFileHandle, ProcessError)
'
'
End Function

