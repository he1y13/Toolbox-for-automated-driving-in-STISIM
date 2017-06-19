VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRFQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright <2017> <Alexander Eriksson & Neville A Stanton, Transportation Research Group, University of Southampton, Uk>
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
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Tested on STISIM Drive Simulation Kernel - Build 3.07.04

'TO USE THIS MODULE THE "COLLECT TIME TO COLLISION DATA" OPTION MUST BE TURNED ON
'Function calls must be made to "" before any of the getter functions can be called
'all function calls should be made in the UPDATE function of the OM_Module
'the class must be defined, and initialized in the initialize function of the OM_Module




'A type containing vehicle information
Type SurrVehicle
    Acceleration As Single    ' vehicle acceleration in ft/s/s
    SixDOF As SixDOFPosition  'Six degree of freedom UDT used to define the vehicle’s inertial orientation and position. The definition  is found in chapter 3, Table 3.10.
    Speed As Single           'Current longitudinal or forward velocity of the vehicle, in feet/second.
    Xpos As Single            '.SixDOF.x
    ypos As Single            '.SixDOF.y
    range As Double           'Current range between this vehicle and the driver’s vehicle, in feet. This value takes into account both longitudinal and lateral positions.
    currentLane As Integer    'Current lane the vehicle is in, as calculated in the function "getCurrLane"
    direction As Integer      'the current vehicle heading: 1:Driver’s direction 2:Right to left across driver path 3:Opposing driver direction4:Left to right across driver path
    Length As Single          'Length of the vehicle, in feet.
    updated As Boolean        'Whether data has been written to the UDT
    LatPos As Single          'Lateral position with respect to the roadway.
End Type

'instantiation of 'vehicles'
Dim VAhead As SurrVehicle
Dim VNextAhead As SurrVehicle
Dim VAheadLeft As SurrVehicle
Dim VNextAheadLeft As SurrVehicle
Dim VAheadRight As SurrVehicle
Dim VNextAheadRight As SurrVehicle
Dim VBehind As SurrVehicle
Dim VBehindLeft As SurrVehicle
Dim VBehindRight As SurrVehicle

'instantiation of delimiters / counters
Dim closestinlane As Integer
Dim secondclosestinlane As Integer
Dim closestbehindinlane As Integer
Dim closestfrontright As Integer
Dim closestfrontleft As Integer
Dim secondclosestfrontright As Integer
Dim secondclosestfrontleft As Integer
Dim closestbackright As Integer
Dim closestbackleft As Integer

' instantiation of a temporary 'container-vehicle'
Dim container As SurrVehicle

'this function gets the vehicle closest ahead of the host vehicle
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getFirstAhead(ByRef vehicle As SurrVehicle) As Boolean
    If VAhead.updated Then
        vehicle = VAhead
        getFirstAhead = True
    Else
        vehicle = vehicle
        getFirstAhead = False
    End If
End Function
'this function gets the vehicle closest ahead of the host vehicle
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getNextAhead(ByRef vehicle As SurrVehicle) As Boolean
    If VNextAhead.updated Then
        vehicle = VNextAhead
        getNextAhead = True
    Else
        vehicle = vehicle
        getNextAhead = False
    End If
End Function
'this function gets the vehicle closest behind the host vehicle
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getBehind(ByRef vehicle As SurrVehicle) As Boolean
    If VBehind.updated Then
        vehicle = VBehind
        getBehind = True
    Else
        vehicle = vehicle
        getBehind = False
    End If
End Function
'this function gets the  closest vehicle ahead of the host vehicle on the left side
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getLeftAhead(ByRef vehicle As SurrVehicle) As Boolean
    If VAheadLeft.updated Then
        vehicle = VAheadLeft
        getLeftAhead = True
    Else
        vehicle = vehicle
        getLeftAhead = False
    End If
End Function
'this function gets the second closest vehicle ahead of the host vehicle on the left side
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getNextLeftAhead(ByRef vehicle As SurrVehicle) As Boolean
    If VNextAheadLeft.updated Then
        vehicle = VNextAheadLeft
        FirstAhead = True
    Else
        vehicle = vehicle
        getNextLeftAhead = False
    End If
End Function
'this function gets the closest vehicle ahead of the host vehicle on the right side
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getRightAhead(ByRef vehicle As SurrVehicle) As Boolean
    If VAheadRight.updated Then
        vehicle = VAheadRight
        getRightAhead = True
    Else
        vehicle = vehicle
        getRightAhead = False
    End If
End Function
'this function gets the second closest vehicle ahead of the host vehicle on the right side
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getNextRightAhead(ByRef vehicle As SurrVehicle) As Boolean
    If VNextAheadRight.updated Then
        vehicle = VNextAheadRight
        getNextRightAhead = True
    Else
        vehicle = vehicle
        getNextRightAhead = False
    End If
End Function
'this function gets the vehicle closest behind of the host vehicle on the right side
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getRightBehind(ByRef vehicle As SurrVehicle) As Boolean
    If VBehindRight.updated Then
        vehicle = VBehindRight
        getRightBehind = True
    Else
        vehicle = vehicle
        getRightBehind = False
    End If
End Function
'this function gets the vehicle closest behind of the host vehicle on the left side
'an instance of 'SurrVehicle' is passed by reference to the function
'and a bool is returned indicating whether a vehicle was found or not
Function getLeftBehind(ByRef vehicle As SurrVehicle) As Boolean
    If VBehindLeft.updated Then
        vehicle = VBehindLeft
        getLeftBehind = True
    Else
        vehicle = vehicle
        getLeftBehind = False
    End If
End Function



'A subroutine to update the UDT's containing vehicle information,
'should be called once every simulation loop, or when traffic needs to be queried in the Update function of the OM module.
'takes the same arguments as the Update function of STISIM plus a maxrange variable limiting the look-ahead distance (in ft)
Sub QueryTraffic(hostV As DYNAMICSParams, dv As OMDynamicVariables, NumEvents As Integer, EDes() As Integer, EDist() As Single, EIndex() As Integer, Events As SimEvents, maxrange As Integer)
    Dim i As Integer
    Call setMaxrange(maxrange)
    Call reset_vehicles
    For i = 0 To NumEvents
        Select Case EDes(i)
            Case EVENTDEFVEHICLE
                With Events.Vehicles(EIndex(i))
                    'write data to the container
                    container.ypos = .SixDOF.y
                    container.Xpos = EDist(i)
                    container.LatPos = .LatPos
                    container.Speed = .Speed
                    container.SixDOF = .SixDOF
                    container.Acceleration = .Acceleration
                    container.range = EDist(i)
                    container.currentLane = getCurrLane(CSng(.LatPos), dv)
                    container.direction = .direction
                    container.Length = .VehicleLength
                    container.updated = False
                    'if vehicles are in the same lane
                    If getCurrLane(CSng(.LatPos), dv) = getCurrLane(CDbl(hostV.YLanePos), dv) Then
                        'if it is in front of the host vehicle
                        If (EDist(i)) > 1 Then
                            If (EDist(i)) < closestinlane Then
                                closestinlane = EDist(i)
                                VAhead = container
                                VAhead.updated = True
                            Else
                                If (EDist(i)) < secondclosestinlane Then
                                    secondclosestinlane = EDist(i)
                                    VNextAhead = container
                                    VNextAhead.updated = True
                                End If
                            End If
                        'if it is behind
                        ElseIf (EDist(i)) < -1 And (EDist(i)) > closestbehindinlane Then
                            closestbehindinlane = (EDist(i))
                            VBehind = container
                            VBehind.updated = True
                        End If
                    'if it is to the right
                    ElseIf (getCurrLane(CSng(.LatPos), dv) - 1) = getCurrLane(CDbl(hostV.YLanePos), dv) Then
                        'if it is in front of the host vehicle
                        If (EDist(i)) > 1 Then
                            If (EDist(i)) < closestfrontright Then
                                closestfrontright = EDist(i)
                                VAheadRight = container
                                VAheadRight.updated = True
                            Else
                                If (EDist(i)) < secondclosestfrontright Then
                                    secondclosestfrontright = EDist(i)
                                    VNextAheadRight = container
                                    VNextAheadRight.updated = True
                                End If
                            End If
                            'if it is behind
                            ElseIf (EDist(i)) < -1 And (EDist(i)) > closestbackright Then
                                closestbackright = (EDist(i))
                                VBehindRight = container
                                VBehindRight.updated = True
                            End If
                    'if it is to the left
                    ElseIf (getCurrLane(CSng(.LatPos), dv) + 1) = getCurrLane(CDbl(hostV.YLanePos), dv) Then
                            If (EDist(i)) > 1 Then
                                If (EDist(i)) < closestfrontleft Then
                                    closestfrontleft = EDist(i)
                                    VAheadLeft = container
                                    VAheadLeft.updated = True
                            Else
                                If (EDist(i)) < secondclosestfrontleft Then
                                    secondclosestfrontleft = .SixDOF.x - hostV.x
                                    VNextAheadLeft = container
                                    VNextAheadLeft.updated = True
                                End If
                            End If
                            'if it is behind
                            ElseIf (EDist(i)) < -1 And (EDist(i)) > closestbackleft Then
                                closestbackleft = EDist(i)
                                VBehindLeft = container
                                VBehindLeft.updated = True
                            End If
                    End If
                End With
        End Select
    Next
End Sub

' A setter subroutine for the maxrange variable
'sets all variables to the maxrange
Private Sub setMaxrange(maxrange As Integer)
    closestinlane = maxrange
    secondclosestinlane = maxrange
    closestbehindinlane = maxrange * -1
    closestfrontright = maxrange
    closestfrontleft = maxrange
    secondclosestfrontright = maxrange
    secondclosestfrontleft = maxrange
    closestbackright = maxrange * -1
    closestbackleft = maxrange * -1
End Sub
'a setter subroutine for the vehicle UDT's
'resets all the UDT's to default, non-valid values
Private Sub reset_vehicles()
    Call resetStruct(VNextAhead)
    Call resetStruct(VAhead)
    Call resetStruct(VBehind)
    Call resetStruct(VNextAheadRight)
    Call resetStruct(VAheadRight)
    Call resetStruct(VNextAheadLeft)
    Call resetStruct(VAheadLeft)
    Call resetStruct(VBehindLeft)
    Call resetStruct(VBehindRight)
End Sub
'a setter subroutine for individual vehilce UDT's
'takes a SurrVehicle UDT by reference and resets it.
Private Sub resetStruct(ByRef vehicle As SurrVehicle)
    vehicle.ypos = -999
    vehicle.Xpos = -999
    vehicle.Speed = -999
    vehicle.Acceleration = -999
    vehicle.range = -999
    vehicle.currentLane = -999
    vehicle.direction = -999
    vehicle.Length = -999
    vehicle.updated = False
End Sub
'A function to calculate the number of lanes on the left side of the road
Private Function noOfLanes(dv As OMDynamicVariables) As Integer
        noOfLanes = Floor((dv.RoadWidthL + dv.RoadWidthR) / dv.LaneWidth)
End Function
'An implementation of the math.floor function
Private Function Floor(x As Variant) As Long
    If Round(x) > x Then
        Floor = Round(x) - 1
    Else
        Floor = Round(x)
    End If
End Function

'Determine car's current lane given its position on the road, lane and road widths


Public Function getCurrLane(ypos As Double, dv As OMDynamicVariables) As Integer
    'Initialise y and define it as the lateral coordinate with the road's left border as the zero point
    Dim y As Single
    y = ypos + dv.RoadWidthL
    'Driving on the left side of the road, return 1 for the leftmost lane and increase the number of the lane moving to the right
    If y < noOfLanes(dv) * dv.LaneWidth And y > 0 Then
        getCurrLane = 1 + Floor(y / dv.LaneWidth)
    'Driving on the right (opposite) side of the road use -1 for the leftmost lane and go more negative, as the car drives further
    'from the road's centreline
    ElseIf ypos > 0 And ypos < dv.RoadWidthR Then
        getCurrLane = -1 - Floor(ypos / dv.LaneWidth)
    'If the car is on the road, where there is only one lane on the left, set its current lane as 1
    ElseIf dv.LaneWidth = dv.RoadWidthL And y > 0 And y < (dv.RoadWidthR + dv.RoadWidthL) Then
        getCurrLane = 1
    'If the car is out of the road, set its current lane as 0
    Else
        getCurrLane = 0
    End If
End Function
