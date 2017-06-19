VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lateral"
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

Option Explicit
Dim prevRW As Integer
Dim prevLane As Integer
Dim prevNoLane As Integer
'A function to calculate the number of lanes on the left side of the road


Public Function noOfLanes(dv As OMDynamicVariables) As Integer
        noOfLanes = Floor((dv.RoadWidthL + dv.RoadWidthR) / dv.LaneWidth)
End Function
'An implementation of the math.floor function
Public Function Floor(x As Variant) As Long
    If Round(x) > x Then
        Floor = Round(x) - 1
    Else
        Floor = Round(x)
    End If
End Function
Public Sub checkLaneStatus(dv As OMDynamicVariables, ByRef LaneTarget As Integer, currentLane As Integer)
    If noOfLanes(dv) < prevNoLane Then
        If Not prevLane = 1 Then
            LaneTarget = LaneTarget - 1
        End If
    End If
    prevRW = dv.RoadWidthL
    prevNoLane = noOfLanes(dv)
    prevLane = currentLane
End Sub
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
        dv.DisplayStrings(5) = "onlyonelane"
    'If the car is out of the road, set its current lane as 0
    Else
        getCurrLane = 0
    End If
End Function

Public Function CalculateOffset(ypos As Double, dv As OMDynamicVariables, LaneTarget As Integer) As Double
    If LaneTarget < (noOfLanes(dv)) And LaneTarget > 0 Then
        
        CalculateOffset = ((LaneTarget - 0.5) * dv.LaneWidth - dv.RoadWidthL)
        prevLane = LaneTarget
    Else
        CalculateOffset = ((prevLane * dv.LaneWidth) - dv.RoadWidthL)
    End If
    
End Function

Private Sub Class_Initialize()
prevLane = 1
End Sub
