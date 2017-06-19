VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bezier"
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


'Initialise parameters to store variables required for trajectory planning
Private V0 As Single
Private V1 As Double
Public interpolating As Boolean
Private TStart As Long
Private T As Double
Private Duration As Double
Private freq_ As Integer


'The function receives arguments with the starting and final parameters and records them to be used for trajectory computation
Public Sub PlanTrajectory(Vs As Single, Vf As Single, TSim As Long, aComfortable As Single, Freq As Integer)
    V0 = Vs
    V1 = Vf
    TStart = TSim
    Duration = ManouverDuration(Abs(Vf - Vs), aComfortable)
    interpolating = True
    freq_ = Freq
End Sub


'Computes the time taken to equalise the speeds with the specified comfortable acceleration and adds a 5 sec margin
Private Function ManouverDuration(dv As Single, aComfortable As Single) As Double
    ManouverDuration = dv / aComfortable
End Function


'Calculates the progress of time within the time interval
Private Function TimeInterval(TCurrent As Long) As Double
    
    TimeInterval = (TCurrent - TStart) / (Duration * freq_)
End Function

'Use the TimeInterval to find the corresponding target speed within the time frame and return the value
Public Function UpdateLinear(TCurrent As Long) As Double
    'Get the T value expressing how far in the manouver the car is
    T = TimeInterval(TCurrent)
    'If the T is smaller than 0 (time before the start of the manouver), use the starting speed, V0,
    'and set the interpolating flag to False
    If (T < 0) Then
        UpdateLinear = V0
        interpolating = False
    'If the T is greater or equal to 1 (the manouver has reached the end, time-wise), use the final speed, V1,
    'and set the interpolating flag to False
    ElseIf (T >= 1) Then
        interpolating = False
        UpdateLinear = V1
    'Otherwise, if the manouver is being performed, return the updated instantaneous target speed
    'computed from the 1st order (linear) Bezier trajectory
    Else
        UpdateLinear = ((1 - T) * V0 + T * V1)
    End If
End Function

'Checks if manouver's final speed, V1, if it has changed and differs by more than 10% frmo the previous V1 value
Public Function V1Changed(V1new As Double) As Boolean

    If (Abs(V1new - V1) > 0.1 * V1) Then
        V1Changed = True
    Else
        V1Changed = False
    End If

End Function

'Set interpolating to False at the beginning
Private Sub Class_Initialize()
    
    interpolating = False
    
End Sub


