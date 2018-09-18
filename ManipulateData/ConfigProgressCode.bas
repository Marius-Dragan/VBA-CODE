
Option Explicit

Dim Cancelled As Boolean, showTime As Boolean, showTimeLeft As Boolean
Dim startTime As Long
Dim BarMin As Long, BarMax As Long, BarVal As Long

#If VBA7 And Win64 Then
Private Declare PtrSafe Function GetTickCount Lib "Kernel32" () As LongPtr

'Title will be the title of the dialogue.
'Status will be the label above the progress bar, and can be changed with SetStatus.
'Min is the progress bar minimum value, only set by calling configure.
'Max is the progress bar maximum value, only set by calling configure.
'CancelButtonText is the caption of the cancel button. If set to vbNullString, it is hidden.
'optShowTimeElapsed controls whether the progress bar computes and displays the time elapsed.
'optShowTimeRemaining controls whether the progress bar estimates and displays the time remaining.
'calling Configure sets the current value equal to Min.
'calling Configure resets the current run time.
Public Sub Configure(ByVal Title As String, ByVal status As String, _
                     ByVal Min As Long, ByVal Max As Long, _
                     Optional ByVal CancelButtonText As String = "Cancel", _
                     Optional ByVal optShowTimeElapsed As Boolean = True, _
                     Optional ByVal optShowTimeRemaining As Boolean = True)

#Else
Private Declare Function GetTickCount Lib "Kernel32" () As Long
Public Sub Configure(ByVal Title As String, ByVal status As String, _
                     ByVal Min As Long, ByVal Max As Long, _
                     Optional ByVal CancelButtonText As String = "Cancel", _
                     Optional ByVal optShowTimeElapsed As Boolean = True, _
                     Optional ByVal optShowTimeRemaining As Boolean = True)
#End If

    Me.Caption = Title
    lblStatus.Caption = status
    BarMin = Min
    BarMax = Max
    BarVal = Min
    CancelButton.Visible = Not CancelButtonText = vbNullString
    CancelButton.Caption = CancelButtonText
    startTime = GetTickCount
    showTime = optShowTimeElapsed
    showTimeLeft = optShowTimeRemaining
    lblRunTime.Caption = ""
    lblRemainingTime.Caption = ""
    Cancelled = False
End Sub

'Set the label text above the status bar
Public Sub SetStatus(ByVal status As String)
    lblStatus.Caption = status
    DoEvents
End Sub

'Set the value of the status bar, a long which is snapped to a value between Min and Max
Public Sub SetValue(ByVal value As Long)

Dim progress As Double, runTime As Long

    If value < BarMin Then value = BarMin
    If value > BarMax Then value = BarMax
    BarVal = value
    progress = (BarVal - BarMin) / (BarMax - BarMin)
    ProgressBar.Width = 352 * progress 'Modify this to reflect the changes in the progress bar to match the width
    lblPercent = Int(progress * 10000) / 100 & "%"
    runTime = GetRunTime()
    If showTime Then lblRunTime.Caption = "Time Elapsed: " & GetRunTimeString(runTime, True)
    If showTimeLeft And progress > 0 Then _
        lblRemainingTime.Caption = "Est. Time Left: " & GetRunTimeString(runTime * (1 - progress) / progress, False)
    DoEvents
End Sub

'Get the time (in milliseconds) since the progress bar "Configure" routine was last called
Public Function GetRunTime() As Long
    GetRunTime = GetTickCount - startTime
End Function

'Get the time (in hours, minutes, seconds) since "Configure" was last called
Public Function GetFormattedRunTime() As String
    GetFormattedRunTime = GetRunTimeString(GetTickCount - startTime)
End Function

'Formats a time in milliseconds as hours, minutes, seconds.milliseconds
'Milliseconds are excluded if showMsecs is set to false
Private Function GetRunTimeString(ByVal runTime As Long, Optional ByVal showMsecs As Boolean = True) As String
    Dim msecs&, hrs&, mins&, secs#
    msecs = runTime
    hrs = Int(msecs / 3600000)
    mins = Int(msecs / 60000) - 60 * hrs
    secs = msecs / 1000 - 60 * (mins + 60 * hrs)
    GetRunTimeString = IIf(hrs > 0, hrs & " hours ", "") _
                     & IIf(mins > 0, mins & " minutes ", "") _
                     & IIf(secs > 0, IIf(showMsecs, secs, Int(secs + 0.5)) & " seconds", "")
End Function

'Returns the current value of the progress bar
Public Function GetValue() As Long
    GetValue = BarVal
End Function

'Returns whether or not the cancel button has been pressed.
'The ProgressDialogue must be polled regularily to detect whether cancel was pressed.
Public Function cancelIsPressed() As Boolean
    cancelIsPressed = Cancelled
End Function

'Recalls that cancel was pressed so that they calling routine can be notified next time it asks.
Private Sub CancelButton_Click()
    Cancelled = True
    lblStatus.Caption = "Cancelled By User. Please Wait."
End Sub
