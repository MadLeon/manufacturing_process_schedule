' =========================================================
' Callback_BatchInputButton: Triggered when the "Multi-Input Mode" toggle
' button is clicked on the custom Ribbon tab.
' =========================================================
' Ribbon button click callback
' ModuleRibbonHandlers

Option Explicit

Public Sub BatchInputButton_Click(control As IRibbonControl)
    ' Call your batch input launcher
    Call LaunchBatchInput
End Sub

Public Sub RandomNumberButton_Click(control As IRibbonControl)
    ' Call your RNG code here
    Call GenerateRandomNumber
End Sub

Public Sub RemoveLeadingZeroButton_Click(control As IRibbonControl)
    ' Call your remove leading zero code here
    Call RemoveLeadingZero
End Sub

Sub OpenGDandTForm(control As IRibbonControl)
    ' Call Input Panel
    Form_GDandTInput.Show
End Sub




