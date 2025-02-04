colour project in vb 6.0
Option Explicit

Private Sub Form_Load()
    ' Set initial values for the sliders
    HScrollRed.Value = 0
    HScrollGreen.Value = 0
    HScrollBlue.Value = 0
    
    ' Set the default color for the result box
    lblResult.BackColor = RGB(0, 0, 0)
End Sub

Private Sub HScrollRed_Change()
    ' Update the resulting color based on the sliders' positions
    UpdateColor
End Sub

Private Sub HScrollGreen_Change()
    ' Update the resulting color based on the sliders' positions
    UpdateColor
End Sub

Private Sub HScrollBlue_Change()
    ' Update the resulting color based on the sliders' positions
    UpdateColor
End Sub

Private Sub UpdateColor()
    ' Combine the values of the sliders to create the resulting color
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    
    Red = HScrollRed.Value
    Green = HScrollGreen.Value
    Blue = HScrollBlue.Value
    
    ' Update the label or shape's background color with the mixed color
    lblResult.BackColor = RGB(Red, Green, Blue)
    
    ' Optionally, display the RGB values in a label
    lblColorValues.Caption = "RGB: " & Red & ", " & Green & ", " & Blue
End Sub




thank You
