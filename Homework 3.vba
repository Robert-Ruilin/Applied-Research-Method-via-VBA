Option Explicit

Public Function maximum_payoff(t As Integer, K As Double, initial_price As Double, discounted_rate As Single) As Double
  Dim sp() As Double
  Dim op() As Double
  Dim disop() As Double
  Dim signal() As Variant
  Dim rtb() As Variant
  Dim i, j As Integer
  Dim max_payoff, time_max_payoff As Double
  ReDim sp(1 To t)
  ReDim op(1 To t)
  ReDim disop(1 To t)
  ReDim signal(1 To t)
  ReDim rtb(1 To t)
  
  i = 1
  sp(i) = initial_price
  op(i) = Application.WorksheetFunction.Max(sp(i) - K, 0)
  If op(i) > 0 Then
    rtb(i) = "The option should be exercised."
  Else
    rtb(i) = "The option shoudl not be exercised."
  End If
  disop(i) = op(i)
  
  Do
  j = Rnd()
  If j > 0.5 Then
    signal(i + 1) = "positive"
    sp(i + 1) = sp(i) + Rnd() * 2
  Else
    signal(i + 1) = "negative"
    sp(i + 1) = sp(i) - Rnd() * 2
  End If
  op(i + 1) = Application.WorksheetFunction.Max(sp(i + 1) - K, 0)
  If op(i + 1) > 0 Then
    rtb(i + 1) = "The option should be exercised."
  Else
    rtb(i + 1) = "The option shoudl not be exercised."
  End If
  disop(i + 1) = op(i + 1) / (1 + discounted_rate / 252) ^ i
  If disop(i + 1) = Application.WorksheetFunction.Max(disop) Then
    max_payoff = disop(i + 1)
    time_max_payoff = i + 1
  End If
  i = i + 1
  Loop Until i = t
  Range("A:F").Columns.AutoFit
  
  maximum_payoff = max_payoff
End Function
