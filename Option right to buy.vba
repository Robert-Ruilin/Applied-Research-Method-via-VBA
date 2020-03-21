Option Explicit

Public Sub option_right_to_buy()
  Dim t As Integer
  Dim sp() As Double
  Dim op() As Double
  Dim disop() As Double
  Dim signal() As Variant
  Dim rtb() As Variant
  Dim K As Double
  Dim i, j As Integer
  Dim max_payoff, time_max_payoff As Double
  Dim initial_price As Double
  Dim discounted_rate As Single
  Dim elapsed_time As Double
  
  Do
  initial_price = InputBox("Please enter the initial price of the stock at time t = 1:", , 100)
  Loop While initial_price < 80 Or initial_price > 120
  
  Do
  discounted_rate = InputBox("Please enter the annual discounted rate at time t = 1:", , 0.02)
  Loop While discounted_rate < 0 Or discounted_rate > 0.2
  
  Do
  t = InputBox("Please enter the simulation time period t (daily):", , 100)
  Loop While t < 1
  
  elapsed_time = Timer '开始时间
  
  ReDim sp(1 To t)
  ReDim op(1 To t)
  ReDim disop(1 To t)
  ReDim signal(1 To t)
  ReDim rtb(1 To t)
  
  K = 100
  i = 1
  Cells(1, 1) = "Time step"
  Cells(i + 1, 1) = i
  sp(i) = initial_price
  Cells(1, 2) = "Stock price"
  Cells(2, 2) = sp(i)
  Cells(1, 3) = "Option payoff"
  op(i) = Application.WorksheetFunction.Max(sp(i) - K, 0)
  Cells(2, 3) = op(i)
  Cells(1, 4) = "Market Signal"
  Cells(1, 5) = "Right to buy"
  If op(i) > 0 Then
    rtb(i) = "The option should be exercised."
  Else
    rtb(i) = "The option shoudl not be exercised."
  End If
  Cells(2, 5) = rtb(i)
  Cells(1, 6) = "Discounted option payoff"
  disop(i) = op(i)
  Cells(2, 6) = disop(i)
  
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
  Cells(i + 2, 1) = i + 1
  Cells(i + 2, 2) = sp(i + 1)
  Cells(i + 2, 3) = op(i + 1)
  Cells(i + 2, 4) = signal(i + 1)
  Cells(i + 2, 5) = rtb(i + 1)
  disop(i + 1) = op(i + 1) / (1 + discounted_rate / 252) ^ i
  Cells(i + 2, 6) = disop(i + 1)
  If disop(i + 1) = Application.WorksheetFunction.Max(disop) Then
    max_payoff = disop(i + 1)
    time_max_payoff = i + 1
  End If
  i = i + 1
  Loop Until i = t
  Range("A:F").Columns.AutoFit
  
  MsgBox "Given by the strike price K = 100" & vbNewLine & "The investor should exercise this call option at time t = " & time_max_payoff & ", the maximum option payoff is " & max_payoff

  elapsed_time = Timer - elapsed_time '结束时间
  Debug.Print elapsed_time
End Sub

