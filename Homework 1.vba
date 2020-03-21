Option Explicit

'1.1
Public Sub Msgbox1()
  MsgBox "Hello World!"
End Sub

Public Sub Msgbox2()
  MsgBox "H e l l o   W o r l d  !"
End Sub

Public Sub MsgBox3()
  str1 = "hello,"
  str2 = "world"
  str3 = "wor"
  str4 = " "
  
  Line1 = ":" & str1 & str4 & str2 & ":" & vbNewLine
  Line2 = ":" & str1 & str4 & str2 & String(2, 2) & ":" & vbNewLine
  Line3 = ":" & str1 & str4 & str3 & ":" & vbNewLine
  Line4 = ":" & str1 & String(2, 2) & str2 & String(2, 2) & ":" & vbNewLine
  Line5 = ":" & str1 & String(3, 2) & str2 & str4 & ":" & vbNewLine
  Line6 = ":" & str1 & str4 & str2 & String(2, 2) & ":" & vbNewLine
  Line7 = ":" & String(3, 2) & str1 & str4 & str3 & ":" & vbNewLine
  Line8 = ":" & str1 & str4 & str3 & String(4, 2) & ":" & vbNewLine
  MsgBox Line1 & Line2 & Line3 & Line4 & Line5 & Line6 & Line7 & Line8
End Sub

Public Sub MsgBox4()
  str1 = "Bond yield:"
  str2 = " "
  str3 = "_"
  
  Line1 = ":" & str1 & str2 & "0.045" & str3 & str3 & ":" & vbNewLine
  Line2 = ":" & String(2, 2) & str1 & str2 & "0.045" & ":" & vbNewLine
  Line3 = ":" & str2 & str1 & str2 & "0.04" & String(2, 2) & ":" & vbNewLine
  Line4 = ":" & str1 & str2 & "4.5%" & str3 & str2 & str3 & ":" & vbNewLine
  Line5 = ":" & str1 & str2 & "4.50" & str3 & str2 & ":" & vbNewLine
  MsgBox Line1 & Line2 & Line3 & Line4 & Line5
End Sub

Public Sub FV_calculation()
  Dim pv As Double
  Dim t As Integer
  Dim ir As Single
  Dim fv As Double
  
  pv = 100#
  t = 1
  ir = 0.0245
  fv = pv * (1 + ir) ^ t
  
  Cells(1, 1).Value = "Present Value:"
  Cells(1, 2).Value = Format(pv, "Currency")
  Cells(2, 1).Value = "Time Horizon:"
  Cells(2, 2).Value = t
  Cells(3, 1).Value = "annual interest rate:"
  Cells(3, 2).Value = Format(ir, "Percent")
  Cells(4, 1).Value = "Future value:"
  Cells(4, 2).Value = Format(fv, "Currency")
  Cells(3, 1).Columns.AutoFit
  Cells(1, 2).Columns.AutoFit
End Sub

'1.2
Public Sub PV_calculation01()
  Dim pv As Double
  Dim fv As Double
  Dim t As Integer
  
  Dim discounted_rate As Single
  fv = InputBox("Please enter future value:", , 100)
  discounted_rate = InputBox("Please enter discounted rate:", , 0.05)
  t = InputBox("Please enter time period:", , 1)
  pv = fv / (1 + discounted_rate) ^ t
  
  str1 = "Future value is " & Format(fv, "Currency") & vbNewLine
  str2 = "Time period is " & t & vbNewLine
  str3 = "Discounted rate is " & Format(discounted_rate, "Percent") & vbNewLine
  str4 = "The discounted value is " & Format(pv, "Currency") & vbNewLine
  MsgBox str1 & str2 & str3 & str4
End Sub

Public Sub PV_calculation02()
  Dim pv As Double
  Dim fv As Double
  Dim t As Integer
  Dim frequency As Integer
  Dim discounted_rate As Single
  
  fv = InputBox("Please enter future value:", , 100)
  discounted_rate = InputBox("Please enter discounted rate:", , 0.05)
  t = InputBox("Please enter time period:", , 1)
  frequency = InputBox("Please enter frequency of time period:", , 1)
  pv = fv / (1 + discounted_rate / frequency) ^ (t * frequency)
  
  str1 = "Future value is " & Format(fv, "Currency") & vbNewLine
  str2 = "Time period is " & t & vbNewLine
  str3 = "Frequency of time period is " & frequency & vbNewLine
  str4 = "Discounted rate is " & Format(discounted_rate, "Percent") & vbNewLine
  str5 = "The discounted value is " & Format(pv, "Currency") & vbNewLine
  MsgBox str1 & str2 & str3 & str4 & str5
End Sub

Private Sub main() '清楚单元格内容、格式、及批注
  Cells.ClearContents
  Cells.ClearFormats
  Cells.ClearComments
End Sub

Public Sub PE_calculation()
  Dim PE As Double
  Dim D As Double
  Dim RR As Single
  
  D = 0.6
  RR = 0.1
  PE = D / RR
  
  Cells(1, 1).Value = "dividend:"
  Cells(1, 2).Value = Format(D, "Currency")
  Cells(2, 1).Value = "rate of return:"
  Cells(2, 2).Value = Format(RR, "Percent")
  Cells(3, 1).Value = "Price of equity share:"
  Cells(3, 2).Value = Format(PE, "Currency")
  Cells(3, 1).Columns.AutoFit
  Cells(2, 2).Columns.AutoFit
End Sub

'1.3
Public Sub Sqr_Calculation()
  Dim RM As Double
  Dim Sqr1 As Double
  Dim Sqr2 As Double
  RM = InputBox("Please enter a real number:", , 100)
  Sqr1 = RM ^ 0.5
  Sqr2 = VBA.Sqr(RM)
  MsgBox "The square root of " & RM & " is " & Format(Sqr2, "0.00")
End Sub

Public Sub Power_Calculation()
  Dim RM As Double
  Dim N As Integer
  Dim Power As Double
  RM = InputBox("Please enter a real number:", , 100)
  N = InputBox("Please enter an integer:", , 2)
  Power = Application.WorksheetFunction.Power(RM, N)
  MsgBox "The power of " & RM & " is " & Format(Power, "0.00")
End Sub

Public Sub Random_Generation()
  Dim rn As Double
  rn = VBA.Rnd()
  MsgBox "The random number is " & Format(rn, "0.00")
End Sub

Public Sub UpperCase_Initials()
  name = InputBox("Please enter your initials:", , "apple")
  term = UCase(name)
  MsgBox "The initials are " & term
End Sub

Public Sub LowerCase_Name()
  name = InputBox("Please enter your name:", , "HE RUILIN")
  term = LCase(name)
  MsgBox "The initials are " & term
End Sub

'1.4
Public Sub PV_retire_calculation01()
  Dim ttr As Integer
  Dim pmt As Double
  Dim ir As Single
  Dim pv As Double
  
  pmt = InputBox("Please input payment per year:", , 30000)
  ttr = InputBox("Please input time to retire:", , 15)
  ir = InputBox("Please input annual interest rate:", , 0.1)
  pv = pmt * (1 / ir - 1 / (ir * (1 + ir) ^ ttr))
  str1 = "The payment per year is " & Format(pmt, "currency") & vbNewLine
  str2 = "Time to retire is " & ttr & " years" & vbNewLine
  str3 = "The annual interest rate is " & Format(ir, "percent") & vbNewLine
  str4 = "The amount is " & Format(pv, "currency") & vbNewLine
  MsgBox str1 & str2 & str3 & str4
End Sub

Public Sub PV_retire_calculation02()
  Dim ttr As Integer
  Dim pmt As Double
  Dim ir As Single
  Dim pv As Double
  Dim i As Integer
  pmt = InputBox("Please input payment per year:", , 30000)
  ttr = InputBox("Please input time to retire:", , 15)
  ir = InputBox("Please input annual interest rate:", , 0.1)
  pv = 0
  For i = 1 To ttr
    pv = pv + 1 / (1 + ir) ^ i
  Next
  pv = pv * pmt
  str1 = "The payment per year is " & Format(pmt, "currency") & vbNewLine
  str2 = "Time to retire is " & ttr & " years" & vbNewLine
  str3 = "The annual interest rate is " & Format(ir, "percent") & vbNewLine
  str4 = "The amount is " & Format(pv, "currency") & vbNewLine
  MsgBox str1 & str2 & str3 & str4
End Sub
