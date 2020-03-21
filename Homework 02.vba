Option Explicit

'2.1
Public Sub Character_Validation01()
  Dim a() As String
  Dim i As Integer
  Dim name_len As Integer
  name_initials = InputBox("Please enter your initials:", , "Robert")
  name_len = Len(name_initials)
  ReDim a(0 To name_len)
  i = 1
  Do Until i > name_len
    a(i) = Mid(name_initials, i, 1)
    If Asc(a(i)) >= 97 And Asc(a(i)) <= 122 Then
      a(i) = UCase(a(i))
      Debug.Print a(i)
    End If
    i = i + 1
  Loop
End Sub

Public Sub Character_Validation02()
  Dim a() As String
  Dim i As Integer
  Dim name_len As Integer
  name_initials = InputBox("Please enter your initials:", , "Robert")
  name_len = Len(name_initials)
  ReDim a(0 To name_len)
  i = 1
  Do While i <= name_len
    a(i) = Mid(name_initials, i, 1)
    If Asc(a(i)) >= 97 And Asc(a(i)) <= 122 Then
      a(i) = UCase(a(i))
      Debug.Print a(i)
    End If
    i = i + 1
  Loop
End Sub

Public Sub Loop_Discounted_Value()
  Dim ttm As Integer
  Dim fv As Double
  Dim dr As Single
  Dim dv As Double
  Do Until ttm > 0
    ttm = InputBox("Please enter the time t:", , 10)
  Loop
  Do Until fv > 0
    fv = InputBox("Please enter the future value:", , 100)
  Loop
  Do Until dr > 0 And dr < 1
    dr = InputBox("Please enter the discounted rate:", , 0.05)
  Loop
  dv = fv / (1 + dr) ^ ttm
  str1 = "time to maturity is " & ttm & "years" & vbNewLine
  str2 = "future value is " & Format(fv, "currency") & vbNewLine
  str3 = "discounted rate is " & Format(dr, "percent") & vbNewLine
  str4 = "the discounted value is " & Format(dv, "currency") & vbNewLine
  MsgBox str1 & str2 & str3 & str4
End Sub

Public Sub white_space_check()
  Dim i As Integer
  Dim value_len As Integer
  Dim a() As String
  input_value = InputBox("Please enter characters:", , "X")
  value_len = Len(input_value)
  ReDim a(0 To value_len)
  i = 1
  Do Until i = value_len
  a(i) = Mid(input_value, i, 1)
  Select Case a(i)
  Case " ", "\n", "\t"
  MsgBox "You entered a white space character!"
  Exit Do
  End Select
  i = i + 1
  Loop
  If i = value_len Then
  MsgBox "The characters in lower case are " & UCase(input_value)
  End If
End Sub
