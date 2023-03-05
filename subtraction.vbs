'Option Explicit
'Dim firstNumber
'Assign first number
firstNumber=InputBox("Please enter a number to subtract:","First number")
firstNumber=Cint(firstNumber)
'Dim secondNumber
'Assign second number
secondNumber=Cint(InputBox("Please enter the second number to subtract"& vbCrLf&"to the sum","Second number",0))
'Store sum to the third variable
'Dim sum
sum=FirstNumber-SecondNumber
'Display sum
MsgBox "The sum is " &  sum
