Dim ii, sOperator, strExpr, y 
strExpr = inputbox("Calculation", "Calculator", "1+1")

' insert spaces around all operators
For Each sOperator in Array("+","-","*","/","%")
  strExpr = Trim( Replace( strExpr, sOperator, Space(1) & sOperator & Space(1)))
Next
' replace all multi spaces with a single space 
Do While Instr( strExpr, Space(2))
  strExpr = Trim( Replace( strExpr, Space(2), Space(1)))
Loop
