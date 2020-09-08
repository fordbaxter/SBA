Public Function get_SBA_7a_Terms(pprice As Double, reprice As Double, sellerfi As Double, workingcap As Double, percentdown As Double, term As String)
' percentdown = down payment expressed as a decimal (.15 typical) to calculate loan value
' Seller financing is not considered part of the equity UNLESS it is on standby the full life of the loan and does not exceed half the total equity.
' SBA defines project cost as "all costs required to complete the change of ownership, regardless of the source of funds"
' SOPs require a minimum of 10% down for a complete change of ownership. For manager buyouts different rules exist.

Dim percentguaranteed As Double: percentguaranteed = 0.75
Dim projectcost As Double: projectcost = 1
Dim downpayment As Double
Dim loanamount As Double
Dim closingcosts As Double


' compute circular loan terms
Dim i: i = 0
While i < 10

    projectcost = pprice + reprice + workingcap + closingcosts
    downpayment = projectcost * percentdown
    loanamount = projectcost - sellerfi - downpayment
    closingcosts = getClosingCosts(loanamount, percentguaranteed, reprice)
   
i = i + 1
Wend

' throw error if loan value is over 7a limit
If loanamount > 5000000 Then
    term = "overmax"
End If

' return requested value
Select Case term
Case Is = "closingcosts"
    get_SBA_7a_Terms = closingcosts
    
Case Is = "projectcost"
    get_SBA_7a_Terms = projectcost

Case Is = "downpayment"
    get_SBA_7a_Terms = downpayment

Case Is = "loanamount"
    get_SBA_7a_Terms = loanamount

Case Is = "loanterm"
    'get_SBA_7a_Terms = getLoanTerm

Case Is = "overmax"
    get_SBA_7a_Terms = 0
    MsgBox ("Loan too large")

Case Else
    get_SBA_7a_Terms = "Error"

End Select
End Function


Public Function getClosingCosts(loanamount As Double, percentguaranteed As Double, reprice As Double)
' gets the closing costs for an SBA 7a deal
' closing costs vary but these are general figures used for an estimate

Dim closingcosts As Double: closingcosts = getGFee(loanamount, percentguaranteed)

If reprice > 0 Then
    closingcosts = closingcosts + 5000 'cost of RE related items paid by buyer
End If

closingcosts = closingcosts + 16500 ' FIB estimated closing costs
getClosingCosts = closingcosts

End Function


Public Function getGFee(loanamount As Double, percentguaranteed As Double)
' returns just the SBA guarantee fee based on the loan amount
' tested against the SBA SOPs on 9/4/20 to work
Dim gfee As Double: gfee = 0
guaranteedportion = loanamount * percentguaranteed

Select Case loanamount
Case Is > 5000000
   ' Loans greater than $5,000,000
   Debug.Print ("Exceeds max loan value, return $0")
   gfee = 0
   
Case Is > 1000000
    ' Loans greater than $1,000,000 but under $5,000,000
    guaranteedportion = guaranteedportion - 1000000
    gfee = 35000 + (guaranteedportion * 0.0375)

Case Is > 700000
    ' Loans greater than $700,000 but less than $1,000,000
    gfee = guaranteedportion * 0.035
        
Case Is > 0
    ' Loans less than $700,000
    gfee = guaranteedportion * 0.02

Case Else
    MsgBox ("Unknown error calculating SBA guarantee fee.")
    
End Select

getGFee = gfee
End Function
