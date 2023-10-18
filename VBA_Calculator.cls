Private Sub cmdCalculate_Click()
'Declare the variables
    Dim Total As Double
    Dim Class1Allo As Double
    Dim Class2Allo As Double
    Dim Class3Allo As Double
    Dim BankAllo As Double
    Dim investmentDuration As Integer
    Dim InterestRateYear1 As Double
    Dim InterestRateYear2 As Double
    Dim InterestRateYear3 As Double
    Dim Banks As Double
    Dim AIandIT As Double
    Dim Consumer As Double
    Dim Energy As Double
    Dim GSC As Double
    Dim Automotive As Double
    Dim AC1_En As Double
    Dim AC1_Co As Double
    Dim AC1_AI As Double
    Dim AC1_TotalAllo As Double
    Dim AC2_Ba As Double
    Dim AC2_AI As Double
    Dim AC2_En As Double
    Dim AC2_TotalAllo As Double
    Dim AC3_Au As Double
    Dim AC3_GSC As Double
    Dim AC3_En As Double
    Dim AC3_TotalAllo As Double
    Dim expectedReturnClass1 As Double
    Dim expectedReturnClass2 As Double
    Dim expectedReturnClass3 As Double
    Dim expectedReturnBank As Double
    Dim expectedTotal As Double
    Dim expectedTotalProfit As Double
    Dim expectedGrowth As Double
    Dim expectedAC1Profit As Double
    Dim expectedAC2Profit As Double
    Dim expectedAC3Profit As Double
    Dim expectedBankProfit As Double
    Dim expectedAC1Growth As Double
    Dim expectedAC2Growth As Double
    Dim expectedAC3Growth As Double
    Dim expectedBankGrowth As Double
    Dim allocationControls_SubSector As Variant
    Dim i As Integer
    Dim assetClasses(1 To 3) As String
    Dim totalAllo_Sub(1 To 3) As Double
    Dim j As Integer
    Dim Allocations() As Variant
    Dim AllocationControls() As Variant
    Dim k As Integer
    Dim totalAllo As Double
    Dim rateControls As Variant
    Dim rateRanges As Variant
    Dim m As Integer
    Dim textBoxes As Variant
    Dim interestRates As Variant
    Dim n As Integer

'Initializing variables from text box values
    Class1Allo = CDbl(Me.txtClass1Allo.Value)
    Class2Allo = CDbl(Me.txtClass2Allo.Value)
    Class3Allo = CDbl(Me.txtClass3Allo.Value)
    BankAllo = CDbl(Me.txtBankAllo.Value)
    Total = CDbl(Me.txtTotal.Value)
    investmentDuration = CInt(Me.txtDuration.Value)
    Banks = CDbl(Me.txtBanks.Value)
    AIandIT = CDbl(Me.txtAIandIT.Value)
    Consumer = CDbl(Me.txtConsumer.Value)
    Energy = CDbl(Me.txtEnergy.Value)
    GSC = CDbl(Me.txtGSC.Value)
    Automotive = CDbl(Me.txtAutomotive.Value)
    AC1_En = CDbl(Me.txtAC1_En.Value)
    AC1_Co = CDbl(Me.txtAC1_Co.Value)
    AC1_AI = CDbl(Me.txtAC1_AI.Value)
    AC2_Ba = CDbl(Me.txtAC2_Ba.Value)
    AC2_AI = CDbl(Me.txtAC2_AI.Value)
    AC2_En = CDbl(Me.txtAC2_En.Value)
    AC3_Au = CDbl(Me.txtAC3_Au.Value)
    AC3_GSC = CDbl(Me.txtAC3_GSC.Value)
    AC3_En = CDbl(Me.txtAC3_En.Value)
    InterestRateYear1 = CDbl(Me.txtYr1.Value)
    InterestRateYear2 = CDbl(Me.txtYr2.Value)
    InterestRateYear3 = CDbl(Me.txtYr3.Value)

'Check for appropriate sub-sector allocation for each asset class
    allocationControls_SubSector = Array(txtAC1_En, txtAC1_Co, txtAC1_AI, txtAC2_Ba, txtAC2_AI, txtAC2_En, txtAC3_Au, txtAC3_GSC, txtAC3_En)
    
    For i = 1 To UBound(allocationControls_SubSector) + 1
        Dim currentAllocation_SubSector As Double
        currentAllocation_SubSector = Val(allocationControls_SubSector(i - 1).Value)
        
        If IsNumeric(currentAllocation_SubSector) Then
            If currentAllocation_SubSector >= 0 Then
                allocationControls_SubSector(i - 1).Value = Format(currentAllocation_SubSector, "0")
            Else
                MsgBox "Error, please ensure allocation is non-negative"
                Exit Sub
            End If
        Else
            MsgBox "Error, please enter a valid numeric allocation."
            Exit Sub
        End If
    Next i

'Check to ensure appropriate sub-sector allocation for each asset class
    assetClasses(1) = "Asset Class 1"
    assetClasses(2) = "Asset Class 2"
    assetClasses(3) = "Asset Class 3"
    
    totalAllo_Sub(1) = AC1_En + AC1_Co + AC1_AI
    totalAllo_Sub(2) = AC2_Ba + AC2_AI + AC2_En
    totalAllo_Sub(3) = AC3_Au + AC3_GSC + AC3_En
    
    For j = 1 To 3
        If totalAllo_Sub(j) > 100 Then
            MsgBox "Error, total allocation of sub-sectors for " & assetClasses(j) & " cannot exceed 100%"
            Exit Sub
        End If
    Next j
'Check for appropriate Asset and Bank class allocation
    Allocations = Array(Class1Allo, Class2Allo, Class3Allo, BankAllo)
    AllocationControls = Array("txtClass1Allo", "txtClass2Allo", "txtClass3Allo", "txtBankAllo")
    
    For k = LBound(Allocations) To UBound(Allocations)
        If IsNumeric(Allocations(k)) And Allocations(k) >= 0 Then
            Me.Controls(AllocationControls(k)).Value = Format(Me.Controls(AllocationControls(k)).Value, "0")
        Else
            MsgBox "Error, please ensure " & Me.Controls(AllocationControls(k)).Tag & " allocation is numeric and non-negative"
            Exit Sub
        End If
    Next k

'Check to ensure appropriate allocation of Asset and Bank classes
    totalAllo = 0
    For k = LBound(Allocations) To UBound(Allocations)
        totalAllo = totalAllo + Allocations(k)
    Next k

    If totalAllo > 100 Then
        MsgBox "Error, total allocation of asset and bank classes cannot exceed 100%"
        Exit Sub
    End If
    
'Check for appropriate rate values
    rateControls = Array(txtBanks, txtAIandIT, txtConsumer, txtEnergy, txtGSC, txtAutomotive)
    rateRanges = Array("-1 to 8", "8 to 15", "-5 to 3", "0 to 8", "-5 to 5", "-10 to 5")
    
    For m = 1 To UBound(rateControls) + 1
        Dim currentRate As Double
        currentRate = Val(rateControls(m - 1).Value)
        
        If IsNumeric(currentRate) Then
            If currentRate >= Split(rateRanges(m - 1), " to ")(0) And currentRate <= Split(rateRanges(m - 1), " to ")(1) Then
                rateControls(m - 1).Value = Format(currentRate, "0")
            Else
                MsgBox "Error, Please ensure your rate figure is between " & rateRanges(m - 1) & "%"
                Exit Sub
            End If
        Else
            MsgBox "Error, Please enter a valid numeric rate."
            Exit Sub
        End If
    Next m
'Check for investment duration
    If IsNumeric(investmentDuration) And investmentDuration >= 1 And investmentDuration <= 3 Then
         Me.txtDuration.Value = Format(Me.txtDuration.Value, "0")
    Else
        MsgBox "Error, invalid number of years. Please ensure your investment duration is between 1 and 3 years!"
        Exit Sub
    End If

'Assign bank rate values
    textBoxes = Array(Me.txtYr1, Me.txtYr2, Me.txtYr3)
    interestRates = Array(InterestRateYear1, InterestRateYear2, InterestRateYear3)
    
    For n = LBound(textBoxes) To UBound(textBoxes)
        If IsNumeric(interestRates(n)) Then
            textBoxes(n).Value = Format(interestRates(n), "0.00")
        Else
            MsgBox "Error, please ensure rate value is numeric"
            Exit Sub
        End If
    Next n
'Calculating expected values
    expectedReturnClass1 = ((((Class1Allo / 100) * Total) * (AC1_En / 100)) * (1 + (Energy / 100)) ^ investmentDuration) + ((((Class1Allo / 100) * Total) * (AC1_Co / 100)) * (1 + (Consumer / 100)) ^ investmentDuration) + ((((Class1Allo / 100) * Total) * (AC1_AI / 100)) * (1 + (AIandIT / 100)) ^ investmentDuration)
    expectedReturnClass2 = ((((Class2Allo / 100) * Total) * (AC2_Ba / 100)) * (1 + (Banks / 100)) ^ investmentDuration) + ((((Class2Allo / 100) * Total) * (AC2_AI / 100)) * (1 + (AIandIT / 100)) ^ investmentDuration) + ((((Class2Allo / 100) * Total) * (AC2_En / 100)) * (1 + (Energy / 100)) ^ investmentDuration)
    expectedReturnClass3 = ((((Class3Allo / 100) * Total) * (AC3_Au / 100)) * (1 + (Automotive / 100)) ^ investmentDuration) + ((((Class3Allo / 100) * Total) * (AC3_GSC / 100)) * (1 + (GSC / 100)) ^ investmentDuration) + ((((Class3Allo / 100) * Total) * (AC3_En / 100)) * (1 + (Energy / 100)) ^ investmentDuration)
    
' Calculate expectedReturnBank based on investmentDuration with variable rates
    If investmentDuration = 1 Then
        expectedReturnBank = (((BankAllo / 100) * Total) * (1 + (InterestRateYear1 / 100)))
    ElseIf investmentDuration = 2 Then
        expectedReturnBank = (((BankAllo / 100) * Total) * (1 + (InterestRateYear1 / 100)) * (1 + (InterestRateYear2 / 100)))
    ElseIf investmentDuration = 3 Then
        expectedReturnBank = (((BankAllo / 100) * Total) * (1 + (InterestRateYear1 / 100)) * (1 + (InterestRateYear2 / 100)) * (1 + (InterestRateYear3 / 100)))
    End If
        
    expectedTotal = expectedReturnClass1 + expectedReturnClass2 + expectedReturnClass3 + expectedReturnBank
    expectedTotalProfit = expectedTotal - Total
    expectedGrowth = (expectedTotalProfit / Total) * 100
    expectedAC1Profit = expectedReturnClass1 - ((Class1Allo / 100) * Total)
    expectedAC2Profit = expectedReturnClass2 - ((Class2Allo / 100) * Total)
    expectedAC3Profit = expectedReturnClass3 - ((Class3Allo / 100) * Total)
    expectedBankProfit = expectedReturnBank - ((BankAllo / 100) * Total)
    expectedAC1Growth = (expectedAC1Profit / Total) * 100
    expectedAC2Growth = (expectedAC2Profit / Total) * 100
    expectedAC3Growth = (expectedAC3Profit / Total) * 100
    expectedBankGrowth = (expectedBankProfit / Total) * 100
    
    
    Me.txtAC1Return.Text = Format(expectedReturnClass1, "£0,000,000,000")
    Me.txtAC2Return.Text = Format(expectedReturnClass2, "£0,000,000,000")
    Me.txtAC3Return.Text = Format(expectedReturnClass3, "£0,000,000,000")
    Me.txtBankReturn.Text = Format(expectedReturnBank, "£0,000,000,000")
    Me.txtTotalReturn.Text = Format(expectedTotal, "£0,000,000,000")
    Me.txtTotalProfit.Text = Format(expectedTotalProfit, "£0,000,000,000")
    Me.txtAC1Profit.Text = Format(expectedAC1Profit, "£0,000,000,000")
    Me.txtAC2Profit.Text = Format(expectedAC2Profit, "£0,000,000,000")
    Me.txtAC3Profit.Text = Format(expectedAC3Profit, "£0,000,000,000")
    Me.txtBankProfit.Text = Format(expectedBankProfit, "£0,000,000,000")
    Me.txtTotalGrowth.Text = Format(expectedGrowth, "0.00\%")
    Me.txtAC1Growth.Text = Format(expectedAC1Growth, "0.00\%")
    Me.txtAC2Growth.Text = Format(expectedAC2Growth, "0.00\%")
    Me.txtAC3Growth.Text = Format(expectedAC3Growth, "0.00\%")
    Me.txtBankGrowth.Text = Format(expectedBankGrowth, "0.00\%")
End Sub
Private Sub cmdClose1_Click()
    Dim response As VbMsgBoxResult

' Ask the user for confirmation
    response = MsgBox("Are you sure you want to close the calculator?", vbYesNo + vbQuestion, "Confirmation")

    If response = vbYes Then
' User confirmed, closes the calculator
        Unload Me
    End If
End Sub

Private Sub cmdClose2_Click()
    Dim response As VbMsgBoxResult

' Ask the user for confirmation
    response = MsgBox("Are you sure you want to close the calculator?", vbYesNo + vbQuestion, "Confirmation")

    If response = vbYes Then
' User confirmed, closes the calculator
        Unload Me
    End If
End Sub

Private Sub cmdClose3_Click()
    Dim response As VbMsgBoxResult

' Ask the user for confirmation
    response = MsgBox("Are you sure you want to close the calculator?", vbYesNo + vbQuestion, "Confirmation")

    If response = vbYes Then
' User confirmed, closes the calculator
        Unload Me
    End If
End Sub

Private Sub cmdClose4_Click()
    Dim response As VbMsgBoxResult

' Ask the user for confirmation
    response = MsgBox("Are you sure you want to close the calculator?", vbYesNo + vbQuestion, "Confirmation")

    If response = vbYes Then
' User confirmed, closes the calculator
        Unload Me
    End If
End Sub

Private Sub cmdReset1_Click()
    Dim ctrl As Control
    Dim response As VbMsgBoxResult

' Ask the user for confirmation
    response = MsgBox("Are you sure you want to reset all entries?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
' Loop through all controls on the UserForm
        For Each ctrl In frmCalculator.Controls
            If TypeName(ctrl) = "TextBox" Then
                ctrl.Value = ""
            End If
        Next ctrl
    End If
End Sub

Private Sub cmdReset2_Click()
    Dim ctrl As Control
    Dim response As VbMsgBoxResult

' Ask the user for confirmation
    response = MsgBox("Are you sure you want to reset all entries?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
' Loop through all controls on the UserForm
        For Each ctrl In frmCalculator.Controls
            If TypeName(ctrl) = "TextBox" Then
                ctrl.Value = ""
            End If
        Next ctrl
    End If
End Sub

Private Sub cmdReset3_Click()
    Dim ctrl As Control
    Dim response As VbMsgBoxResult

' Ask the user for confirmation
    response = MsgBox("Are you sure you want to reset all entries?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
' Loop through all controls on the UserForm
        For Each ctrl In frmCalculator.Controls
            If TypeName(ctrl) = "TextBox" Then
                ctrl.Value = ""
            End If
        Next ctrl
    End If
End Sub

Private Sub cmdReset4_Click()
    Dim ctrl As Control
    Dim response As VbMsgBoxResult

' Ask the user for confirmation
    response = MsgBox("Are you sure you want to reset all entries?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
' Loop through all controls on the UserForm
        For Each ctrl In frmCalculator.Controls
            If TypeName(ctrl) = "TextBox" Then
                ctrl.Value = ""
            End If
        Next ctrl
    End If
End Sub
