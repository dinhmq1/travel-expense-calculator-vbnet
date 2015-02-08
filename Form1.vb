'Anh Dinh, Assignment_13, due: 11/06/14
Public Class Form1

    Dim blnOK As Boolean = False
    Const dblPerDayMeals As Double = 37  'reimbursement for meals per day
    Const dblParking As Double = 10      'reimbursement for parking fees per day
    Const dblTaxi As Double = 20         'reimbursement for taxis fees per day
    Const dblLodging As Double = 95      'reimbursement for lodging fees per day
    Const dblPerMile As Double = 0.27    'reimbursement for miles driven

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        Dim dblTotalSpent As Double                 'variable for total amount spent on trip
        Dim dblTotalAirfare As Double               'variable for total airfare cost
        Dim dblTotalMeals As Double                 'variable for total meal cost
        Dim dblTotalCarRental As Double             'variable for total Car rental cost
        Dim dblTotalParking As Double               'variable for total parking cost
        Dim dblTotalTaxi As Double                  'variable for total taxi cost
        Dim dblTotalConferenceSeminar As Double     'variable for total conference seminar registration cost
        Dim dblTotalLodging As Double               'variable for total total lodging cost
        Dim dblDaysSpent As Double                  'variable for total days spent on trip
        Dim dblMilesDriven As Double                'variable for total miles driven
        Dim dblCoveredParking As Double             'variable for total parking fees per day
        Dim dblCoveredTaxi As Double                'variable for total Taxi fees per day
        Dim dblCoveredLodging As Double             'variable for total Lodging fees per day
        Dim dblAmountReimbursedMeals As Double      'variable for total amount reimbursed for meals
        Dim dblAmountReimbursedMileage As Double    'variable for total amount reimbursed for mileage
        Dim dblAmountReimbursedParking As Double    'variable for total amount reimbursed for parking fees
        Dim dblAmountReimbursedTaxi As Double       'variable for total amount reimbursed for Taxi
        Dim dblAmountReimbursedLodging As Double    'variable for total amount reimbursed for Lodging
        Dim dblTotalAmountReimbursed As Double      'variable for total amount reimbursed 
        Dim dblAmountSavedParking As Double         'variable for total savings on parking fees
        Dim dblAmountSavedTaxi As Double            'variable for total savings on taxi fees
        Dim dblAmountSavedLodging As Double         'variable for total savings on lodging fees
        Dim dblTotalAmountSaved As Double           'variable for total savings 


        'begin validation function
        ValidateData(dblDaysSpent, dblTotalAirfare, dblTotalMeals, dblTotalCarRental, dblMilesDriven, dblTotalParking, dblTotalTaxi, dblTotalConferenceSeminar, dblTotalLodging)

        If blnOK = True Then    'if true proceed

            'total amount spent on trip
            dblTotalSpent = CalcTotalSpent(dblTotalAirfare, dblTotalCarRental, dblTotalConferenceSeminar, dblTotalLodging, dblTotalMeals, dblTotalParking, dblTotalTaxi)
            dblCoveredParking = CalcUnallowedParking(dblDaysSpent, dblTotalParking)     'total amount per day on parking
            dblCoveredTaxi = CalcUnallowedTaxi(dblDaysSpent, dblTotalTaxi)              'total amount per day on taxi
            dblCoveredLodging = CalcUnallowedLodging(dblDaysSpent, dblTotalLodging)     'total amount per day on lodging
            dblAmountReimbursedMeals = CalcMeals(dblDaysSpent, dblTotalMeals)           'total amount reimbursed for meals
            dblAmountReimbursedMileage = CalcMileage(dblMilesDriven)                    'total amount reimbursed for mileage
            dblAmountReimbursedParking = CalcParking(dblDaysSpent, dblTotalParking)     'total amount reimbursed for parking
            dblAmountReimbursedTaxi = CalcTaxi(dblDaysSpent, dblTotalTaxi)              'total amount reimbursed for Taxi
            dblAmountReimbursedLodging = CalcLodging(dblDaysSpent, dblTotalLodging)     'total amount reimbursed for lodging
            dblAmountSavedParking = CalcSavedParking(dblDaysSpent, dblTotalParking)     'total amount saved on parking
            dblAmountSavedTaxi = CalcSavedTaxi(dblDaysSpent, dblTotalTaxi)              'total amount saved on taxi
            dblAmountSavedLodging = CalcSavedLodging(dblDaysSpent, dblTotalLodging)     'total amount saved on lodging

            'total amount saved
            dblTotalAmountSaved = dblAmountSavedParking + dblAmountSavedTaxi + dblAmountSavedLodging
            'total amount reimbursed
            dblTotalAmountReimbursed = (dblAmountReimbursedMeals + dblAmountReimbursedMileage + dblAmountReimbursedParking + dblAmountReimbursedTaxi + dblAmountReimbursedLodging)
            'dblTotalAmountReimbursed = dblTotalSpent - dblTotalAmountSaved
            'begin function display
            DisplayExpenses(dblTotalSpent, dblCoveredParking, dblCoveredTaxi, dblCoveredLodging, dblAmountReimbursedMeals, dblAmountReimbursedMileage, dblAmountReimbursedParking,
                            dblAmountReimbursedTaxi, dblAmountReimbursedLodging, dblTotalAmountReimbursed, dblAmountSavedParking, dblAmountSavedTaxi, dblAmountSavedLodging, dblTotalAmountSaved)

        End If

    End Sub


    Private Function CalcTotalSpent(ByVal dblTotalAirfare As Double, ByVal dblTotalMeals As Double, ByVal dblTotalCarRental As Double, ByVal dblTotalParking As Double, ByVal dblTotalTaxi As Double, ByVal dblTotalConferenceSeminar As Double, ByVal dblTotalLodging As Double)

        Dim dblTotalSpent As Double

        dblTotalSpent = dblTotalAirfare + dblTotalCarRental + dblTotalConferenceSeminar + dblTotalLodging + dblTotalMeals + dblTotalParking + dblTotalTaxi

        Return dblTotalSpent
    End Function

    Private Function CalcMeals(ByVal dblDaysSpent As Double, ByVal dblTotalMeals As Double)

        Dim dblAmountReimbursedMeals As Double

        dblAmountReimbursedMeals = dblTotalMeals / dblDaysSpent 'output will be amount spent on meals per day

        If dblAmountReimbursedMeals <= 37 Then               'if <= $37 per day, reimburse full amount
            dblAmountReimbursedMeals = dblTotalMeals
        Else
            dblAmountReimbursedMeals = dblTotalMeals - (dblPerDayMeals * dblDaysSpent)  'if > $37 per day, reimbursement will be deducted from total amount spent on meals
        End If

        Return dblAmountReimbursedMeals
    End Function

    Private Function CalcMileage(ByVal dblMilesDriven As Double)

        Dim dblAmountReimbursedMileage As Double

        dblAmountReimbursedMileage = dblMilesDriven * dblPerMile

        Return dblAmountReimbursedMileage
    End Function

    Private Function CalcParking(ByVal dblDaysSpent As Double, ByVal dblTotalParking As Double)

        Dim dblAmountReimbursedParking As Double

        dblAmountReimbursedParking = dblTotalParking / dblDaysSpent  'output amount spent on parking per day
        If dblAmountReimbursedParking < 10 Then             'if < 10 per day proceed with reimbursement

            dblAmountReimbursedParking = (dblParking * dblDaysSpent) - dblTotalParking 'allowed parking reimbursement multiply by days on trip minus total parking spent outputs amount reimbursed on parking
            txtParkingFees.BackColor = Color.White
            txtParkingFees.ForeColor = Color.Black
        Else
            dblAmountReimbursedParking = 0
        End If

        Return dblAmountReimbursedParking
    End Function

    Private Function CalcTaxi(ByVal dblDaysSpent As Double, ByVal dblTotalTaxi As Double)

        Dim dblAmountReimbursedTaxi As Double

        dblAmountReimbursedTaxi = dblTotalTaxi / dblDaysSpent       'output amount spent on taxi per day

        If dblAmountReimbursedTaxi < 20 Then                        'if < 20 per day proceed with reimbursement

            dblAmountReimbursedTaxi = (dblDaysSpent * dblTaxi) - dblTotalTaxi 'allowed taxi fees reimbursement multiply by days on trip minus total taxi fees spent outputs amount reimbursed on taxi fees        
            txtTaxiCharges.BackColor = Color.White
            txtTaxiCharges.ForeColor = Color.Black

        Else
            dblAmountReimbursedTaxi = 0
        End If

        Return dblAmountReimbursedTaxi
    End Function

    Private Function CalcLodging(ByVal dblDaysSpent As Double, ByVal dblTotalLodging As Double)

        Dim dblAmountReimbursedLodging As Double

        dblAmountReimbursedLodging = dblTotalLodging / dblDaysSpent 'output amount spent on lodging per day

        If dblAmountReimbursedLodging < 95 Then                     'if < 95 per day proceed with reimbursement

            dblAmountReimbursedLodging = (dblDaysSpent * dblLodging) - dblTotalLodging 'allowed lodging fees reimbursement multiply by days on trip minus total lodging fees spent outputs amount reimbursed on lodging fees
            txtLodgingCharges.BackColor = Color.White
            txtLodgingCharges.ForeColor = Color.Black
        Else

            dblAmountReimbursedLodging = 0

        End If

        Return dblAmountReimbursedLodging
    End Function

    Private Function CalcUnallowedParking(ByVal dblDaysSpent As Double, ByVal dblTotalParking As Double)

        Dim dblCoveredParking As Double

        dblCoveredParking = dblTotalParking / dblDaysSpent  'output amount spent on parking per day

        If dblCoveredParking > 10 Then                      'if > 10 per day, no reimbursement allowed
            dblCoveredParking = dblCoveredParking           'amount per day spent on parking
            txtParkingFees.BackColor = Color.Red
            txtParkingFees.ForeColor = Color.Yellow
            MessageBox.Show("Your parking fees exceeded the limit, no reimbursement will be allowed")
        End If

        Return dblCoveredParking

    End Function

    Private Function CalcUnallowedTaxi(ByVal dblDaysSpent As Double, ByVal dblTotalTaxi As Double)

        Dim dblCoveredTaxi As Double

        dblCoveredTaxi = dblTotalTaxi / dblDaysSpent        'output amount spent on taxi fees per day

        If dblCoveredTaxi > 20 Then                         'if > 20 per day, no reimbursement allowed

            dblCoveredTaxi = dblCoveredTaxi                 'amount per day spent on taxi fees

            txtTaxiCharges.BackColor = Color.Red

            txtTaxiCharges.ForeColor = Color.Yellow

            MessageBox.Show("Your taxi fees exceeded the limit, no reimbursement will be allowed")

        End If

        Return dblCoveredTaxi
    End Function

    Private Function CalcUnallowedLodging(ByVal dblDaysSpent As Double, ByVal dblTotalLodging As Double)

        Dim dblCoveredLodging As Double

        dblCoveredLodging = dblTotalLodging / dblDaysSpent          'output amount spent on lodging fees per day

        If dblCoveredLodging > 95 Then                              'if > 95 per day, no reimbursement allowed
            dblCoveredLodging = dblCoveredLodging                   'amount per day spent on lodging fees
            txtLodgingCharges.BackColor = Color.Red
            txtLodgingCharges.ForeColor = Color.Yellow
            MessageBox.Show("Your lodging fees exceeded the limit, no reimbursement will be allowed")

        End If

        Return dblCoveredLodging
    End Function

    Private Function CalcSavedParking(ByVal dblDaysSpent As Double, ByVal dblTotalParking As Double)

        Dim dblAmountSavedParking As Double

        dblAmountSavedParking = dblTotalParking / dblDaysSpent      'output amount spent on parking per day

        If dblAmountSavedParking < 10 Then                          'if < 10 per day, proceed with savings calculation
            dblAmountSavedParking = (dblParking - dblAmountSavedParking) * dblDaysSpent   'amount saved per day
        Else
            dblAmountSavedParking = 0
        End If

        Return dblAmountSavedParking
    End Function

    Private Function CalcSavedTaxi(ByVal dblDaysSpent As Double, ByVal dblTotalTaxi As Double)

        Dim dblAmountSavedTaxi As Double

        dblAmountSavedTaxi = dblTotalTaxi / dblDaysSpent        'output amount spent on taxi fees per day

        If dblAmountSavedTaxi < 20 Then                         ' if < 20 per day, proceed with savings
            dblAmountSavedTaxi = (dblTaxi - dblAmountSavedTaxi) * dblDaysSpent    'amount saved per day
        Else
            dblAmountSavedTaxi = 0
        End If

        Return dblAmountSavedTaxi
    End Function
    Private Function CalcSavedLodging(ByVal dblDaysSpent As Double, ByVal dblTotalLodging As Double)

        Dim dblAmountSavedLodging As Double

        dblAmountSavedLodging = dblTotalLodging / dblDaysSpent          'output amount spent on lodging per day

        If dblAmountSavedLodging < 95 Then                              'if < 95 per day, proceed with savings
            dblAmountSavedLodging = (dblLodging - dblAmountSavedLodging) * dblDaysSpent   'amount saved per day
        Else
            dblAmountSavedLodging = 0
        End If

        Return dblAmountSavedLodging
    End Function

    Private Sub ValidateData(ByRef dblDaysSpent As Double, ByRef dblTotalAirfare As Double, ByRef dblTotalMeals As Double, ByRef dblTotalCarRental As Double, ByRef dblMilesDriven As Double, ByRef dblTotalParking As Double, ByRef dblTotalTaxi As Double, ByRef dblTotalConferenceSeminar As Double, ByRef dblTotalLodging As Double)

        'validate only numerical values
        If IsNumeric(txtNumberofDaysOnTrip.Text) Then

            dblDaysSpent = CDbl(txtNumberofDaysOnTrip.Text)
            txtNumberofDaysOnTrip.BackColor = Color.White
            blnOK = True

        Else
            blnOK = False
            MessageBox.Show("You entered " & txtNumberofDaysOnTrip.Text & ", please only use numerical values.")
            txtNumberofDaysOnTrip.Clear()
            txtNumberofDaysOnTrip.Focus()
            txtNumberofDaysOnTrip.BackColor = Color.Yellow

            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtNumberofDaysOnTrip.Text) < 0 Then

            MessageBox.Show("You entered " & txtNumberofDaysOnTrip.Text & ", please only use positive values.")
            txtNumberofDaysOnTrip.Clear()
            txtNumberofDaysOnTrip.Focus()
            txtNumberofDaysOnTrip.BackColor = Color.Yellow

            Exit Sub
        End If

        'validate only numerical values
        If IsNumeric(txtAirfareCharges.Text) Then

            dblTotalAirfare = txtAirfareCharges.Text
            txtAirfareCharges.BackColor = Color.White
            blnOK = True

        Else
            blnOK = False

            MessageBox.Show("You entered " & txtAirfareCharges.Text & ", please use numerical values only.")
            txtAirfareCharges.Clear()
            txtAirfareCharges.Focus()
            txtAirfareCharges.BackColor = Color.Yellow
            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtAirfareCharges.Text) < 0 Then

            MessageBox.Show("You entered " & txtAirfareCharges.Text & ", please only use positive values.")
            txtAirfareCharges.Clear()
            txtAirfareCharges.Focus()
            txtAirfareCharges.BackColor = Color.Yellow

            Exit Sub
        End If

        'validate only numerical values
        If IsNumeric(txtMealCharges.Text) Then

            dblTotalMeals = CDbl(txtMealCharges.Text)
            txtMealCharges.BackColor = Color.White
            blnOK = True
        Else
            blnOK = False
            MessageBox.Show("You entered " & txtMealCharges.Text & ", please use numerical values only.")
            txtMealCharges.Clear()
            txtMealCharges.Focus()
            txtMealCharges.BackColor = Color.Yellow
            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtMealCharges.Text) < 0 Then

            MessageBox.Show("You entered " & txtMealCharges.Text & ", please only use positive values.")
            txtMealCharges.Clear()
            txtMealCharges.Focus()
            txtMealCharges.BackColor = Color.Yellow

            Exit Sub
        End If

        'validate only numerical values
        If IsNumeric(txtCarRentalCharges.Text) Then

            dblTotalCarRental = CDbl(txtCarRentalCharges.Text)
            txtCarRentalCharges.BackColor = Color.White
            blnOK = True
        Else
            blnOK = False
            MessageBox.Show("You entered " & txtCarRentalCharges.Text & ", please use numerical values only.")
            txtCarRentalCharges.Clear()
            txtCarRentalCharges.Focus()
            txtCarRentalCharges.BackColor = Color.Yellow
            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtCarRentalCharges.Text) < 0 Then

            MessageBox.Show("You entered " & txtCarRentalCharges.Text & ", please only use positive values.")
            txtCarRentalCharges.Clear()
            txtCarRentalCharges.Focus()
            txtCarRentalCharges.BackColor = Color.Yellow

            Exit Sub
        End If

        'validate only numerical values
        If IsNumeric(txtNumberOfMilesDriven.Text) Then

            dblMilesDriven = CDbl(txtNumberOfMilesDriven.Text)
            txtNumberOfMilesDriven.BackColor = Color.White
            blnOK = True
        Else
            blnOK = False
            MessageBox.Show("You entered " & txtNumberOfMilesDriven.Text & ", please use numerical values only.")
            txtNumberOfMilesDriven.Clear()
            txtNumberOfMilesDriven.Focus()
            txtNumberOfMilesDriven.BackColor = Color.Yellow
            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtNumberOfMilesDriven.Text) < 0 Then

            MessageBox.Show("You entered " & txtNumberOfMilesDriven.Text & ", please only use positive values.")
            txtNumberOfMilesDriven.Clear()
            txtNumberOfMilesDriven.Focus()
            txtNumberOfMilesDriven.BackColor = Color.Yellow

            Exit Sub
        End If

        'validate only numerical values
        If IsNumeric(txtParkingFees.Text) Then

            dblTotalParking = CDbl(txtParkingFees.Text)
            txtParkingFees.BackColor = Color.White
            blnOK = True
        Else
            blnOK = False
            MessageBox.Show("You entered " & txtParkingFees.Text & ", please use numerical values only.")
            txtParkingFees.Clear()
            txtParkingFees.Focus()
            txtParkingFees.BackColor = Color.Yellow
            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtParkingFees.Text) < 0 Then

            MessageBox.Show("You entered " & txtParkingFees.Text & ", please only use positive values.")
            txtParkingFees.Clear()
            txtParkingFees.Focus()
            txtParkingFees.BackColor = Color.Yellow

            Exit Sub
        End If

        'validate only numerical values
        If IsNumeric(txtTaxiCharges.Text) Then

            dblTotalTaxi = CDbl(txtTaxiCharges.Text)
            txtTaxiCharges.BackColor = Color.White
            blnOK = True
        Else
            blnOK = False
            MessageBox.Show("You entered " & txtTaxiCharges.Text & ", please use numerical values only.")
            txtTaxiCharges.Clear()
            txtTaxiCharges.Focus()
            txtTaxiCharges.BackColor = Color.Yellow
            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtTaxiCharges.Text) < 0 Then

            MessageBox.Show("You entered " & txtTaxiCharges.Text & ", please only use positive values.")
            txtTaxiCharges.Clear()
            txtTaxiCharges.Focus()
            txtTaxiCharges.BackColor = Color.Yellow

            Exit Sub
        End If


        If IsNumeric(txtConferenceSeminarRegistrationFees.Text) Then

            dblTotalConferenceSeminar = CDbl(txtConferenceSeminarRegistrationFees.Text)
            txtConferenceSeminarRegistrationFees.BackColor = Color.White
            blnOK = True
        Else
            blnOK = False
            MessageBox.Show("You entered " & txtConferenceSeminarRegistrationFees.Text & ", please use numerical values only.")
            txtConferenceSeminarRegistrationFees.Clear()
            txtConferenceSeminarRegistrationFees.Focus()
            txtConferenceSeminarRegistrationFees.BackColor = Color.Yellow
            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtConferenceSeminarRegistrationFees.Text) < 0 Then

            MessageBox.Show("You entered " & txtConferenceSeminarRegistrationFees.Text & ", please only use positive values.")
            txtConferenceSeminarRegistrationFees.Clear()
            txtConferenceSeminarRegistrationFees.Focus()
            txtConferenceSeminarRegistrationFees.BackColor = Color.Yellow

            Exit Sub
        End If

        If IsNumeric(txtLodgingCharges.Text) Then

            dblTotalLodging = CDbl(txtLodgingCharges.Text)
            txtLodgingCharges.BackColor = Color.White
            blnOK = True
        Else
            blnOK = False
            MessageBox.Show("You entered " & txtLodgingCharges.Text & ", please use numerical values only.")
            txtLodgingCharges.Clear()
            txtLodgingCharges.Focus()
            txtLodgingCharges.BackColor = Color.Yellow
            Exit Sub
        End If

        'validate only positive values
        If CDbl(txtLodgingCharges.Text) < 0 Then

            MessageBox.Show("You entered " & txtLodgingCharges.Text & ", please only use positive values.")
            txtLodgingCharges.Clear()
            txtLodgingCharges.Focus()
            txtLodgingCharges.BackColor = Color.Yellow

            Exit Sub
        End If

    End Sub

    Sub DisplayExpenses(ByVal dblTotalSpent As Double, ByVal dblCoveredParking As Double, ByVal dblCoveredTaxi As Double, ByVal dblCoveredLodging As Double, ByVal dblAmountReimbursedMeals As Double, ByVal dblAmountReimbursedMileage As Double,
                        ByVal dblAmountReimbursedParking As Double, dblAmountReimbursedTaxi As Double, ByVal dblAmountReimbursedLodging As Double, ByVal dblTotalAmountReimbursed As Double, ByVal dblAmountSavedParking As Double,
                        ByVal dblAmountSavedTaxi As Double, ByVal dblAmountSavedLodging As Double, ByVal dblTotalAmountSaved As Double)

        lblTotalTravelExpenses.Text = "Total Parking Fees Per Day : " & FormatCurrency(dblCoveredParking.ToString) & vbNewLine _
                                    & "Total Taxi Charges Per Day : " & FormatCurrency(dblCoveredTaxi.ToString) & vbNewLine _
                                    & "Total Lodging Fees Per Day : " & FormatCurrency(dblCoveredLodging.ToString) & vbNewLine _
                                    & "-----------------------------------------------------------------------------" & vbNewLine _
                                    & "Total Reimbursement for Meals : " & FormatCurrency(dblAmountReimbursedMeals.ToString) & vbNewLine _
                                    & "Total Reimbursement for Mileage : " & FormatCurrency(dblAmountReimbursedMileage.ToString) & vbNewLine _
                                    & "Total Reimbursement for Parking : " & FormatCurrency(dblAmountReimbursedParking.ToString) & vbNewLine _
                                    & "Total Reimbursement for Taxi : " & FormatCurrency(dblAmountReimbursedTaxi.ToString) & vbNewLine _
                                    & "Total Reimbursement for Lodging : " & FormatCurrency(dblAmountReimbursedLodging.ToString) & vbNewLine _
                                    & "-----------------------------------------------------------------------------" & vbNewLine _
                                    & "Total Amount Saved on Parking : " & FormatCurrency(dblAmountSavedParking.ToString) & " Per Day" & vbNewLine _
                                    & "Total Amount Saved on Taxi : " & FormatCurrency(dblAmountSavedTaxi.ToString) & " Per Day" & vbNewLine _
                                    & "Total Amount Saved on Lodging : " & FormatCurrency(dblAmountSavedLodging.ToString) & " Per Day" & vbNewLine _
                                    & "-----------------------------------------------------------------------------" & vbNewLine _
                                    & "Total Amount Spent : " & FormatCurrency(dblTotalSpent.ToString) & vbNewLine _
                                    & "Total Amount of Reimbursement Payout : " & FormatCurrency(dblTotalAmountReimbursed.ToString) & vbNewLine _
                                    & "Total Amount Saved : " & FormatCurrency(dblTotalAmountSaved.ToString)


    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Close()

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        txtNumberofDaysOnTrip.Clear()
        txtNumberofDaysOnTrip.BackColor = Color.White
        txtNumberofDaysOnTrip.ForeColor = Color.Black
        txtNumberofDaysOnTrip.Text = 0
        txtAirfareCharges.Clear()
        txtAirfareCharges.BackColor = Color.White
        txtAirfareCharges.ForeColor = Color.Black
        txtAirfareCharges.Text = 0
        txtMealCharges.Clear()
        txtMealCharges.BackColor = Color.White
        txtMealCharges.ForeColor = Color.Black
        txtMealCharges.Text = 0
        txtCarRentalCharges.Clear()
        txtCarRentalCharges.BackColor = Color.White
        txtCarRentalCharges.ForeColor = Color.Black
        txtCarRentalCharges.Text = 0
        txtNumberOfMilesDriven.Clear()
        txtNumberOfMilesDriven.BackColor = Color.White
        txtNumberOfMilesDriven.ForeColor = Color.Black
        txtNumberOfMilesDriven.Text = 0
        txtParkingFees.Clear()
        txtParkingFees.BackColor = Color.White
        txtParkingFees.ForeColor = Color.Black
        txtParkingFees.Text = 0
        txtTaxiCharges.Clear()
        txtTaxiCharges.BackColor = Color.White
        txtTaxiCharges.ForeColor = Color.Black
        txtTaxiCharges.Text = 0
        txtConferenceSeminarRegistrationFees.Clear()
        txtConferenceSeminarRegistrationFees.Text = 0
        txtLodgingCharges.Clear()
        txtLodgingCharges.BackColor = Color.White
        txtLodgingCharges.ForeColor = Color.Black
        txtLodgingCharges.Text = 0
        lblTotalTravelExpenses.Text = String.Empty

    End Sub
End Class
