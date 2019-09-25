Module initializeForm
    Public Sub initForm()
        'currentYear = Year(Now)
        'Call resetSelectTypeFrameToRed
        'Call resetCalcFrameToNormal
        'Call modelList
        msgbox("bam")
        Exit Sub
        '*********************
        'nissanFOM = "Nathan.Smith3@nissan-usa.com" 'no direct emails to FOM at this time
        'nissanFOM = "smull@tonygroup.com"
        gwForm.dlrCodeTB.Text = "98009"
        '*********************

        With gwForm
            '.StartUpPosition = 3
            '.Height = 300
            '.Width = 390

            'initialize dropdowns

            With gwForm.outOfWarrCB 'init drop down
                .Items.Add("Time")
                '    .Clear
                '    .AddItem "Time"
                '.AddItem "Miles"
                '.AddItem "Time and Miles"
                '.AddItem "Force Goodwill"
            End With
            With gwForm.goodwillCB 'init drop down
                '    .Clear
                '    .AddItem " 0% DI and/or CP"
                '.AddItem "10% DI and/or CP"
                '.AddItem "20% DI and/or CP"
                '.AddItem "30% DI and/or CP"
                '.AddItem "40% DI and/or CP"
            End With

            'reset option buttons and textboxes
            .repairOnlyOB.Checked = False
            .rentalOnlyOB.Checked = False
            .repairAndRentalOB.Checked = False

            .dlrCodeTB.BackColor = vbWhite

            With .ROnumberTB
                .Text = ""
                .BackColor = vbWhite
            End With

            .lineNumberTB.Text = ""

            With .openDateTB
                .Text = "mm/dd/yyyy"
                .ForeColor = RGB(211, 211, 211)
                .BackColor = vbWhite
            End With
            With .currentMileageTB
                .Text = ""
                .BackColor = vbWhite
            End With
            With .failedPartNumberTB
                .Text = ""
                .BackColor = vbWhite
            End With
            With .vinTB
                .Text = ""
                .BackColor = vbWhite
            End With
            With .modelCB
                .Text = ""
                .BackColor = vbWhite
            End With
            With .inServDateTB
                .Text = "mm/dd/yyyy"
                .ForeColor = RGB(211, 211, 211)
                .BackColor = vbWhite
            End With

            .origOwnerChkBx.Value = False
            .multiNissanChkBx.Value = False

            With .brandedNoOB
                .Value = False
                .ForeColor = vbNormal
            End With
            With .brandedYesOB
                .Value = False
                .ForeColor = vbNormal
            End With
            With .detailsTB
                .Text = ""
                .BackColor = vbWhite
            End With
            With .dsaYesOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .dsaNoOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .comebackNoOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .comebackYesOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .photosNoOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .photosYesOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .preauthNoOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .preauthYesOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .techlineNoOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .techlineYesOB
                .Checked = False
                .ForeColor = vbNormal
            End With
            With .rentalTB
                .Text = ""
                .BackColor = vbWhite
            End With
            With .rentalOutTB
                .Text = "mm/dd/yyyy"
                .ForeColor = RGB(211, 211, 211)
                .BackColor = vbWhite
            End With
            With .partsOrderedTB
                .Text = "mm/dd/yyyy"
                .ForeColor = RGB(211, 211, 211)
                .BackColor = vbWhite
            End With
            With .partsArrivedTB
                .Text = "mm/dd/yyyy"
                .ForeColor = RGB(211, 211, 211)
                .BackColor = vbWhite
            End With
            With .repairsCompletedTB
                .Text = "mm/dd/yyyy"
                .ForeColor = RGB(211, 211, 211)
                .BackColor = vbWhite
            End With
            With .rentalReturnedTB
                .Text = "mm/dd/yyyy"
                .ForeColor = RGB(211, 211, 211)
                .BackColor = vbWhite
            End With

            .totalDaysTB.Text = ""
            .totalRentalAmtTB.Text = ""

            With .requestedDaysTB
                .Text = ""
                .BackColor = vbWhite
            End With
            With .requestedAmtTB
                .Text = ""
                .BackColor = vbWhite
            End With

            'setup calculator frame
            With .fgPartsTB
                .Visible = True
                .Text = ""
            End With
            With .fgLaborTB
                .Visible = True
                .Text = ""
            End With
            .fgExpensesTB.Text = ""
            .fgTotalCostTB.Text = ""
            .fgPercentTB.Text = "0"
            With .cpdiPartsTB
                .Visible = True
                .Text = ""
            End With
            With .cpdiLaborTB
                .Visible = True
                .Text = ""
            End With
            .cpdiExpensesTB.Text = ""
            .cpdiTotalCostTB.Text = ""
            .cpdiPercentTB.Text = "0"
            With .partsTotalTB
                .Visible = True
                .Text = ""
            End With
            With .laborTotalTB
                .Visible = True
                .Text = ""
            End With
            .expensesTotalTB.Text = ""
            .totalCostTB.Text = ""
            .totalPercentTB.Text = "0"

            .questionsFrame.Visible = False
            .rentalFrame.Visible = False
            .calcFrame.Visible = False

            .resetBtn.Visible = False
            .emailBtn.Visible = False


        End With

        boolOutlook = False

        'gwForm.modelYearSB.Max = 2100 'sets max date
        'gwForm.modelYearTB.Value = currentYear 'set default date

        'reset calculator fields
        fgPartAmt = 0
        fgLaborAmt = 0
        fgExpAmt = 0
        fgTotalAmt = 0
        cpPartAmt = 0
        cpLaborAmt = 0
        cpExpAmt = 0
        cpTotalAmt = 0
        diPartAmt = 0
        diLaborAmt = 0
        diExpAmt = 0
        diTotalAmt = 0
        fgPercent = 0
        cpPercent = 0
        diPercent = 0
        percentTotal = 0


        'gwForm.repairAndRentalOB.Value = True
        'gwForm.ROnumberTB.Value = "123456"
        'gwForm.lineNumberTB.Value = "B"
        'gwForm.openDateTB = "8/10/2019"
        'gwForm.currentMileageTB.Value = "12987"
        'gwForm.failedPartNumberTB.Value = "12345-ABC56 STARTER MOTOR ASSY"
        'gwForm.vinTB.Value = "3N1AB7AP9FY301872"
        'gwForm.modelCB.Value = "Altima"
        'gwForm.inServDateTB.Value = "1/1/2010"
        'gwForm.origOwnerChkBx.Value = True
        'gwForm.multiNissanChkBx.Value = False
        'gwForm.brandedNoOB = True
        'gwForm.outOfWarrCB.Value = "Force Goodwill"
        'gwForm.goodwillCB.Value = "40% DI and/or CP"
        'gwForm.detailsTB.Value = "Some text goes here to explain this request for goodwill coverage Some text goes here to explain this request for goodwill coverage.Some text goes here to explain this request for goodwill coverageSome text goes here to explain this request for goodwill coverageSome text goes here to explain this request for goodwill coverage"
        'gwForm.dsaYesOB.Value = True
        'gwForm.comebackNoOB.Value = True
        'gwForm.photosNoOB.Value = True
        'gwForm.preauthNoOB.Value = True
        'gwForm.techlineYesOB.Value = True
        'gwForm.rentalTB.Value = "Some text that explains the rental or other expenses goes here.Some text goes here to explain this request for goodwill coverageSome text goes here to explain this request for goodwill coverageSome text goes here to explain this request for goodwill coverageSome text goes here to explain this request for goodwill coverageSome text goes here to explain this request for goodwill coverage"
        'gwForm.rentalOutTB.Value = "08/01/2019"
        'gwForm.partsOrderedTB.Value = "08/02/2019"
        'gwForm.partsArrivedTB.Value = "08/08/2019"
        'gwForm.repairsCompletedTB.Value = "08/09/2019"
        'gwForm.rentalReturnedTB.Value = "08/10/2019"
        'gwForm.totalDaysTB.Value = ""
        'gwForm.totalRentalAmtTB.Value = "150"
        'gwForm.requestedDaysTB.Value = "7"
        'gwForm.requestedAmtTB.Value = "150"
        'gwForm.fgPartsTB.Value = "5000"
        'gwForm.fgLaborTB.Value = "6000"
        'gwForm.fgExpensesTB.Value = ""
        'gwForm.cpPartsTB.Value = "7000"
        'gwForm.cpLaborTB.Value = "8000"
        'gwForm.cpExpensesTB.Value = "1000"
        'gwForm.diPartsTB.Value = "1100"
        'gwForm.diLaborTB.Value = "1200"
        'gwForm.diExpensesTB.Value = "1000"

    End Sub


End Module
