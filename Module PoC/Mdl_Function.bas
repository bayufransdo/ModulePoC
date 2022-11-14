Attribute VB_Name = "Mdl_Function"
Option Explicit
'This to Close form
Sub CloseForm(CloseMode As Integer, Cancel As Integer)
    If CloseMode <> 1 Then
        Dim askQuestion As Integer
        Cancel = 1
        askQuestion = MsgBox("Are you want to exit?", vbQuestion + vbYesNo, "Point of Change")
        If askQuestion = vbYes Then
            Cancel = 0
        End If
    End If
End Sub

Sub FrmReportCCValue()
    'To make FrmReport Change Content combo box  list
    If FrmReport.ComboCC.value = "PLAN" Then
        If FrmReport.ComboProcess.value = "SPIRALLING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!A4:A5"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!D4:D9"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!G4:G5"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!J4:J6"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!M4:M5"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "WELDING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!Q4:Q5"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!T4:T9"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!W4:W5"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!Z4:Z6"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AC4:AC5"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FINISHING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!A22:A23"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!D22:D25"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!G22:G23"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!J22:J24"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!M22:M23"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "APPEARANCE" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!Q22:Q23"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!T22"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!W22:W23"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!Z22:Z23"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AC22:AC23"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FORMING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!A37:A38"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!D37:D42"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!G37:G38"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!J37:J39"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!M37:M38"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    ElseIf FrmReport.ComboCC.value = "UNPLAN" Then
        If FrmReport.ComboProcess.value = "SPIRALLING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!B4:B8"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!E4:E11"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!H4"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!K4:K5"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!N4"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "WELDING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!R4:R7"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!U4:U16"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!X4"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AA4"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AD4"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FINISHING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!B22:B26"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!E22:E31"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!H22"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!K22:K26"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!N22:N25"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "APPEARANCE" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!R22:R26"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!U22:U25"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!X22"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AA22:AA25"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AD22:AD25"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FORMING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!B37:B41"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!E37:E43"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!E37"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!K37:K40"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!N37:N40"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    ElseIf FrmReport.ComboCC.value = "OTHERS" Then
        If FrmReport.ComboProcess.value = "SPIRALLING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!C4:C5"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!F4:F5"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "WELDING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!S4:S5"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!V4:V8"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FINISHING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "APPEARANCE" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!S22:S23"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FORMING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!C37:C38"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    Else
    End If
End Sub

Sub frmsearchclickCCValue()
    'To make FrmSearchClick Change Content combo box  list
    If FrmSearchClick.ComboCC.value = "PLAN" Then
        FrmSearchClick.ComboChangeContent.Visible = True
        FrmSearchClick.TxtChangeContent.Visible = True
        If FrmSearchClick.ComboProcess.value = "SPIRALLING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!A4:A5"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!D4:D9"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!G4:G5"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!J4:J6"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!M4:M5"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "WELDING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!Q4:Q5"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!T4:T9"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!W4:W5"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!Z4:Z6"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AC4:AC5"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "FINISHING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!A22:A23"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!D22:D25"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!G22:G23"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!J22:J24"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!M22:M23"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "APPEARANCE" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!Q22:Q23"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!T22"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!W22:W23"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!Z22:Z23"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AC22:AC23"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "FORMING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!A37:A38"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!D37:D42"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!G37:G38"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!J37:J39"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!M37:M38"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    ElseIf FrmSearchClick.ComboCC.value = "UNPLAN" Then
        FrmSearchClick.ComboChangeContent.Visible = True
        FrmSearchClick.TxtChangeContent.Visible = True
        If FrmSearchClick.ComboProcess.value = "SPIRALLING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!B4:B8"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!E4:E11"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!H4"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!K4:K5"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!N4"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "WELDING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!R4:R7"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!U4:U16"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!X4"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AA4"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AD4"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "FINISHING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!B22:B26"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!E22:E31"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!H22"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!K22:K26"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!N22:N25"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "APPEARANCE" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!R22:R26"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!U22:U25"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!X22"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AA22:AA25"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AD22:AD25"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "FORMING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!B37:B41"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!E37:E43"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!H37"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!K37:K40"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!N37:N40"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    ElseIf FrmSearchClick.ComboCC.value = "OTHERS" Then
        FrmSearchClick.ComboChangeContent.Visible = True
        FrmSearchClick.TxtChangeContent.Visible = True
        If FrmSearchClick.ComboProcess.value = "SPIRALLING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!C4:C5"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!F4:F5"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "WELDING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!S4:S5"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!V4:V8"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "FINISHING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "APPEARANCE" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!S22:S23"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmSearchClick.ComboProcess.value = "FORMING" Then
            If FrmSearchClick.ComboPOC.value = "MAN" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!C37:C38"
            ElseIf FrmSearchClick.ComboPOC.value = "MACHINE" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "METHOD" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MEASUREMENT" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmSearchClick.ComboPOC.value = "MATERIAL" Then
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmSearchClick.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    Else
        FrmSearchClick.ComboChangeContent.Visible = False
        FrmSearchClick.TxtChangeContent.Visible = False
    End If
End Sub

Sub CloseTreatment()
        FrmCheckItem.CBContinue.value = False
        FrmCheckItem.CBScrap.value = False
        FrmCheckItem.CBHold.value = False
        FrmCheckItem.CBSorting.value = False
        FrmCheckItem.CBContinue.Enabled = False
        FrmCheckItem.CBScrap.Enabled = False
        FrmCheckItem.CBHold.Enabled = False
        FrmCheckItem.CBSorting.Enabled = False
        
    'SCRAP / REJECT FORM
        FrmCheckItem.Label3.Enabled = False
        FrmCheckItem.Label4.Enabled = False
        FrmCheckItem.Label5.Enabled = False
        FrmCheckItem.TextBox1.Enabled = False
        FrmCheckItem.TextBox2.Enabled = False
        FrmCheckItem.TextBox3.Enabled = False
        FrmCheckItem.CmdSubmit1.Enabled = False

    'SORTING FORM
        FrmCheckItem.Label6.Enabled = False
        FrmCheckItem.Label7.Enabled = False
        FrmCheckItem.Label8.Enabled = False
        FrmCheckItem.Label9.Enabled = False
        FrmCheckItem.Label10.Enabled = False
        FrmCheckItem.Label11.Enabled = False
        FrmCheckItem.TextBox4.Enabled = False
        FrmCheckItem.TextBox5.Enabled = False
        FrmCheckItem.txtN.Enabled = False
        FrmCheckItem.txtR.Enabled = False
        FrmCheckItem.txtHasil.Enabled = False
        FrmCheckItem.CmdSubmit2.Enabled = False
        
    'ON HOLD FORM
        FrmCheckItem.Label12.Enabled = False
        FrmCheckItem.Label13.Enabled = False
        FrmCheckItem.Label14.Enabled = False
        FrmCheckItem.TextBox9.Enabled = False
        FrmCheckItem.TextBox10.Enabled = False
        FrmCheckItem.TextBox11.Enabled = False
        FrmCheckItem.CmdSubmit3.Enabled = False

    'CONTINUE FORM
        FrmCheckItem.Label15.Enabled = False
        FrmCheckItem.TextBox12.Enabled = False
        FrmCheckItem.CmdSubmit4.Enabled = False
        FrmCheckItem.TextBox12.value = ""
        
        'To make value off checkbox are cleared when user changed checkbox
    'SCRAP / REJECT FORM
        FrmCheckItem.TextBox1.value = ""
        FrmCheckItem.TextBox2.value = ""
        FrmCheckItem.TextBox3.value = ""

    'SORTING FORM
        FrmCheckItem.TextBox4.value = ""
        FrmCheckItem.TextBox5.value = ""
        FrmCheckItem.txtN.value = ""
        FrmCheckItem.txtR.value = ""
        FrmCheckItem.txtHasil.value = ""
        
    'ON HOLD FORM
        FrmCheckItem.TextBox9.value = ""
        FrmCheckItem.TextBox10.value = ""
        FrmCheckItem.TextBox11.value = ""
End Sub

Sub OpenTreatment()
        FrmCheckItem.CBContinue.value = False
        FrmCheckItem.CBScrap.value = False
        FrmCheckItem.CBHold.value = False
        FrmCheckItem.CBSorting.value = False
        FrmCheckItem.CBContinue.Enabled = True
        FrmCheckItem.CBScrap.Enabled = True
        FrmCheckItem.CBHold.Enabled = True
        FrmCheckItem.CBSorting.Enabled = True
        
    'SCRAP / REJECT FORM
        FrmCheckItem.Label3.Enabled = True
        FrmCheckItem.Label4.Enabled = True
        FrmCheckItem.Label5.Enabled = True
        FrmCheckItem.TextBox1.Enabled = True
        FrmCheckItem.TextBox2.Enabled = True
        FrmCheckItem.TextBox3.Enabled = True
        FrmCheckItem.CmdSubmit1.Enabled = True

    'SORTING FORM
        FrmCheckItem.Label6.Enabled = True
        FrmCheckItem.Label7.Enabled = True
        FrmCheckItem.Label8.Enabled = True
        FrmCheckItem.Label9.Enabled = True
        FrmCheckItem.Label10.Enabled = True
        FrmCheckItem.Label11.Enabled = True
        FrmCheckItem.TextBox4.Enabled = True
        FrmCheckItem.TextBox5.Enabled = True
        FrmCheckItem.txtN.Enabled = True
        FrmCheckItem.txtR.Enabled = True
        FrmCheckItem.txtHasil.Enabled = True
        FrmCheckItem.CmdSubmit2.Enabled = True
        
    'ON HOLD FORM
        FrmCheckItem.Label12.Enabled = True
        FrmCheckItem.Label13.Enabled = True
        FrmCheckItem.Label14.Enabled = True
        FrmCheckItem.TextBox9.Enabled = True
        FrmCheckItem.TextBox10.Enabled = True
        FrmCheckItem.TextBox11.Enabled = True
        FrmCheckItem.CmdSubmit3.Enabled = True

    'CONTINUE FORM
        FrmCheckItem.Label15.Enabled = True
        FrmCheckItem.TextBox12.Enabled = True
        FrmCheckItem.CmdSubmit4.Enabled = True
        
        'To make value off checkbox are cleared when user changed checkbox
    'SCRAP / REJECT FORM
        FrmCheckItem.TextBox1.value = ""
        FrmCheckItem.TextBox2.value = ""
        FrmCheckItem.TextBox3.value = ""

    'SORTING FORM
        FrmCheckItem.TextBox4.value = ""
        FrmCheckItem.TextBox5.value = ""
        FrmCheckItem.txtN.value = ""
        FrmCheckItem.txtR.value = ""
        FrmCheckItem.txtHasil.value = ""
        
    'ON HOLD FORM
        FrmCheckItem.TextBox9.value = ""
        FrmCheckItem.TextBox10.value = ""
        FrmCheckItem.TextBox11.value = ""
End Sub

Sub ClearFrmCheckItem()
    FrmCheckItem.TextBox1.value = ""
    FrmCheckItem.TextBox2.value = ""
    FrmCheckItem.TextBox3.value = ""
    FrmCheckItem.TextBox4.value = ""
    FrmCheckItem.TextBox5.value = ""
    FrmCheckItem.TextBox9.value = ""
    FrmCheckItem.TextBox10.value = ""
    FrmCheckItem.TextBox11.value = ""
    FrmCheckItem.TextBox12.value = ""
    FrmCheckItem.TextBox13.value = ""
    FrmCheckItem.TextBox14.value = ""
    FrmCheckItem.TextBox15.value = ""
    FrmCheckItem.TextBox16.value = ""
    FrmCheckItem.TextBox17.value = ""
    FrmCheckItem.TextBox18.value = ""
    FrmCheckItem.TextBox19.value = ""
    FrmCheckItem.TextBox20.value = ""
    FrmCheckItem.TextBox21.value = ""
    FrmCheckItem.TextBox22.value = ""
    FrmCheckItem.TextBox23.value = ""
    FrmCheckItem.TextBox24.value = ""
    FrmCheckItem.TextBox25.value = ""
    FrmCheckItem.txtN.value = ""
    FrmCheckItem.txtR.value = ""
    FrmCheckItem.txtHasil.value = ""
    
    FrmCheckItem.Label41.Enabled = False
    FrmCheckItem.TextBox25.Enabled = False
    
    FrmCheckItem.CheckBox1.value = False
    FrmCheckItem.CheckBox2.value = False
    FrmCheckItem.CheckBox3.value = False
    FrmCheckItem.CheckBox4.value = False
    
    FrmCheckItem.TextBox13.SetFocus
End Sub

Sub FrmStartActive()
    'To make userform is autofit with any screen
    With Application
        .WindowState = xlMaximized
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
    
    'To make content in center of app
    With FrmStart.Frame1
        .Top = (Application.Height / 100) * 12.5
        .Left = (Application.Width / 100) * 15
    End With
    
    With FrmStart.Frame2
        .Left = (FrmStart.Width - FrmStart.Width) + 5
        .Top = FrmStart.Height - 60
    End With
    'To make clock realtime
    Do While ThisWorkbook.Sheets("DataLogin").Range("K47") = "" And ThisWorkbook.Sheets("DataLogin").Range("K49") = ""
        FrmStart.LblDate.Caption = Format(Now, "dd mmmm yyyy")
        FrmStart.LblTime.Caption = Format(Now, "hh:mm:ss")
        DoEvents
    Loop
End Sub

Sub FrmLoginButton()
    'To make login according to radio button value
    If FrmLogin.Opt_Leader.value = True Then
        ThisWorkbook.Sheets("DataLogin").Range("K47", "K49").ClearContents
        Call LeaderLogin
    ElseIf FrmLogin.Opt_Member.value = True Then
        ThisWorkbook.Sheets("DataLogin").Range("K47", "K49").ClearContents
        Call MemberLogin
    ElseIf FrmLogin.Opt_Member.value = False And FrmLogin.Opt_Leader.value = False And TxtName.value = "" Then
        MsgBox "Please input your EE correctly!", vbCritical, "Point of Change"
    ElseIf FrmLogin.Opt_Member.value = False And FrmLogin.Opt_Leader.value = False Then
        MsgBox "please choose the radio button!", vbCritical, "Point of Change"
    End If
End Sub

Sub FrmLoginTextboxChange()
    Dim ws As Worksheet
    Dim Rng As Range
    Set ws = ThisWorkbook.Sheets("DataLogin")
    Set Rng = ws.Range("F:F").Find(TxtEE.value, , , xlWhole)
    FrmLogin.TxtEE.PasswordChar = "*"
    
    If IsNumeric(FrmLogin.TxtEE.value) Then
        FrmLogin.LblReq.Caption = ""
        If Rng Is Nothing Then
            FrmLogin.TxtName.value = ""
        Else
            FrmLogin.TxtName.value = Rng.Offset(0, -1)
        End If
    Else
        FrmLogin.TxtEE.value = ""
        FrmLogin.TxtName.value = ""
        FrmLogin.LblReq.Caption = "Fill it With Number!"
    End If
End Sub

Sub FrmReportComboSizeChange()
    'To make size of content based on user want
    Dim Zsize As Integer
    If FrmReport.ComboSize.value = "" Then
        MsgBox "Please Input size Correctly!", vbCritical, "Point of Change - Error"
    Else
        Zsize = FrmReport.ComboSize.value
        ThisWorkbook.Sheets("DataLogin").Range("K44").value = FrmReport.ComboSize.value
        FrmReport.Frame11.Zoom = Zsize
    
        With FrmReport.Frame10
            If FrmReport.ComboSize.value = 100 Then
                .Top = -1
                .Left = -1
                If Application.Height < 750 Then
                    Frame11.Left = (Application.Width / 6) - 20
                End If
            ElseIf FrmReport.ComboSize.value = 90 Then
                .Top = 30
                .Left = 60
                If Application.Height < 750 Then
                    Frame11.Left = Application.Width / 6
                End If
                
            ElseIf FrmReport.ComboSize.value = 80 Then
                .Top = 70
                .Left = 110
                If Application.Height < 750 Then
                    Frame11.Left = (Application.Width / 6) + 10
                End If
            ElseIf FrmReport.ComboSize.value = 70 Then
                .Top = 110
                .Left = 160
                If Application.Height < 750 Then
                    Frame11.Left = (Application.Width / 6) + 20
                End If
            End If
        End With

    End If
End Sub

Sub FrmSearchActive()
    With Application
        .WindowState = xlMaximized
        Me.Left = .Left
        Me.Width = .Width
    End With
    
    With FrmSearch.Framewawa
        .Left = (FrmSearch.Width - FrmSearch.Width) + 5
        .Top = FrmSearch.Height - 60
    End With
    
    With FrmSearch.Frame1
        .Width = (Application.Width * 97) / 100
    End With
    
    With FrmSearch.ListBox1
        .Width = (Application.Width * 95.5) / 100
    End With
    FrmSearch.Label1.Width = FrmSearch.Width
    
    flag = True
    With FrmSearch.lblToday
        Do While flag = True
            FrmSearch.lblToday.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
            DoEvents
        Loop
    End With
End Sub

Sub FrmCheckingItemCmd1()
    If CBScrap.value = False Then
        MsgBox "Please Check the CheckBox!", vbExclamation, "Point of Change"
    ElseIf CBScrap.value = True Then
        Dim ws As Worksheet
        Dim IDfier As Long
        
        Set ws = ThisWorkbook.Sheets("DataReport")
        IDfier = ws.Range("A:A").Find(Label46.Caption, , , xlWhole).Row
        
        ws.Cells(IDfier, 20).value = "Opened"
        ws.Cells(IDfier, 21).value = TextBox1.value
        ws.Cells(IDfier, 22).value = TextBox2.value
        ws.Cells(IDfier, 23).value = TextBox3.value
        ws.Cells(IDfier, 24).value = ""
        ws.Cells(IDfier, 25).value = ""
        ws.Cells(IDfier, 26).value = ""
        ws.Cells(IDfier, 27).value = ""
        ws.Cells(IDfier, 28).value = ""
        ws.Cells(IDfier, 29).value = ""
        ws.Cells(IDfier, 30).value = ""
        ws.Cells(IDfier, 31).value = ""
        ws.Cells(IDfier, 32).value = ""
    
        ws.Cells(IDfier, 33).value = TextBox13.value
        ws.Cells(IDfier, 34).value = TextBox14.value
        ws.Cells(IDfier, 35).value = TextBox15.value
        ws.Cells(IDfier, 36).value = TextBox17.value
        ws.Cells(IDfier, 37).value = TextBox18.value
        ws.Cells(IDfier, 38).value = TextBox19.value
        ws.Cells(IDfier, 39).value = TextBox16.value
        ws.Cells(IDfier, 40).value = TextBox20.value
        ws.Cells(IDfier, 41).value = TextBox21.value
        ws.Cells(IDfier, 42).value = TextBox22.value
        ws.Cells(IDfier, 43).value = TextBox23.value
        ws.Cells(IDfier, 44).value = TextBox24.value
        If CheckBox1.value = True Then
            ws.Cells(IDfier, 45).value = "OK"
        ElseIf CheckBox2.value = True Then
            ws.Cells(IDfier, 45).value = "NG"
        End If
        ws.Cells(IDfier, 46).value = TextBox25.value
        Call ClearFrmCheckItem
        FrmCheckItem.Hide
    End If
End Sub

Sub FrmCheckingItemCmd2()
    If CBSorting.value = False Then
        MsgBox "Please Check the CheckBox!", vbExclamation, "Point of Change"
    ElseIf CBSorting.value = True Then
        Dim ws As Worksheet
        Dim IDfier As Long
        
        Set ws = ThisWorkbook.Sheets("DataReport")
        IDfier = ws.Range("A:A").Find(Label46.Caption, , , xlWhole).Row
        
        ws.Cells(IDfier, 20).value = "Opened"
        ws.Cells(IDfier, 21).value = ""
        ws.Cells(IDfier, 22).value = ""
        ws.Cells(IDfier, 23).value = ""
        ws.Cells(IDfier, 24).value = TextBox5.value
        ws.Cells(IDfier, 25).value = TextBox4.value
        ws.Cells(IDfier, 26).value = txtN.value
        ws.Cells(IDfier, 27).value = txtR.value
        ws.Cells(IDfier, 28).value = txtHasil.value
        ws.Cells(IDfier, 29).value = ""
        ws.Cells(IDfier, 30).value = ""
        ws.Cells(IDfier, 31).value = ""
        ws.Cells(IDfier, 32).value = ""
    
        ws.Cells(IDfier, 33).value = TextBox13.value
        ws.Cells(IDfier, 34).value = TextBox14.value
        ws.Cells(IDfier, 35).value = TextBox15.value
        ws.Cells(IDfier, 36).value = TextBox17.value
        ws.Cells(IDfier, 37).value = TextBox18.value
        ws.Cells(IDfier, 38).value = TextBox19.value
        ws.Cells(IDfier, 39).value = TextBox16.value
        ws.Cells(IDfier, 40).value = TextBox20.value
        ws.Cells(IDfier, 41).value = TextBox21.value
        ws.Cells(IDfier, 42).value = TextBox22.value
        ws.Cells(IDfier, 43).value = TextBox23.value
        ws.Cells(IDfier, 44).value = TextBox24.value
        If CheckBox1.value = True Then
            ws.Cells(IDfier, 45).value = "OK"
        ElseIf CheckBox2.value = True Then
            ws.Cells(IDfier, 45).value = "NG"
        End If
        ws.Cells(IDfier, 46).value = TextBox25.value
        Call ClearFrmCheckItem
        FrmCheckItem.Hide
    End If
End Sub
Sub FrmCheckingItemCmd3()
    If CBHold.value = False Then
        MsgBox "Please Check the CheckBox!", vbExclamation, "Point of Change"
    ElseIf CBHold.value = True Then
        Dim ws As Worksheet
        Dim IDfier As Long
        
        Set ws = ThisWorkbook.Sheets("DataReport")
        IDfier = ws.Range("A:A").Find(Label46.Caption, , , xlWhole).Row
        
        ws.Cells(IDfier, 20).value = "Opened"
        ws.Cells(IDfier, 21).value = ""
        ws.Cells(IDfier, 22).value = ""
        ws.Cells(IDfier, 23).value = ""
        ws.Cells(IDfier, 24).value = ""
        ws.Cells(IDfier, 25).value = ""
        ws.Cells(IDfier, 26).value = ""
        ws.Cells(IDfier, 27).value = ""
        ws.Cells(IDfier, 28).value = ""
        ws.Cells(IDfier, 29).value = TextBox11.value
        ws.Cells(IDfier, 30).value = TextBox10.value
        ws.Cells(IDfier, 31).value = TextBox9.value
        ws.Cells(IDfier, 32).value = ""
    
        ws.Cells(IDfier, 33).value = TextBox13.value
        ws.Cells(IDfier, 34).value = TextBox14.value
        ws.Cells(IDfier, 35).value = TextBox15.value
        ws.Cells(IDfier, 36).value = TextBox17.value
        ws.Cells(IDfier, 37).value = TextBox18.value
        ws.Cells(IDfier, 38).value = TextBox19.value
        ws.Cells(IDfier, 39).value = TextBox16.value
        ws.Cells(IDfier, 40).value = TextBox20.value
        ws.Cells(IDfier, 41).value = TextBox21.value
        ws.Cells(IDfier, 42).value = TextBox22.value
        ws.Cells(IDfier, 43).value = TextBox23.value
        ws.Cells(IDfier, 44).value = TextBox24.value
        If CheckBox1.value = True Then
            ws.Cells(IDfier, 45).value = "OK"
        ElseIf CheckBox2.value = True Then
            ws.Cells(IDfier, 45).value = "NG"
        End If
        ws.Cells(IDfier, 46).value = TextBox25.value
        Call ClearFrmCheckItem
        FrmCheckItem.Hide
    End If
End Sub

Sub FrmCheckingItemCmd4()
    If CBContinue.value = False Then
        MsgBox "Please Check the CheckBox!", vbExclamation, "Point of Change"
    ElseIf CBContinue.value = True Then
        Dim ws As Worksheet
        Dim IDfier As Long
        
        Set ws = ThisWorkbook.Sheets("DataReport")
        IDfier = ws.Range("A:A").Find(Label46.Caption, , , xlWhole).Row
        
        ws.Cells(IDfier, 20).value = "Closed"
        ws.Cells(IDfier, 21).value = ""
        ws.Cells(IDfier, 22).value = ""
        ws.Cells(IDfier, 23).value = ""
        ws.Cells(IDfier, 24).value = ""
        ws.Cells(IDfier, 25).value = ""
        ws.Cells(IDfier, 26).value = ""
        ws.Cells(IDfier, 27).value = ""
        ws.Cells(IDfier, 28).value = ""
        ws.Cells(IDfier, 29).value = ""
        ws.Cells(IDfier, 30).value = ""
        ws.Cells(IDfier, 31).value = ""
        ws.Cells(IDfier, 32).value = TextBox12.value
    
        ws.Cells(IDfier, 33).value = TextBox13.value
        ws.Cells(IDfier, 34).value = TextBox14.value
        ws.Cells(IDfier, 35).value = TextBox15.value
        ws.Cells(IDfier, 36).value = TextBox17.value
        ws.Cells(IDfier, 37).value = TextBox18.value
        ws.Cells(IDfier, 38).value = TextBox19.value
        ws.Cells(IDfier, 39).value = TextBox16.value
        ws.Cells(IDfier, 40).value = TextBox20.value
        ws.Cells(IDfier, 41).value = TextBox21.value
        ws.Cells(IDfier, 42).value = TextBox22.value
        ws.Cells(IDfier, 43).value = TextBox23.value
        ws.Cells(IDfier, 44).value = TextBox24.value
        If CheckBox1.value = True Then
            ws.Cells(IDfier, 45).value = "OK"
        ElseIf CheckBox2.value = True Then
            ws.Cells(IDfier, 45).value = "NG"
        Else
            ws.Cells(IDfier, 45).value = ""
        End If
        ws.Cells(IDfier, 46).value = TextBox25.value
        Call ClearFrmCheckItem
        FrmCheckItem.Hide
    End If
End Sub

Sub FrmCheckingItemTxtN()
    If IsNumeric(txtN.value) Then
        Label17.Visible = False
        If txtR.value <> "" Then
            txtHasil.value = (txtR.value / txtN.value)
        Else
        End If
    ElseIf Not IsNumeric(txtN.value) Then
        Label17.Visible = True
        txtN.value = ""
        txtHasil.value = ""
    End If
End Sub

Sub FrmCheckingItemTxtR()
    If IsNumeric(txtR.value) Then
        Label17.Visible = False
        If txtN.value <> "" Then
            txtHasil.value = (txtR.value / txtN.value)
        Else
        End If
    ElseIf Not IsNumeric(txtR.value) Then
        Label17.Visible = True
        txtR.value = ""
        txtHasil.value = ""
    End If
End Sub

Sub FrmCheckingItemTxtHasil()
    Dim lDecPoint As Long

    With Me
        If IsNumeric(txtHasil.value) Then
            lDecPoint = InStr(.txtHasil.value, ".")
            If lDecPoint > 0 And Len(.txtHasil.value) - lDecPoint >= 2 Then
                .txtHasil.value = Left(.txtHasil.value, lDecPoint + 2)
            End If
        End If
    End With
End Sub

Sub FrmStatusActive()
    On Error Resume Next
    Dim ws As Worksheet: Dim Rng As Long
    Set ws = ThisWorkbook.Sheets("DataReport")
    Rng = ws.Cells(Rows.count, 1).End(xlUp).Row
    
    FrmStatus.Image1.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\total.jpg")
    FrmStatus.Image2.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\tprocess.jpg")
    FrmStatus.Image3.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\tmodel.jpg")
    FrmStatus.Image4.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\t5m.jpg")
    FrmStatus.Image5.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\tindication.jpg")
    FrmStatus.Image6.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\tbreakdown.jpg")
    FrmStatus.Image7.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\fman.jpg")
    FrmStatus.Image8.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\fmachine.jpg")
    FrmStatus.Image9.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\fmethod.jpg")
    FrmStatus.Image10.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\fmaterial.jpg")
    FrmStatus.Image11.Picture = LoadPicture(ThisWorkbook.Path & "\Chart\fmeasurement.jpg")
      
    FrmStatus.Label22.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label23.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label24.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label25.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label26.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label27.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label39.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label41.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label43.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label45.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    FrmStatus.Label47.Caption = "Total From Date : " & ws.Range("F3") & " - " & ws.Range("F" & Rng).value
    
    flag = True
    Do While flag = True
        FrmStatus.LblToday1.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.LblToday2.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.LblToday3.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.LblToday4.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.LblToday5.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.LblToday6.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.Label40.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.Label42.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.Label44.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.Label46.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        FrmStatus.Label48.Caption = "Today : " & Format(Now, "dd mmmm yyyy | hh:mm:ss")
        DoEvents
    Loop
End Sub

Sub FrmStatusInitialize()
    On Error Resume Next
    Kill (ThisWorkbook.Path & "\Chart\total.jpg")
    Kill (ThisWorkbook.Path & "\Chart\tprocess.jpg")
    Kill (ThisWorkbook.Path & "\Chart\tmodel.jpg")
    Kill (ThisWorkbook.Path & "\Chart\t5m.jpg")
    Kill (ThisWorkbook.Path & "\Chart\tindication.jpg")
    Kill (ThisWorkbook.Path & "\Chart\tbreakdown.jpg")
    Kill (ThisWorkbook.Path & "\Chart\fman.jpg")
    Kill (ThisWorkbook.Path & "\Chart\fmachine.jpg")
    Kill (ThisWorkbook.Path & "\Chart\fmethod.jpg")
    Kill (ThisWorkbook.Path & "\Chart\fmaterial.jpg")
    Kill (ThisWorkbook.Path & "\Chart\fmeasurement.jpg")
    Call chartExport("total")
    Call chartExport("tprocess")
    Call chartExport("tmodel")
    Call chartExport("t5m")
    Call chartExport("tindication")
    Call chartExport("tbreakdown")
    Call chartExport("fman")
    Call chartExport("fmachine")
    Call chartExport("fmethod")
    Call chartExport("fmaterial")
    Call chartExport("fmeasurement")
End Sub

Sub DataLoginCmd1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DataLogin")
    If ws.Range("K50") <> "ER" Then
    End If
    ws.Range("K50") = "ER"
    ws.Range("K47", "K49").ClearContents
    ThisWorkbook.Application.WindowState = xlMaximized
    ThisWorkbook.Save
    ThisWorkbook.Application.Visible = False
    FrmStart.Show
End Sub

Sub DataLoginCmd2()
    ThisWorkbook.Sheets("Summary").Activate
End Sub
