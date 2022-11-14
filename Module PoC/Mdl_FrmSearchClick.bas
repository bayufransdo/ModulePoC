Attribute VB_Name = "Mdl_FrmSearchClick"
Option Explicit

'This for  FrmSearchClick Update button
Sub UpdateButton()
 
End Sub

'FrmSearchClick value
Sub FrmSearchClickValue()
    'To fill all textbox and comboBox value based on Listbox value that user double clicking
    FrmSearchClick.ComboPOC.value = FrmSearch.ListBox1.Column(3)
    FrmSearchClick.ComboProductSafety.value = FrmSearch.ListBox1.Column(4)
    FrmSearchClick.TxtDate.value = Format(FrmSearch.ListBox1.Column(5), "dd mmmm yyyy")
    FrmSearchClick.TxtTime.value = Format(FrmSearch.ListBox1.Column(6), "hh:mm:ss")
    FrmSearchClick.TxtModelType.value = FrmSearch.ListBox1.Column(7)
    FrmSearchClick.TxtPartNo.value = FrmSearch.ListBox1.Column(8)
    FrmSearchClick.TxtMachineNo.value = FrmSearch.ListBox1.Column(9)
    FrmSearchClick.TxtMachineName.value = FrmSearch.ListBox1.Column(10)
    FrmSearchClick.TxtAffectedLot.value = FrmSearch.ListBox1.Column(12)
    FrmSearchClick.ComboChangeContent.value = FrmSearch.ListBox1.Column(13)
    FrmSearchClick.ComboProcess.value = FrmSearch.ListBox1.Column(14)
    FrmSearchClick.TxtProcessNo.value = FrmSearch.ListBox1.Column(15)
    FrmSearchClick.TxtTreatment.value = FrmSearch.ListBox1.Column(16)
    FrmSearchClick.ComboCC.value = FrmSearch.ListBox1.Column(17)
    FrmSearchClick.TxtStatus.value = FrmSearch.ListBox1.Column(18)
    FrmSearchClick.Caption = "Update Form No : " & FrmSearch.ListBox1.Column(0)
    FrmSearchClick.Show
End Sub

Sub FrmSearchclickStatus()
    
End Sub
