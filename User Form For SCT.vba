Dim swDoc As SldWorks.ModelDoc2
Dim swDoc1 As SldWorks.ModelDoc2

Private Sub CheckBox1_Click()

Frame2.Enabled = CheckBox3.Value And CheckBox1.Value

End Sub

Private Sub CheckBox3_Change()

Frame2.Enabled = CheckBox3.Value And CheckBox1.Value

End Sub

Private Sub ComboBox1_DropButtonClick()

Dim Gwinty As Variant
Gwinty = Array("M2", "M3", "M4", "M6", "M8", "M10", "M12")
ComboBox1.List() = Gwinty

End Sub

Private Sub ComboBox2_DropButtonClick()

Dim Typy As Variant
Typy = Array("Walcowa", "Stożkowa", "Specjalna")
ComboBox2.List() = Typy

End Sub

Private Sub ExitButton_Click()
    
    Unload Me
    
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub OkButton_Click()

    'Creating new Assembly for the screw
    Set swApp = Application.SldWorks
    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly)
    Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
    'Hiding the Form
    ParametryPołączenia.Hide
    
    ' Check if Solidworks document is opened or not
    If swDoc Is Nothing Then
        MsgBox "Solidworks document haven't opened correctly."
        Exit Sub
    End If
    'Adding the parts:
    If CheckBox2.Value = True Then
        'First Opening the part
        Set swDoc1 = swApp.OpenDoc6("P:\!PRJ_SW\4522 00 0817 HS50 (E240)\4522 200 DA Blacha pod kostkę.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
        'PartName = "P:\!PRJ_SW\!SOLIDWORKS Data\CopiedParts\DIN 931 - M3 X 4 oc.slprt"
        'Set swComponent = swAssembly.AddComponent5(PartName, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)
    End If
    
End Sub

Private Sub SpinButton1_Change()

ParametryPołączenia.TextBox1.Value = SpinButton1.Value

End Sub

Private Sub SpinButton2_Change()

ParametryPołączenia.TextBox2.Value = SpinButton2.Value

End Sub

Private Sub UserForm_Click()

End Sub
