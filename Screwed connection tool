Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swDoc As SldWorks.ModelDoc2
Dim swDoc1 As SldWorks.ModelDoc2
Dim swDoc2 As SldWorks.ModelDoc2
Dim swDoc3 As SldWorks.ModelDoc2
Dim swDoc4 As SldWorks.ModelDoc2
Dim swAssembly As SldWorks.AssemblyDoc
Dim swComponent As SldWorks.Component2


Dim partName As String
Dim partName1 As String
Dim partName2 As String
Dim partName3 As String

Sub main()

Set swApp = Application.SldWorks
'Creating variable for an assembly template
Dim defaultTemplate As String
defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)


'Creating new Assembly
defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly)
Set swDoc4 = swApp.NewDocument(defaultTemplate, 0, 0, 0)
'-------------------------------------------------------
'------------------Adding-Files-------------------------
'-------------------------------------------------------
'Opening 1st file
Set swDoc = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Śruby\Śruba z łbem sześciokątnym czarna oksydacja DIN 931.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
partName = swDoc.GetTitle
'Checking if we get a file name
If Len(partName) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'Inserting part into an assembly
'Set swAssembly = swDoc4
'Set swComponent = swAssembly.AddComponent5(partName, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)
'--------------------------------------------------------
Set swDoc1 = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Podkładki\Podkładka okrągła ocynk DIN 125.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "DIN 125 - M8 oc", 0, 0)
partName1 = swDoc1.GetTitle
'Checking if we get a file name
If Len(partName1) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'Inserting part into an assembly
'Set swAssembly = swDoc4
'Set swComponent = swAssembly.AddComponent5(partName1, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)
'----------------------------------------------------------
Set swDoc2 = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Podkładki\Podkładka sprężysta ocynk DIN 127.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "DIN 127 - M8 oc", 0, 0)
partName2 = swDoc2.GetTitle
'Checking if we get a file name
If Len(partName2) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'Inserting part into an assembly
'Set swAssembly = swDoc4
'Set swComponent = swAssembly.AddComponent5(partName2, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)
'-----------------------------------------------------------
Set swDoc3 = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Nakrętki\Nakrętka sześciokątna drobnozwojna stal nierdzewna DIN 934.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
partName3 = swDoc3.GetTitle
'Checking if we get a file name
If Len(partName3) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'---------------------------------------------------------
'Inserting parts into an assembly
Set swAssembly = swDoc4
Set swComponent = swAssembly.AddComponent5(partName, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)
Set swComponent = swAssembly.AddComponent5(partName1, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0.01, 0, 0)
Set swComponent = swAssembly.AddComponent5(partName2, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0.02, 0, 0)
Set swComponent = swAssembly.AddComponent5(partName3, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0.03, 0, 0)

'Closing opened part files
swApp.CloseDoc partName
swApp.CloseDoc partName1
swApp.CloseDoc partName2
swApp.CloseDoc partName3
'--------------------------------------------------------
'--------------Adding-Mates------------------------------
'--------------------------------------------------------

End Sub
