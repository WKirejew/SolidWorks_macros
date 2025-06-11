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
Dim swDocExt As ModelDocExtension
Dim swAssy As AssemblyDoc
Dim swModel As ModelDoc2
Dim boolstat As Boolean, stat As Boolean
Dim strings As Variant
Dim matefeature As SldWorks.Feature
Dim MateName As String
Dim FirstSelection As String
Dim SecondSelection As String
Dim śruba As String
Dim podkładka As String
Dim sprężyna As String
Dim nakrętka As String
Dim AssemblyTitle As String
Dim AssemblyName As String
Dim mateError As Long

' Open assembly
Set swModel = swAssembly
Set swAssy = swModel
' Get title of assembly document
AssemblyTitle = swModel.GetTitle
' Split the title into two strings using the period as the delimiter
strings = Split(AssemblyTitle, ".")
' Use AssemblyName when mating the component with the assembly
AssemblyName = strings(0)
boolstat = True
'---------------------------------------------------------
' Get the name of the components for the mates
śruba = swcomponent.Name2()
podkładka = swcomponent1.Name2()
sprężyna = swcomponent2.Name2()
nakrętka = swcomponent3.Name2()
'-----------------------------------------------------------
'---------------------Mates for the screw-------------------
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba wzdłużne"
FirstSelection = "Plane1@" + śruba + "@" + AssemblyName
SecondSelection = "Płaszczyzna przednia (XY)@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.Name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba poprzeczne"
FirstSelection = "Plane2@" + śruba & "@" + AssemblyName
SecondSelection = "Płaszczyzna górna (XZ)@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate1
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.Name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba czoło"
FirstSelection = "Plane3@" + śruba & "@" + AssemblyName
SecondSelection = "Płaszczyzna prawa (ZY)@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.Name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
'---------------------Mates for the pad---------------------
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba-podkładka wzdłużne"
FirstSelection = "Plane1@" + śruba & "@" + AssemblyName
SecondSelection = "Plane1@" + podkładka & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.Name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba-podkładka wzdłużne"
FirstSelection = "Plane2@" + śruba & "@" + AssemblyName
SecondSelection = "Plane2@" + podkładka & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.Name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba-podkładka wzdłużne"
FirstSelection = "Plane3@" + śruba & "@" + AssemblyName
SecondSelection = "Plane3@" + podkładka & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0.075, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.Name = MateName
swModel.ClearSelection2 (True)

End Sub
