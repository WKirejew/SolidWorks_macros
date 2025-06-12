Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swDoc0 As SldWorks.ModelDoc2
Dim swDoc As SldWorks.ModelDoc2
Dim swDoc1 As SldWorks.ModelDoc2
Dim swDoc2 As SldWorks.ModelDoc2
Dim swDoc3 As SldWorks.ModelDoc2
Dim swDoc4 As SldWorks.ModelDoc2
Dim swAssembly As SldWorks.AssemblyDoc
Dim swcomponent As SldWorks.Component2
Dim swcomponent1 As SldWorks.Component2
Dim swcomponent2 As SldWorks.Component2
Dim swcomponent3 As SldWorks.Component2
Dim swcomponent4 As SldWorks.Component2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swDim As SldWorks.Dimension

Dim dimValue As Variant
Dim partName As String
Dim partName1 As String
Dim partName2 As String
Dim partName3 As String
Dim partName4 As String

Dim Głowa_śruby As Double
Dim D_podkładka As Double
Dim D_sprężyna As Double

Dim swDocExt As ModelDocExtension
Dim swAssy As AssemblyDoc
Dim swModel As ModelDoc2
Dim boolstat As Boolean, stat As Boolean
Dim strings As Variant
Dim matefeature As SldWorks.Feature
Dim MateName As String
Dim FirstSelection As String
Dim SecondSelection As String
Dim Sruba As String
Dim podkładka As String
Dim podkładka2 As String
Dim sprężyna As String
Dim nakrętka As String
Dim AssemblyTitle As String
Dim AssemblyName As String
Dim mateError As Long

Public Rozmiar As String
Public Gwint As String
Public L0 As Double
Public L1 As Double
Public srb As Boolean
Public pdkl1 As Boolean
Public pdkl2 As Boolean
Public pdkls As Boolean
Public nkrtk As Boolean
Public srb_n As String
Public pdkl1_n As String
Public pdkl2_n As String
Public pdkls_n As String
Public nkrtk_n As String


Sub main()

Set swApp = Application.SldWorks
'Creating variable for an assembly template
Dim defaultTemplate As String
defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
'Opening the User Form for data
ParametryPołączenia.Show
MsgBox Rozmiar
'Creating new Assembly
defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly)
Set swDoc0 = swApp.NewDocument(defaultTemplate, 0, 0, 0)
'-------------------------------------------------------
'------------------Adding-Files-------------------------
'-------------------------------------------------------
'Opening 1st file
Set swDoc = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Śruby\" + srb_n + ".SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
partName = swDoc.GetTitle
'Checking if we get a file name
If Len(partName) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'--------------------------------------------------------
Set swDoc1 = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Podkładki\" + pdkl1_n + ".SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "DIN 125 - M8 oc", 0, 0)
partName1 = swDoc1.GetTitle
'Checking if we get a file name
If Len(partName1) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'----------------------------------------------------------
Set swDoc2 = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Podkładki\" + pdkls_n + ".SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "DIN 127 - M8 oc", 0, 0)
partName2 = swDoc2.GetTitle
'Checking if we get a file name
If Len(partName2) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'-----------------------------------------------------------
Set swDoc3 = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Nakrętki\" + nkrtk_n + ".SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
partName3 = swDoc3.GetTitle
'Checking if we get a file name
If Len(partName3) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'-----------------------------------------------------------
Set swDoc4 = swApp.OpenDoc6("P:\!PRJ_SW\!SOLIDWORKS Data\browser\Organic\Podkładki\" + pdkl2_n + ".SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "DIN 125 - M8 oc", 0, 0)
partName4 = swDoc4.GetTitle
'Checking if we get a file name
If Len(partName3) = 0 Then
    MsgBox "Fail to get Part title."
    Exit Sub
End If
'---------------------------------------------------------
'Passing widths of elements
Set swSelMgr = swDoc.SelectionManager

'Finding and Saving thickness of a head of a screw
boolstat = swDoc.Extension.SelectByID2("BaseHead@" + partName, "EXTRUSION", 0, 0, 0, True, 0, Nothing, 0)
Set swDim = swDoc.Parameter("Head_ht@BaseHead")
dimValue = swDim.GetSystemValue3(swThisConfiguration, Empty)
Głowa_śruby = dimValue(0)
'Thickness of a pad
boolstat = swDoc1.Extension.SelectByID2("Base-Revolve@" + partName1, "REVOLUTION", 0, 0, 0, True, 0, Nothing, 0)
Set swDim = swDoc1.Parameter("Thickness@Sketch1")
dimValue = swDim.GetSystemValue3(swThisConfiguration, Empty)
D_podkładka = dimValue(0)
'Thickness of a spring pad
boolstat = swDoc2.Extension.SelectByID2("Base-Revolve@" + partName2, "REVOLUTION", 0, 0, 0, True, 0, Nothing, 0)
Set swDim = swDoc2.Parameter("Thickness@Sketch1")
dimValue = swDim.GetSystemValue3(swThisConfiguration, Empty)
D_sprężyna = dimValue(0)

'---------------------------------------------------------
'Inserting parts into an assembly
Set swAssembly = swDoc0
Set swcomponent = swAssembly.AddComponent5(partName, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)
Set swcomponent1 = swAssembly.AddComponent5(partName1, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0.01, 0, 0)
Set swcomponent2 = swAssembly.AddComponent5(partName2, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0.02, 0, 0)
Set swcomponent3 = swAssembly.AddComponent5(partName3, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0.03, 0, 0)
Set swcomponent4 = swAssembly.AddComponent5(partName4, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0.04, 0, 0)
'Unfixing first component
swcomponent.Select True
swAssembly.UnfixComponent

'Closing opened part files
swApp.CloseDoc partName
swApp.CloseDoc partName1
swApp.CloseDoc partName2
swApp.CloseDoc partName3
'--------------------------------------------------------
'--------------Adding-Mates------------------------------
'--------------------------------------------------------

' Open assembly
Set swModel = swAssembly
Set swAssy = swModel
' Get title of assembly document
AssemblyTitle = swAssy.GetTitle
' Split the title into two strings using the period as the delimiter
strings = Split(AssemblyTitle, ".")
' Use AssemblyName when mating the component with the assembly
AssemblyName = strings(0)
boolstat = True
Set swDocExt = swModel.Extension
Set swSelMgr = swDoc.SelectionManager
'---------------------------------------------------------
' Get the name of the components for the mates
Sruba = swcomponent.Name2()
podkładka = swcomponent1.Name2()
sprężyna = swcomponent2.Name2()
nakrętka = swcomponent3.Name2()
podkładka2 = swcomponent4.Name2()
'-----------------------------------------------------------
'---------------------Mates for the screw-------------------
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba wzdłużne"
FirstSelection = "Plane1@" + Sruba + "@" + AssemblyName
SecondSelection = "Płaszczyzna przednia (XY)@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba poprzeczne"
FirstSelection = "Plane2@" + Sruba & "@" + AssemblyName
SecondSelection = "Płaszczyzna górna (XZ)@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate1
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba czołowe"
FirstSelection = "Plane3@" + Sruba & "@" + AssemblyName
SecondSelection = "Płaszczyzna prawa (ZY)@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
'---------------------Mates for the pad---------------------
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba-podkładka wzdłużne"
FirstSelection = "Plane1@" + Sruba & "@" + AssemblyName
SecondSelection = "Plane1@" + podkładka & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba-podkładka poprzeczne"
FirstSelection = "Plane2@" + Sruba & "@" + AssemblyName
SecondSelection = "Plane2@" + podkładka & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignANTI_ALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Śruba-podkładka czołowe"
FirstSelection = "Plane3@" + Sruba & "@" + AssemblyName
SecondSelection = "Plane3@" + podkładka & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateDISTANCE, swMateAlignANTI_ALIGNED, False, Głowa_śruby, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
'---------------------Mates for the 2nd pad-----------------
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-podkładka wzdłużne"
FirstSelection = "Plane1@" + podkładka & "@" + AssemblyName
SecondSelection = "Plane1@" + podkładka2 & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-podkładka poprzeczne"
FirstSelection = "Plane2@" + podkładka & "@" + AssemblyName
SecondSelection = "Plane2@" + podkładka2 & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignANTI_ALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-podkładka czołowe"
FirstSelection = "Plane3@" + podkładka & "@" + AssemblyName
SecondSelection = "Plane3@" + podkładka2 & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Dim DistDbl As Double
DistDbl = 12 / 1000
Set matefeature = swAssy.AddMate5(swMateDISTANCE, swMateAlignANTI_ALIGNED, True, DistDbl, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
'-------------------Mates for the spring pad----------------
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-sprężyna wzdłużne"
FirstSelection = "Plane1@" + podkładka2 & "@" + AssemblyName
SecondSelection = "Plane1@" + sprężyna & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-sprężyna poprzeczne"
FirstSelection = "Plane2@" + podkładka2 & "@" + AssemblyName
SecondSelection = "Plane2@" + sprężyna & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignANTI_ALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-sprężyna czołowe"
FirstSelection = "Plane3@" + podkładka2 & "@" + AssemblyName
SecondSelection = "Plane3@" + sprężyna & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignANTI_ALIGNED, True, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
'----------------------Mates for the nut--------------------
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-sprężyna wzdłużne"
FirstSelection = "Plane3@" + nakrętka & "@" + AssemblyName
SecondSelection = "Plane1@" + sprężyna & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-sprężyna poprzeczne"
FirstSelection = "Plane2@" + nakrętka & "@" + AssemblyName
SecondSelection = "Plane2@" + sprężyna & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateCOINCIDENT, swMateAlignANTI_ALIGNED, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)
'-----------------------------------------------------------
' Create the name of the mate and the names of the planes to use for the mate
MateName = "Podkładka-sprężyna czołowe"
FirstSelection = "Plane1@" + nakrętka & "@" + AssemblyName
SecondSelection = "Plane3@" + sprężyna & "@" + AssemblyName
swModel.ClearSelection2 (True)
' Select the planes to mate
boolstat = swDocExt.SelectByID2(FirstSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
boolstat = swDocExt.SelectByID2(SecondSelection, "PLANE", 0, 0, 0, True, 1, Nothing, swSelectOptionDefault)
' Add the mate
Set matefeature = swAssy.AddMate5(swMateDISTANCE, swMateAlignALIGNED, False, D_sprężyna, 0, 0, 0, 0, 0, 0, 0, False, False, 0, mateError)
matefeature.name = MateName
swModel.ClearSelection2 (True)


End Sub

Sub Case_śruba(name As String)


End Sub
