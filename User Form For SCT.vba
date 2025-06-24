Private Sub UserForm_Initialize()

Dim A As String
Dim B As String
Dim C As String
Dim D As String
Dim E As String
Dim F As String
Dim G As String
Dim H As String
Dim I As String
Dim K As String
Dim L As String
Dim M As String
Dim N As String
Dim O As String
Dim P As String
Dim R As String
Dim S As String
Dim T As String
Dim U As String
Dim V As String
Dim W As String
Dim Y As String
Dim Z As String
Dim X As String
Dim AA As String
Dim AB As String
Dim swDoc As SldWorks.ModelDoc2
Dim swDoc1 As SldWorks.ModelDoc2

End Sub

Private Sub CheckBox1_Click()

Frame2.Enabled = CheckBox3.Value And CheckBox1.Value
CheckBox3.Enabled = CheckBox1.Value
If CheckBox1.Value = False Then
    CheckBox3.Value = CheckBox1.Value
End If
ComboBox3.Enabled = CheckBox1.Value

End Sub

Private Sub CheckBox3_Change()

Frame2.Enabled = CheckBox3.Value And CheckBox1.Value
ComboBox5.Enabled = CheckBox3.Value
If CheckBox3.Value = True Then
    ComboBox5.Value = ComboBox3.Value
End If

End Sub

Private Sub CheckBox4_Change()

ComboBox6.Enabled = CheckBox4.Value

End Sub

Private Sub CheckBox5_Change()

ComboBox7.Enabled = CheckBox5.Value

End Sub

Private Sub ComboBox1_DropButtonClick()

Dim Gwinty As Variant
Gwinty = Array("M2", "M3", "M4", "M6", "M8", "M10", "M12")
ComboBox1.List() = Gwinty

End Sub

Private Sub ComboBox3_Change()

If CheckBox3.Value = True Then
    ComboBox5.Value = ComboBox3.Value
End If

End Sub

Private Sub ComboBox3_DropButtonClick()

Dim Pdkl As Variant
Pdkl = Array("Podkładka okrągła czarna oksydacja DIN 125", "Podkładka okrągła ocynk DIN 125", "Podkładka okrągła stal nierdzewna DIN 125", "Podkładka okrągła powiększona czarna oksydacja DIN_9021", "Podkładka okrągła powiększona ocynk DIN 9021", "Podkładka okrągła powiększona stal nierdzewna DIN 9021", "Podkładka falista czarna oksydacja DIN 137", "Podkładka falista ocynk DIN 137", "Podkładka ząbkowana czarna oksydacja DIN 6798", "Podkładka ząbkowana ocynk DIN 6798", "Podkładka ząbkowana stal nierdzewna DIN 6798")
ComboBox3.List() = Pdkl

End Sub

Private Sub ComboBox4_DropButtonClick()

Dim Typy As Variant
Typy = Array("Śruba z łbem sześciokątnym czarna oksydacja DIN 931", "Śruba z łbem sześciokątnym ocynk DIN 931", "Śruba z łbem sześciokątnym ocynk DIN 933", "Śruba z łbem sześciokątnym stal nierdzewna DIN 931", "Wkręt soczewka ampul czarna oksydacja DIN 7380", "Wkręt soczewka ampul ocynk DIN 7380", "Wkręt soczewka ampul stal nierdzewna DIN 7380", "Wkręt stożkowy ampul czarna oksydacja DIN 7991", "Wkręt stożkowy ampul czarna stal nierdzewna DIN 7991", "Wkręt stożkowy ampul ocynk DIN 7991", "Wkręt stożkowy krzyżak czarna oksydacja DIN 965", "Wkręt stożkowy krzyżak ocynk DIN_965", "Wkręt stożkowy krzyżak stal nierdzewna DIN 965", "Śruba z łbem radełkowanym czarna oksydacja DIN 653", "Śruba z łbem radełkowanym niskim ocynk DIN 653", "Śruba z łbem radełkowanym stal nierdzewna DIN 653")
ComboBox4.List() = Typy

End Sub

Private Sub ComboBox4_Change()

End Sub


Private Sub ComboBox5_Change()


End Sub

Private Sub ComboBox5_DropButtonClick()

Dim pdkl1 As Variant
pdk1 = Array("Podkładka okrągła czarna oksydacja DIN 125", "Podkładka okrągła ocynk DIN 125", "Podkładka okrągła stal nierdzewna DIN 125", "Podkładka okrągła powiększona czarna oksydacja DIN_9021", "Podkładka okrągła powiększona ocynk DIN 9021", "Podkładka okrągła powiększona stal nierdzewna DIN 9021", "Podkładka falista czarna oksydacja DIN 137", "Podkładka falista ocynk DIN 137", "Podkładka ząbkowana czarna oksydacja DIN 6798", "Podkładka ząbkowana ocynk DIN 6798", "Podkładka ząbkowana stal nierdzewna DIN 6798")
ComboBox5.List() = pdk1

End Sub

Private Sub ComboBox6_Change()

End Sub

Private Sub ComboBox6_DropButtonClick()

Dim pds As Variant
pds = Array("Podkładka sprężysta czarna oksydacja DIN 127", "Podkładka sprężysta ocynk DIN 127", "Podkładka sprężysta stal nierdzewna DIN 127")
ComboBox6.List() = pds


End Sub

Private Sub ComboBox7_Change()

Select Case ComboBox7.Value
Case A
End Select
End Sub

Private Sub ComboBox7_DropButtonClick()


A = "Nakrętka sześciokątna niska stal nierdzewna DIN 439"
B = "Nakrętka sześciokątna ocynk DIN 934"
C = "Nakrętka sześciokątna samokontrująca czarna oksydacja DIN 985"
D = "Nakrętka sześciokątna samokontrująca ocynk DIN 985"
E = "Nakrętka sześciokątna samokontrująca stal nierdzewna DIN 985"
F = "Nakrętka sześciokątna stal nierdzewna DIN 934"
G = "Nakrętka z uchem czarna oksydacja DIN 582"
H = "Nakrętka z uchem ocynk DIN 582"
I = "Nakrętka z uchem stal nierdzewna DIN 582"
J = "Nakrętka kołpakowa czarna oksydacja DIN 1587"
K = "Nakrętka kołpakowa drobnozwojna czarna oksydacja DIN 1587"
L = "Nakrętka kołpakowa drobnozwojna ocynk DIN 1587"
M = "Nakrętka kołpakowa drobnozwojna stal nierdzewna DIN 1587"
N = "Nakrętka kołpakowa ocynk DIN 1587"
O = "Nakrętka kołpakowa stal nierdzewna DIN 1587"
P = "Nakrętka skrzydełkowa czarna oksydacja DIN 315"
R = "Nakrętka skrzydełkowa ocynk DIN 315"
S = "Nakrętka skrzydełkowa stal nierdzewna DIN 315"
T = "Nakrętka skrzydełkowa stal nierdzewna DIN 315"
U = "Nakrętka sześciokątna drobnozwojna czarna oksydacja DIN 934"
V = "Nakrętka sześciokątna drobnozwojna ocynk DIN 934"
W = "Nakrętka sześciokątna drobnozwojna stal nierdzewna DIN 934"
X = "Nakrętka sześciokątna niska czarna oksydacja DIN 439"
Y = "Nakrętka sześciokątna niska drobnozwojna czarna oksydacja DIN 439"
Z = "Nakrętka sześciokątna niska drobnozwojna ocynk DIN 439"
AA = "Nakrętka sześciokątna niska drobnozwojna stal nierdzewna DIN 439"
AB = "Nakrętka sześciokątna niska ocynk DIN 439"

Dim Typy As Variant
Typy = Array(B, F, AB, A, C, D, E, U, V, W, X, Y, Z, AA, P, R, S, T, J, K, L, M, N, O, G, H, I)
ComboBox7.List() = Typy

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub SpinButton1_Change()

ParametryPołączenia.TextBox1.Value = SpinButton1.Value

End Sub

Private Sub SpinButton2_Change()

ParametryPołączenia.TextBox2.Value = SpinButton2.Value

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub ExitButton_Click()
    
    Exit1 = True
    Unload Me
    
End Sub

Private Sub OkButton_Click()
    
Exit1 = False
Select Case ComboBox3.Value
    Case "Podkładka falista czarna oksydacja DIN 137"
        pdk_DIN = "DIN 137"
        pdk_mat = " co"
    Case "Podkładka falista ocynk DIN 137"
        pdk_DIN = "DIN 137"
        pdk_mat = " oc"
    'Case "Podkładka falista stal nierdzewna DIN 137"
        'pdk_DIN = "DIN 137"
        'pdk_mat = " sn"
    Case "Podkładka okrągła czarna oksydacja DIN 125"
        pdk_DIN = "DIN 125"
        pdk_mat = " co"
    Case "Podkładka okrągła ocynk DIN 125"
        pdk_DIN = "DIN 125"
        pdk_mat = " oc"
    Case "Podkładka okrągła stal nierdzewna DIN 125"
        pdk_DIN = "DIN 125"
        pdk_mat = " sn"
    Case "Podkładka okrągła powiększana czarna oksydacja DIN_9021"
        pdk_DIN = "DIN 9021"
        pdk_mat = " co"
    Case "Podkładka okrągła powiększana ocynk DIN 9021"
        pdk_DIN = "DIN 9021"
        pdk_mat = " oc"
    Case "Podkładka okrągła powiększana stal nierdzewna DIN 9021"
        pdk_DIN = "DIN 9021"
        pdk_mat = " sn"
    Case "Podkładka ząbkowana czarna oksydacja DIN 6798"
        pdk_DIN = "DIN 6798"
        pdk_mat = " co"
    Case "Podkładka ząbkowana ocynk DIN 6798"
        pdk_DIN = "DIN 6798"
        pdk_mat = " oc"
    Case "Podkładka ząbkowana stal nierdzewna DIN 6798"
        pdk_DIN = "DIN 6798"
        pdk_mat = " sn"
End Select
 
Select Case ComboBox4.Value
    Case "Śruba z łbem radełkowanym czarna oksydacja DIN 653"
        srb_DIN = "DIN 653"
        srb_mat = " co"
        srb_typ = "Śruby\"
    Case "Śruba z łbem radełkowanym ocynk DIN 653"
        srb_DIN = "DIN 653"
        srb_mat = " oc"
        srb_typ = "Śruby\"
    Case "Śruba z łbem radełkowanym stal nierdzewna DIN 653"
        srb_DIN = "DIN 653"
        srb_mat = " sn"
        srb_typ = "Śruby\"
    Case "Śruba z łbem sześciokątnym czarna oksydacja DIN 931"
        srb_DIN = "DIN 931"
        srb_mat = " co"
        srb_typ = "Śruby\"
    Case "Śruba z łbem sześciokątnym ocynk DIN 931"
        srb_DIN = "DIN 931"
        srb_mat = " oc"
        srb_typ = "Śruby\"
    Case "Śruba z łbem sześciokątnym stal nierdzewna DIN 931"
        srb_DIN = "DIN 931"
        srb_mat = " sn"
        srb_typ = "Śruby\"
    Case "Śruba z łbem sześciokątnym ocynk DIN 933"
        srb_DIN = "DIN 933"
        srb_mat = " oc"
        srb_typ = "Śruby\"
    Case "Wkręt soczewka ampul czarna oksydacja DIN 7380"
        srb_DIN = "DIN 7380"
        srb_mat = " co"
        srb_typ = "Wkręty soczewki\"
    Case "Wkręt soczewka ampul ocynk DIN 7380"
        srb_DIN = "DIN 7380"
        srb_mat = " oc"
        srb_typ = "Wkręty soczewki\"
    Case "Wkręt soczewka ampul stal nierdzewna DIN 7380"
        srb_DIN = "DIN 7380"
        srb_mat = " sn"
        srb_typ = "Wkręty soczewki\"
    Case "Wkręt stożkowy ampul czarna oksydacja DIN 7991"
        srb_DIN = "DIN 7991"
        srb_mat = " co"
        srb_typ = "Wkręty stożki\"
    Case "Wkręt stożkowy ampul ocynk DIN 7991"
        srb_DIN = "DIN 7991"
        srb_mat = " oc"
        srb_typ = "Wkręty stożki\"
    Case "Wkręt stożkowy ampul stal nierdzewna DIN 7991"
        srb_DIN = "DIN 7991"
        srb_mat = " sn"
        srb_typ = "Wkręty stożki\"
    Case "Wkręt stożkowy krzyżak czarna oksydacja DIN 965"
        srb_DIN = "DIN 965"
        srb_mat = " co"
        srb_typ = "Wkręty stożki\"
    Case "Wkręt stożkowy krzyżak ocynk DIN 965"
        srb_DIN = "DIN 965"
        srb_mat = " oc"
        srb_typ = "Wkręty stożki\"
    Case "Wkręt stożkowy krzyżak stal nierdzewna DIN 965"
        srb_DIN = "DIN 965"
        srb_mat = " sn"
        srb_typ = "Wkręty stożki\"
End Select

Select Case ComboBox5.Value
    Case "Podkładka falista czarna oksydacja DIN 137"
        pd2_DIN = "DIN 137"
        pd2_mat = " co"
    Case "Podkładka falista ocynk DIN 137"
        pd2_DIN = "DIN 137"
        pd2_mat = " oc"
    'Case "Podkładka falista stal nierdzewna DIN 137"
        'pd2_DIN = "DIN 137"
        'pd2_mat = " sn"
    Case "Podkładka okrągła czarna oksydacja DIN 125"
        pd2_DIN = "DIN 125"
        pd2_mat = " co"
    Case "Podkładka okrągła ocynk DIN 125"
        pd2_DIN = "DIN 125"
        pd2_mat = " oc"
    Case "Podkładka okrągła stal nierdzewna DIN 125"
        pd2_DIN = "DIN 125"
        pd2_mat = " sn"
    Case "Podkładka okrągła powiększana czarna oksydacja DIN_9021"
        pd2_DIN = "DIN 9021"
        pd2_mat = " co"
    Case "Podkładka okrągła powiększana ocynk DIN 9021"
        pd2_DIN = "DIN 9021"
        pd2_mat = " oc"
    Case "Podkładka okrągła powiększana stal nierdzewna DIN 9021"
        pd2_DIN = "DIN 9021"
        pd2_mat = " sn"
    Case "Podkładka ząbkowana czarna oksydacja DIN 6798"
        pd2_DIN = "DIN 6798"
        pd2_mat = " co"
    Case "Podkładka ząbkowana ocynk DIN 6798"
        pd2_DIN = "DIN 6798"
        pd2_mat = " oc"
    Case "Podkładka ząbkowana stal nierdzewna DIN 6798"
        pd2_DIN = "DIN 6798"
        pd2_mat = " sn"
End Select
  
Select Case ComboBox6.Value
    Case "Podkładka sprężysta czarna oksydacja DIN 127"
        pds_DIN = "DIN 127"
        pds_mat = " co"
    Case "Podkładka sprężysta ocynk DIN 127"
        pds_DIN = "DIN 127"
        pds_mat = " oc"
    Case "Podkładka sprężysta stal nierdzewna DIN 127"
        pds_DIN = "DIN 127"
        pds_mat = " sn"
End Select

    Rozmiar = ComboBox1.Value
    L0 = TextBox1.Value
    L1 = TextBox2.Value
    srb = CheckBox2.Value
    pdkl1 = CheckBox1.Value
    pdkl2 = CheckBox3.Value
    pdkls = CheckBox4.Value
    nkrtk = CheckBox5.Value
    srb_n = ComboBox4.Value
    pdkl1_n = ComboBox3.Value
    pdkl2_n = ComboBox5.Value
    pdkls_n = ComboBox6.Value
    nkrtk_n = ComboBox7.Value
    Unload ParametryPołączenia

End Sub

