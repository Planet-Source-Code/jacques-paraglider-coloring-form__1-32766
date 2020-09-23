Attribute VB_Name = "CouleurFenetre"
Public Couleur$(200)
Public TempoCouleur$
Public intStartX%
Public intStartY%
Public intLastX%
Public intLastY%


Public Sub CoulFenetre(Obj As Object)

' *************** Couleur de la Fenêtre ****************
Couleur$(0) = "220/226\222"  ' Couleur Centrale
Couleur$(1) = "0/47\60"   ' couleur arrondi
Couleur$(2) = "0/59\75"   ' couleur arrondi
Couleur$(3) = "0/70\90"   ' couleur arrondi
Couleur$(4) = "0/82\105"   ' couleur arrondi
Couleur$(5) = "0/94\120"    ' couleur arrondi
Couleur$(6) = "0/82\105"    ' couleur arrondi
Couleur$(7) = "0/70\90"     ' couleur arrondi
Couleur$(8) = "0/59\75"     ' couleur arrondi
Couleur$(9) = "0/47\60"     ' couleur arrondi
Couleur$(10) = "0/94\120"    ' derniere couleur arrondi
Couleur$(11) = "41/36\24"    ' ombre
Couleur$(12) = "49/52\30"    ' ombre
Couleur$(13) = "0/47\60"    ' dégradé vers couleur centrale
Couleur$(14) = "0/47\60"    ' dégradé vers couleur centrale
Couleur$(15) = "0/59\75"    ' dégradé vers couleur centrale
Couleur$(16) = "0/59\75"    ' dégradé vers couleur centrale
Couleur$(17) = "0/70\90"   ' dégradé vers couleur centrale
Couleur$(18) = "0/70\90" ' dégradé vers couleur centrale
Couleur$(19) = "0/82\105" ' dégradé vers couleur centrale
Couleur$(20) = "0/94\120" '
Couleur$(21) = "0/106\135" '
Couleur$(22) = "0/117\150" '
TempoCouleur$ = Couleur$(0)
intStartX = 0
intStartY = 0
intLastX = Obj.Width - 1
intLastY = Obj.Height - 1
Enlarge% = 15
Obj.AutoRedraw = True
For X = 1 To 22
TempoCouleur$ = Couleur$(X)
Call ExtraitCouleur(TempoCouleur$, CouleurExadecimal$)
CouleurT = Val(CouleurExadecimal$)
Obj.ForeColor = CouleurT
Obj.Line (intStartX, intStartY)-(intLastX, intLastY), , B
    intStartX = intStartX + Enlarge%
    intStartY = intStartY + Enlarge%
    intLastX = intLastX - Enlarge%
    intLastY = intLastY - Enlarge%
Next
TempoCouleur$ = Couleur$(0)
Call ExtraitCouleur(TempoCouleur$, CouleurExadecimal$)
CouleurT = Val(CouleurExadecimal$)
Obj.ForeColor = CouleurT
Obj.Line (intStartX, intStartY)-(intLastX, intLastY), , BF

End Sub

Private Sub ExtraitCouleur(TempoCouleur$, CouleurExadecimal$)

If TempoCouleur$ = "" Then Exit Sub
R1$ = "/"
R2$ = "\"
Pos1% = InStr(TempoCouleur$, R1$)
Pos2% = InStr(TempoCouleur$, R2$)
Dif% = (Pos2% - Pos1%) - 1
Rouge$ = Left$(TempoCouleur$, Pos1% - 1)
Vert$ = Mid$(TempoCouleur$, Pos1% + 1, Dif%)
Bleu$ = Mid$(TempoCouleur$, Pos2% + 1, 3)
R% = Val(Rouge$): V% = Val(Vert$): B% = Val(Bleu$)
        total = " RGB(" + totalRouge + "," + totalVert + "," + totalBleu + ")"
        BB$ = Hex(B%)
        If Len(BB$) = 1 Then BB$ = "0" + BB$
        GG$ = Hex(V%)
        If Len(GG$) = 1 Then GG$ = "0" + GG$
        RR$ = Hex(R%)
        If Len(RR$) = 1 Then RR$ = "0" + RR$
        CouleurExadecimal$ = "&H00" + BB$ + GG$ + RR$ + "&"
End Sub



