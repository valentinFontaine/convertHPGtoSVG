Option Explicit

Dim args
Set args = Wscript.Arguments
AnalyseFichier args(0), args(1) 


Sub AnalyseFichier(inputFile, outputFile)

    Dim ecritureStream 
    Dim lectureStream 
    
    Dim mode, taillePolice, ligne, groupeRaster, parametreRaster, commande, terminaisonRaster    'Strings
    Dim nbLigne
    Dim texte
    Dim penDown   ' boolean

    Dim inputP(3), echelle(3), anciennePosition(1), nouvellePosition(1), tab_1parametres(0), tab_2parametres(1), tab_4parametres(3)
    Dim i, j     ' Integer
    Dim angle, largeurTrait 'Double
    Dim pointilles, couleur

    Dim horaire               
    Dim grandArc              
    Dim rayon                 
    Dim rayon_ellipse(1), point_centre(1)
    Dim angle_base             
    Dim pi

    Dim terminator, imprimeTerm, anchor, baseLineShift
    Dim labelOrigin(0)

    Dim polyligne
    polyligne=array()

    pi = 3.14159265 
    mode = "Raster"
    penDown = False
    
    Set ecritureStream = CreateObject("ADODB.Stream")
    Set lectureStream = CreateObject("ADODB.Stream")

    lectureStream.Charset = "iso-8859-1"
    lectureStream.Open
    lectureStream.LoadFromFile inputFile

    ecritureStream.Charset = "utf-8"
    ecritureStream.Open
    ecritureStream.WriteText "<?xml version=""1.0"" standalone=""no""?>" & vbNewLine
    ecritureStream.WriteText "<!DOCTYPE svg PUBLIC ""-//W3C//DTD SVG 1.1//EN"" ""http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd"">" & vbNewLine
    ecritureStream.WriteText "<svg width=""210mm"" height=""297mm"" viewBox=""0 0 210 297"" version=""1.1"" xmlns = ""http://www.w3.org/2000/svg"" > " & vbNewLine

    ecritureStream.WriteText ligne & vbNewLine
    ecritureStream.WriteText Mid(ligne, 5, 5) & vbNewLine

    nbLigne = 0
    While Not lectureStream.EOS
    
        nbLigne = nbLigne + 1
        ligne = lectureStream.ReadText(-2)
        For i = 1 To Len(ligne)
            
            If IsLetter(Mid(ligne, i, 1)) And mode = "HPG-2L" Then
                i = i + 1
                commande = Mid(ligne, i - 1, 1) & Mid(ligne, i, 1)

                If commande = "IN" Then
                    'Initialisation de toutes les variables.
                    inputP(0) = 0
                    inputP(1) = 0
                    inputP(2) = 0
                    inputP(3) = 0
                    echelle(0) = 0
                    echelle(1) = 0
                    echelle(2) = 0
                    echelle(3) = 0
                    anciennePosition(0) = 0
                    anciennePosition(1) = 0
                    nouvellePosition(0) = 0
                    nouvellePosition(1) = 0
                    tab_2parametres(0) = 0
                    tab_2parametres(1) = 0
                    penDown = False
                    angle = 0
                    taillePolice = "10mm"
                    largeurTrait = 0.35
                    couleur = "black"

                ElseIf commande = "IP" Then
                
                    Call rempliTableau(i, ligne, inputP)

                    tab_4parametres(0) = 2.5 + inputP(0) * 0.025
                    tab_4parametres(1) = 297 - 10 - inputP(1) * 0.025
                    tab_4parametres(2) = 2.5 + inputP(2) * 0.025
                    tab_4parametres(3) = 297 - 10 - inputP(3) * 0.025
                
                    ecritureStream.WriteText "<rect x=""" & Replace(Min(tab_4parametres(0), tab_4parametres(2)), ",", ".") & """ y=""" & Replace(Min(tab_4parametres(1), tab_4parametres(3)), ",", ".") & """ width=""" & Replace(Abs(tab_4parametres(2) - tab_4parametres(0)), ",", ".") & """ height=""" & Replace(Abs(tab_4parametres(1) - tab_4parametres(3)), ",", ".") & """ " & pointilles & " style=""fill:none;stroke-linecap:round;stroke-width:0.05;stroke:powderblue""  />" & vbNewLine
                ElseIf commande = "SC" Then
                    Call rempliTableau(i, ligne, echelle)

                '--------------- Position Stylo ----------------    
                ElseIf commande = "PU" Then
                    penDown = False
                    If rempliPolyligne(i, ligne, polyligne) Then
                        For j = LBound(polyligne, 2) To UBound(polyligne, 2)
                            tab_2parametres(0) = polyligne(0, j)
                            tab_2parametres(1) = polyligne(1, j)
                            Call changeRepere(tab_2parametres, inputP, echelle, False)
                            anciennePosition(0) = nouvellePosition(0)
                            anciennePosition(1) = nouvellePosition(1)
                            nouvellePosition(0) = tab_2parametres(0)
                            nouvellePosition(1) = tab_2parametres(1)
                        Next
                    End If
                ElseIf commande = "PD" Then
                    penDown = True
                    If rempliPolyligne(i, ligne, polyligne) Then
                        texte = Replace(nouvellePosition(0), ",", ".") & "," & Replace(nouvellePosition(1), ",", ".")
                        For j = LBound(polyligne, 2) To UBound(polyligne, 2)
                            tab_2parametres(0) = polyligne(0, j)
                            tab_2parametres(1) = polyligne(1, j)
                            
                            Call changeRepere(tab_2parametres, inputP, echelle, False)
                            
                            anciennePosition(0) = nouvellePosition(0)
                            anciennePosition(1) = nouvellePosition(1)
                            nouvellePosition(0) = tab_2parametres(0)
                            nouvellePosition(1) = tab_2parametres(1)
                            texte = texte & " " & Replace(tab_2parametres(0), ",", ".") & "," & Replace(tab_2parametres(1), ",", ".")
                        Next
                        ecritureStream.WriteText "<polyline points=""" & texte & """ " & pointilles & " style=""fill:none;stroke:" & couleur & ";stroke-linecap:round;stroke-linejoin:round;stroke-width:" & Replace(largeurTrait, ",", ".") & """ />" & vbNewLine
              
                    End If
                ElseIf commande = "PA" Then
                    anciennePosition(0) = nouvellePosition(0)
                    anciennePosition(1) = nouvellePosition(1)
                    Call rempliTableau(i, ligne, nouvellePosition)
                    Call changeRepere(nouvellePosition, inputP, echelle, false)
                    If penDown Then
                        ecritureStream.WriteText "<line x1=""" & Replace(anciennePosition(0), ",", ".") & """ x2=""" & Replace(nouvellePosition(0), ",", ".") & """ y1=""" & Replace(anciennePosition(1), ",", ".") & """ y2=""" & Replace(nouvellePosition(1), ",", ".") & """ " & pointilles & " style=""fill:none;stroke:" & couleur & " ;stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & """ />" & vbNewLine
                    End If   

                 ElseIf commande = "PR" Then
                    anciennePosition(0) = nouvellePosition(0)
                    anciennePosition(1) = nouvellePosition(1)
                    Call rempliTableau(i, ligne, nouvellePosition)
                    Call changeRepere(nouvellePosition, inputP, echelle, True)
                    nouvellePosition(0) = nouvellePosition(0) + anciennePosition(0)
                    nouvellePosition(1) = nouvellePosition(1) + anciennePosition(1)
                    If penDown Then
                        ecritureStream.WriteText "<line x1=""" & Replace(anciennePosition(0), ",", ".") & """ x2=""" & Replace(nouvellePosition(0), ",", ".") & """ y1=""" & Replace(anciennePosition(1), ",", ".") & """ y2=""" & Replace(nouvellePosition(1), ",", ".") & """ " & pointilles & " style=""fill:none;stroke:" & couleur & ";stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & """ />" & vbNewLine
                    End If

                ' ------------ Rectangles -----------
                 ElseIf commande = "EA" Then
                    Call rempliTableau(i, ligne, tab_2parametres)
                    Call changeRepere(tab_2parametres, inputP, echelle, False)
                  
                    If tab_2parametres(0) - nouvellePosition(0) <> 0 And tab_2parametres(1) - nouvellePosition(1) <> 0 Then
                        ecritureStream.WriteText "<rect x=""" & Replace(Min(nouvellePosition(0), tab_2parametres(0)), ",", ".") & _
                        """ y=""" & Replace(Min(nouvellePosition(1), tab_2parametres(1)), ",", ".") & _
                        """ width=""" & Replace(Abs(tab_2parametres(0) - nouvellePosition(0)), ",", ".") & _
                        """ height=""" & Replace(Abs(tab_2parametres(1) - nouvellePosition(1)), ",", ".") & """ " & pointilles & _
                        " style=""fill:none;stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & ";stroke:" & couleur & """  />" & vbNewLine
                    Else
                         ecritureStream.WriteText "<line x1=""" & Replace(nouvellePosition(0), ",", ".") & _
                        """ y1=""" & Replace(nouvellePosition(1), ",", ".") & _
                        """ x2=""" & Replace(tab_2parametres(0), ",", ".") & _
                        """ y2=""" & Replace(tab_2parametres(1), ",", ".") & """ " & pointilles & _
                        " style=""fill:none;stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & ";stroke:" & couleur & """  />" & vbNewLine
                    End If
                 ElseIf commande = "ER" Then
                    Call rempliTableau(i, ligne, tab_2parametres)
                    Call changeRepere(tab_2parametres, inputP, echelle, True)
                    
                    tab_2parametres(0) = nouvellePosition(0) + tab_2parametres(0)
                    tab_2parametres(1) = nouvellePosition(1) + tab_2parametres(1)
                    
                    If tab_2parametres(0) - nouvellePosition(0) <> 0 And tab_2parametres(1) - nouvellePosition(1) <> 0 Then
                        ecritureStream.WriteText "<rect x=""" & Replace(Min(nouvellePosition(0), tab_2parametres(0)), ",", ".") & _
                        """ y=""" & Replace(Min(nouvellePosition(1), tab_2parametres(1)), ",", ".") & _
                        """ width=""" & Replace(Abs(tab_2parametres(0) - nouvellePosition(0)), ",", ".") & _
                        """ height=""" & Replace(Abs(tab_2parametres(1) - nouvellePosition(1)), ",", ".") & """ " & pointilles & _
                        " style=""fill:none;stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & ";stroke:" & couleur & """  />" & vbNewLine
                    Else
                         ecritureStream.WriteText "<line x1=""" & Replace(nouvellePosition(0), ",", ".") & _
                        """ y1=""" & Replace(nouvellePosition(1), ",", ".") & _
                        """ x2=""" & Replace(tab_2parametres(0), ",", ".") & _
                        """ y2=""" & Replace(tab_2parametres(1), ",", ".") & """ " & pointilles & _
                        " style=""fill:none;stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & ";stroke:" & couleur & """  />" & vbNewLine
                    End If               

                ' ---------------- ARCS --------------------
                ElseIf commande = "CI" Then
                
                    Call rempliTableau(i, ligne, tab_2parametres)
                    
                    'Call changeRepere(tab_2parametres, inputP, echelle, True)
                    ecritureStream.WriteText "<circle  cx=""" & Replace(anciennePosition(0), ",", ".") & """ cy=""" & Replace(anciennePosition(1), ",", ".") & """ " & _
                                                " r=""" & Replace(tab_2parametres(0), ",", ".") & """ " & _
                                                "style=""fill:none;stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & ";stroke:" & couleur & """  />" & vbNewLine
                    

                ElseIf commande = "AT" Then
                    Call rempliTableau(i, ligne, tab_4parametres)
                    anciennePosition(0) = nouvellePosition(0)
                    anciennePosition(1) = nouvellePosition(1)
                    
                    'conversion des points récupérés dans le plan
                    tab_2parametres(0) = tab_4parametres(0)
                    tab_2parametres(1) = tab_4parametres(1)
                    Call changeRepere(tab_2parametres, inputP, echelle, False)
                    nouvellePosition(0) = tab_4parametres(2)
                    nouvellePosition(1) = tab_4parametres(3)
                    Call changeRepere(nouvellePosition, inputP, echelle, False)
              
                    point_centre(0) = ((nouvellePosition(1) - tab_2parametres(1)) * (anciennePosition(0) * anciennePosition(0) - tab_2parametres(0) * tab_2parametres(0) + anciennePosition(1) * anciennePosition(1) - tab_2parametres(1) * tab_2parametres(1)) + (anciennePosition(1) - tab_2parametres(1)) * (tab_2parametres(0) * tab_2parametres(0) - nouvellePosition(0) * nouvellePosition(0) + tab_2parametres(1) * tab_2parametres(1) - nouvellePosition(1) * nouvellePosition(1))) / (2 * ((tab_2parametres(0) - nouvellePosition(0)) * (anciennePosition(1) - tab_2parametres(1)) - (nouvellePosition(1) - tab_2parametres(1)) * (tab_2parametres(0) - anciennePosition(0))))
                    point_centre(1) = (2 * (tab_2parametres(0) - anciennePosition(0)) * point_centre(0) + anciennePosition(0) * anciennePosition(0) - tab_2parametres(0) * tab_2parametres(0) + anciennePosition(1) * anciennePosition(1) - tab_2parametres(1) * tab_2parametres(1)) / (2 * (anciennePosition(1) - tab_2parametres(1)))
                    
                    'point_centre = calcul_centre_cercle(anciennePosition, tab_2parametres, nouvellePosition)
                    
                    rayon = Sqr((anciennePosition(0) - point_centre(0)) * (anciennePosition(0) - point_centre(0)) + (anciennePosition(1) - point_centre(1)) * (anciennePosition(1) - point_centre(1)))
                    
                    'Vérifie si le 2eme point est entre le point 1 et 3
                    If (((anciennePosition(0) - point_centre(0)) * (tab_2parametres(0) - point_centre(0)) + (anciennePosition(1) - point_centre(1)) * (tab_2parametres(1) - point_centre(1))) / (Sqr((anciennePosition(0) - point_centre(0)) * (anciennePosition(0) - point_centre(0)) + (anciennePosition(1) - point_centre(1)) * (anciennePosition(1) - point_centre(1))) * Sqr((tab_2parametres(0) - point_centre(0)) * (tab_2parametres(0) - point_centre(0)) + (tab_2parametres(1) - point_centre(1)) * (tab_2parametres(1) - point_centre(1))))) _
                        > (((anciennePosition(0) - point_centre(0)) * (nouvellePosition(0) - point_centre(0)) + (anciennePosition(1) - point_centre(1)) * (nouvellePosition(1) - point_centre(1))) / (Sqr((anciennePosition(0) - point_centre(0)) * (anciennePosition(0) - point_centre(0)) + (anciennePosition(1) - point_centre(1)) * (anciennePosition(1) - point_centre(1))) * Sqr((nouvellePosition(0) - point_centre(0)) * (nouvellePosition(0) - point_centre(0)) + (nouvellePosition(1) - point_centre(1)) * (nouvellePosition(1) - point_centre(1))))) Then
                        grandArc = 0
                    Else
                        grandArc = 1
                    End If
                    
                    'Calcul du produit vectoriel pour vérifier l'orientation du cercle
                    If ((anciennePosition(0) - point_centre(0)) * (tab_2parametres(1) - point_centre(1)) - (anciennePosition(1) - point_centre(1)) * (tab_2parametres(0) - point_centre(0))) > 0 Then
                        horaire = 1
                    Else
                        horaire = 0
                    End If
                    
                    
                    ecritureStream.WriteText "<path d =""M" & Replace(anciennePosition(0), ",", ".") & "," & Replace(anciennePosition(1), ",", ".") & _
                                                " A " & Replace(rayon, ",", ".") & " " & Replace(rayon, ",", ".") & " 0 " & Replace(grandArc, ",", ".") & " " & Replace(horaire, ",", ".") & _
                                                " " & Replace(nouvellePosition(0), ",", ".") & "," & Replace(nouvellePosition(1), ",", ".") & _
                                                """ " & pointilles & " style=""fill:none;stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & ";stroke:" & couleur & """  />" & vbNewLine


               ElseIf commande = "AA" Then
                    Call rempliTableau(i, ligne, tab_4parametres)
                    
                    anciennePosition(0) = nouvellePosition(0)
                    anciennePosition(1) = nouvellePosition(1)
                    Call changeRepereInverse(anciennePosition, inputP, echelle, False)
                    
                    
                    'conversion des points récupérés dans le plan
                    point_centre(0) = tab_4parametres(0)
                    point_centre(1) = tab_4parametres(1)

                    rayon_ellipse(0) = Sqr((anciennePosition(0) - point_centre(0)) * (anciennePosition(0) - point_centre(0)) + (anciennePosition(1) - point_centre(1)) * (anciennePosition(1) - point_centre(1)))
                    rayon_ellipse(1) = Sqr((anciennePosition(0) - point_centre(0)) * (anciennePosition(0) - point_centre(0)) + (anciennePosition(1) - point_centre(1)) * (anciennePosition(1) - point_centre(1)))
                    
                    'on se place dans un repère polaire de centre point_centre
                    'angle base est la coordonée d'angle polaire du point d'origine du trait
                    'angle est l'angle formé par les deux arcs, Pour calculer les coordonnées
                    'd'arrivée de l'arc, on converti dans le repère cartésiens le point (rayon, angle + angle_base)
                    '(je suis sûr qu'il y a plus simple...)
                    If (anciennePosition(0) - point_centre(0)) > 0 Then
                        angle_base = Atn((anciennePosition(1) - point_centre(1)) / (anciennePosition(0) - point_centre(0)))
                    ElseIf (anciennePosition(0) - point_centre(0)) < 0 Then
                        If (anciennePosition(1) - point_centre(1)) >= 0 Then
                             angle_base = Atn((anciennePosition(1) - point_centre(1)) / (anciennePosition(0) - point_centre(0))) + pi
                        Else
                            angle_base = Atn((anciennePosition(1) - point_centre(1)) / (anciennePosition(0) - point_centre(0))) - pi
                        End If
                    Else
                        If (anciennePosition(1) - point_centre(1)) > 0 Then
                            angle_base = pi / 2
                        ElseIf (anciennePosition(1) - point_centre(1)) < 0 Then
                            angle_base = -pi / 2
                        Else
                            angle_base = 0
                        End If
                    End If
                    
                    angle = tab_4parametres(2) * pi / 180
                    
                    nouvellePosition(0) = point_centre(0) + rayon_ellipse(0) * Cos(angle + angle_base)
                    nouvellePosition(1) = point_centre(1) + rayon_ellipse(0) * Sin(angle + angle_base)
                    Call changeRepere(nouvellePosition, inputP, echelle, False)
                    Call changeRepere(anciennePosition, inputP, echelle, False)
                    Call changeRepere(point_centre, inputP, echelle, False)

                    'ATTENTION !!!! hypothèse angle de l'ellipse parallèle aux axes du repères
                    rayon_ellipse(1) = Sqr(((nouvellePosition(1) - point_centre(1)) * (nouvellePosition(1) - point_centre(1)) * (anciennePosition(0) - point_centre(0)) * (anciennePosition(0) - point_centre(0)) - (anciennePosition(1) - point_centre(1)) * (anciennePosition(1) - point_centre(1)) * (nouvellePosition(0) - point_centre(0)) * (nouvellePosition(0) - point_centre(0)) ) / ( (anciennePosition(0) - point_centre(0)) * (anciennePosition(0) - point_centre(0)) - (nouvellePosition(0) - point_centre(0)) * (nouvellePosition(0) - point_centre(0))))

                    rayon_ellipse(0) = Sqr((rayon_ellipse(1) * rayon_ellipse(1) * (nouvellePosition(0) - point_centre(0)) * (nouvellePosition(0) - point_centre(0))) / ( rayon_ellipse(1) * rayon_ellipse(1) - (nouvellePosition(1) - point_centre(1)) * (nouvellePosition(1) - point_centre(1))) )

                    'Call changeRepere(rayon_ellipse, inputP, echelle, True)

                    'Vérifie si le 2eme point est entre le point 1 et 3
                    If Abs(angle) <= pi Then
                        grandArc = 0
                    Else
                        grandArc = 1
                    End If
                    
                    'Calcul du produit vectoriel pour vérifier l'orientation du cercle
                    If angle <= 0 Then
                        horaire = 1
                    Else
                        horaire = 0
                    End If
                    
                    ecritureStream.WriteText "<path d =""M" & Replace(anciennePosition(0), ",", ".") & "," & Replace(anciennePosition(1), ",", ".") & _
                                                " A " & Replace(rayon_ellipse(0), ",", ".") & " " & Replace(rayon_ellipse(1), ",", ".") & " 0 " & Replace(grandArc, ",", ".") & " " & Replace(horaire, ",", ".") & _
                                                " " & Replace(nouvellePosition(0), ",", ".") & "," & Replace(nouvellePosition(1), ",", ".") & _
                                                """ " & pointilles & " style=""fill:none;stroke-linecap:round;stroke-width:" & Replace(largeurTrait, ",", ".") & ";stroke:" & couleur & """ />" & vbNewLine

                'Styles de ligne
               ElseIf commande = "SP" Then
                    'penDown = False
                    Call rempliTableau(i, ligne, tab_1parametres)
                    Select Case tab_1parametres(0)
                        Case 1
                            couleur = "black"
                        Case 2
                            couleur = "red"
                        Case 3
                            couleur = "green"
                        Case 4
                            couleur = "yellow"
                        Case 5
                            couleur = "blue"
                        Case 6
                            couleur = "magenta"
                        Case 7
                            couleur = "cyan"
                        Case Else
                            couleur = "white"
                    End Select
                ElseIf commande = "PW" Then
                    penDown = False
                    Call rempliTableau(i, ligne, tab_1parametres)
                    largeurTrait = tab_1parametres(0)
                ElseIf commande = "LT" Then
                    Call rempliTableau(i, ligne, tab_1parametres)
                    Select Case Abs(tab_1parametres(0))
                        Case 1
                            pointilles = "stroke-dasharray=""5,5"""
                        Case 2
                            pointilles = "stroke-dasharray=""2.5, 5"""
                        Case 3
                            pointilles = "stroke-dasharray=""4, 1, 0.1, 1"""
                        Case 4
                            pointilles = "stroke-dasharray=""8, 1, 0.1, 1"""
                        Case 5
                            pointilles = "stroke-dasharray=""7, 1, 1 ,1"""
                        Case 6
                            pointilles = "stroke-dasharray=""5, 1, 1, 1, 1, 1"""
                        Case 7
                            pointilles = "stroke-dasharray=""7, 1, 0.1, 1, 0.1, 1"""
                        Case 8
                            pointilles = "stroke-dasharray=""7, 1, 0.1, 1, 1, 1, 0.1, 1"""
                        Case Else
                            pointilles = ""
                    End Select

                '--------- TEXTE ----------
                ElseIf commande = "DT" Then
                    
                    i = i + 1
                    terminator = Mid(ligne, i, 1)
                    If Mid(ligne, i + 1, 1) = "," Then
                        i = i + 2
                        imprimeTerm = CInt(Mid(ligne, i, 1))
                    End If
                ElseIf commande = "LO" Then
                    
                    Call rempliTableau(i, ligne, labelOrigin)
                    
                    Select Case CInt(labelOrigin(0))
                        Case 1, 2, 3, 11, 12, 13
                            anchor = "start"
                        Case 4, 14, 5, 15, 6, 16
                            anchor = "middle"
                        Case 7, 8, 9, 17, 18, 19
                            anchor = "end"
                        Case Else
                            anchor = "start"
                    End Select
                    Select Case CInt(labelOrigin(0))
                        Case 1, 4, 7
                            baseLineShift = "text-top"
                        Case 11, 14, 17
                            baseLineShift = "text-after-edge"
                        Case 12, 2, 5, 15, 8, 18
                            baseLineShift = "middle"
                        Case 3, 6, 9
                            baseLineShift = "hanging"
                        Case 13, 16, 19
                            baseLineShift = "text-before-edge"
                        Case Else
                            baseLineShift = "text-top"
                    End Select
                ElseIf commande = "LB" Then
                    texte = ""
                    i = i + 1
                    While Mid(ligne, i, 1) <> terminator
                    
                        texte = texte & RemplaceCaractereSpeciaux(Mid(ligne, i, 1))
                        i = i + 1
                    Wend
                    If imprimeTerm = 0 Then
                         texte = texte & RemplaceCaractereSpeciaux(Mid(ligne, i, 1))
                    End If
                    ecritureStream.WriteText "<text xml:space=""preserve"" x =""" & Replace(nouvellePosition(0), ",", ".") & """ y=""" & Replace(nouvellePosition(1), ",", ".") & """" & _
                                                " transform=""rotate(" & angle & " " & Replace(nouvellePosition(0), ",", ".") & " " & Replace(nouvellePosition(1), ",", ".") & ")"" " & _
                                                " style =""font-family:ISOCPEUR;font-size:" & Replace(taillePolice, ",", ".") & ";dominant-baseline:" & baseLineShift & ";text-anchor:" & anchor & """>" & _
                                                texte & "</text>" & vbNewLine

                ElseIf commande = "DI" Then
                    Call rempliTableau(i, ligne, tab_2parametres)
                    'Transformation miroir pour s'adapter au svg
                    tab_2parametres(1) = -tab_2parametres(1)
                    angle = calcul_angle_trigo(tab_2parametres)
                ElseIf commande = "SI" Then
                    Call rempliTableau(i, ligne, tab_2parametres)
                    taillePolice = CStr(Replace(tab_2parametres(1) * 10 / 0.7, ",", "."))
                ElseIf commande = "SD" Then
                    Call rempliPolyligne(i, ligne, polyligne)
                    
                    For j = LBound(polyligne, 2) To UBound(polyligne, 2)
                        If polyligne(0, j) = 4 Then
                            taillePolice = CStr(polyligne(1, j) * 0.352778) & "mm"
                        End If
                    Next
                Else 
                    'MsgBox commande
                End If
            'Si le caractère est ESC
            ElseIf Mid(ligne, i, 1) = Chr(27) Then
                i = i + 1
                parametreRaster = Mid(ligne, i, 1)
                If parametreRaster = "E" Then
                    mode = "Raster"
                ElseIf parametreRaster = "%" Then
                     Call rempliTableau(i, ligne, tab_1parametres)
                     
                     
                     i = i + 1
                     terminaisonRaster = Mid(ligne, i, 1)
                     
                     If terminaisonRaster = "A" Then
                        mode = "Raster"
                        inputP(0) = 0
                        inputP(1) = 0
                        inputP(2) = 0
                        inputP(3) = 0
                        echelle(0) = 0
                        echelle(1) = 0
                        echelle(2) = 0
                        echelle(3) = 0
                    ElseIf terminaisonRaster = "B" Then
                        mode = "HPG-2L"
                    End If

                ElseIf parametreRaster = "&" Then
                
                ElseIf parametreRaster = "*" Then
                    i = i + 1
                    groupeRaster = Mid(ligne, i, 1)
                    
                    If groupeRaster = "p" Then
                        Call rempliTableau(i, ligne, tab_1parametres)
                        i = i + 1
                        terminaisonRaster = Mid(ligne, i, 1)
                        
                        If terminaisonRaster = "X" Then
                            anciennePosition(0) = nouvellePosition(0)
                            nouvellePosition(0) = 2.5 + tab_1parametres(0) * 25.4 / 300
                            
                        ElseIf terminaisonRaster = "Y" Then
                            anciennePosition(1) = nouvellePosition(1)
                            nouvellePosition(1) = 15 + tab_1parametres(0) * 25.4 / 300
                           
                            texte = ""
                            i = i + 1
                            While i <= Len(ligne) And Mid(ligne, i, 1) <> Chr(27)
                                texte = texte & RemplaceCaractereSpeciaux(Mid(ligne, i, 1))
                                i = i + 1
                            Wend

                            If Mid(ligne, i, 1) = Chr(27) Then 
                                i = i - 1 
                            End If

                            If texte <> "" Then
                                ecritureStream.WriteText "<text xml:space=""preserve"" x =""" & Replace(nouvellePosition(0), ",", ".") & """ y=""" & Replace(nouvellePosition(1), ",", ".") & """" & _
                                                " style =""font-family:Calibri;font-weight:bold;font-size:" & Replace(taillePolice, ",", ".") & ";dominant-baseline:Middle;text-anchor:start"">" & _
                                                texte & "</text>" & vbNewLine
                            End If
                        End If
                    End If
                ElseIf parametreRaster = "(" Then
                    i = i + 1
                    groupeRaster = Mid(ligne, i, 1)
                    If groupeRaster = "s" Then
                        
                        While Mid(ligne, i, 1) <> "p" And Not(isUpperCase(Mid(ligne, i, 1))) And i <= len(ligne)
                            i = i + 1
                        Wend
                        
                        Call rempliTableau(i, ligne, tab_1parametres)
                        
                        taillePolice = CStr(tab_1parametres(0) * 0.352778)
                       ' taillePolice = CStr(tab_1parametres(0)) & "mm"
                        
                    End If
                End If
            End If
        Next
    Wend
    lectureStream.Close

    ecritureStream.WriteText "</svg>" & vbNewLine
    ecritureStream.SaveToFile outputFile, 2
    ecritureStream.Close
    MsgBox "ecriture terminée"
End Sub 

Function IsLetter(strValue)
    Dim intPos 
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function

Function IsNumber(strValue)
    Dim intPos 
    For intPos = 1 To Len(strValue)
        Select Case strValue
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
                IsNumber = True
            Case Else
                IsNumber = False
                Exit For
        End Select
    Next
End Function

Function isUpperCase(strValue)

    isUpperCase = IsLetter(strValue) And (UCase(strValue) = strValue)

End Function
Function Min(a , b ) 

    If a < b Then
       Min = a
    Else
        Min = b
    End If
    
End Function

Function Max(a , b ) 

    If a > b Then
       Max = a
    Else
        Max = b
    End If
    
End Function
'Fonction qui rempli les paramètres d'un tableau donnée en HPG-2L
Sub rempliTableau(i, ligne, ByRef tableau)

    Dim compteur 
    Dim texte   
    Dim j      
    Dim nombreDemarre 
    
    'effacement du tableau
    For j = LBound(tableau) To UBound(tableau)
        tableau(j) = 0
    Next
    
    compteur = 0
    Do
        nombreDemarre = False
        Do
            i = i + 1
            texte = texte & Mid(ligne, i, 1)
            
            If IsNumber(Mid(ligne, i, 1)) Then
                nombreDemarre = True
            End If
                
        'Attention, la virgule n'est pas obligatoire, on peut directement avoir le signes + et -
        Loop While Mid(ligne, i, 1) <> "," And Mid(ligne, i, 1) <> ";" And Not (nombreDemarre And Mid(ligne, i, 1) = "-") And Not (nombreDemarre And Mid(ligne, i, 1) = "+") And Not (nombreDemarre And Mid(ligne, i, 1) = " ") And Not (IsLetter(Mid(ligne, i, 1))) And i < Len(ligne)
        If nombreDemarre Then
            tableau(compteur) = CDbl(Replace(Left(texte, Len(texte) - IIf(IsNumber(Right(texte, 1)), 0, 1)), ".", ","))
            texte = ""
            compteur = compteur + 1
            If Mid(ligne, i, 1) = "-" Then
                i = i - 1
            End If
        End If
            
    ' Attention, le point virgule n'est pas forcément obligatoire
    Loop While Mid(ligne, i, 1) <> ";" And compteur <= UBound(tableau) And Not (IsLetter(Mid(ligne, i, 1))) And i < Len(ligne)
    i = i - 1
        
End Sub

Sub changeRepere(ByRef coordonnees, ByRef inputP, ByRef echelle, relatif)
    
    Dim tableauVide(3)

    tableauVide(0) = 0
    tableauVide(1) = 0
    tableauVide(2) = 0 
    tableauVide(3) = 0
    
    'On test si une echelle a été définie
    If (inputP(0) = inputP(2)) And (inputP(1) = inputP(3)) Then
        
        'on convertie du repère unités HPG vers SVG
         coordonnees(0) = IIf(relatif, 0, 2.5) + coordonnees(0) * 0.025 'marge gauche
         coordonnees(1) = IIf(relatif, 0, 297 - 10) - coordonnees(1) * 0.025 'Hauteur Page - marge supérieure
         
    Else
        'conversion d'abord dans le repère des unités HPG
        coordonnees(0) = IIf(relatif, 0, Min(inputP(0), inputP(2))) + (coordonnees(0) - echelle(0)) * (Max(inputP(0), inputP(2)) - Min(inputP(0), inputP(2))) / (echelle(1) - echelle(0))
        coordonnees(1) = IIf(relatif, 0, Min(inputP(1), inputP(3))) + (coordonnees(1) - echelle(2)) * (Max(inputP(1), inputP(3)) - Min(inputP(1), inputP(3))) / (echelle(3) - echelle(2))
        Call changeRepere(coordonnees, tableauVide, tableauVide, relatif)
    End If

End Sub

Function rempliPolyligne(i, ligne, ByRef polyligne)

    Dim compteur
    Dim texte      
    Dim XouY       
    Dim nombreDemarre
    
    'effacement de la polyligne
    ReDim polyligne(1, 0)
    compteur = 0
    XouY = 0
    
    Do
        nombreDemarre = False
        Do
            i = i + 1
            texte = texte & Mid(ligne, i, 1)
            
            If IsNumber(Mid(ligne, i, 1)) Then
                nombreDemarre = True
            End If
                
        'Attention, la virgule n'est pas obligatoire, on peut directement avoir le signes + et -
        Loop While Mid(ligne, i, 1) <> "," And Mid(ligne, i, 1) <> ";" And Not (nombreDemarre And Mid(ligne, i, 1) = "-") And Not (nombreDemarre And Mid(ligne, i, 1) = "+") And Not (IsLetter(Mid(ligne, i, 1))) And i < Len(ligne)
        
        If nombreDemarre Then
            ReDim Preserve polyligne(1, compteur)
            polyligne(XouY, compteur) = CDbl(Replace(Left(texte, Len(texte) - IIf(IsNumber(Right(texte, 1)), 0, 1)), ".", ","))
            texte = ""
            
            If XouY <> 0 Then
                XouY = 0
                compteur = compteur + 1
            Else
                XouY = 1
            End If
            rempliPolyligne = True
            If Mid(ligne, i, 1) = "-" Then
                i = i - 1
            End If
        Else
        
            i = i - 1
            rempliPolyligne = False
            Exit Function
        End If
    ' Attention, le point virgule n'est pas forcément obligatoire
    Loop While Mid(ligne, i + 1, 1) <> ";" And Not (IsLetter(Mid(ligne, i + 1, 1))) And i + 1 < Len(ligne)
    i = i - 1
        
End Function

Sub changeRepereInverse(ByRef coordonnees, ByRef inputP, ByRef echelle, relatif)
    
    Dim tableauVide(3)
    
    'On test si une echelle a été définie
    If (inputP(0) = inputP(2)) And (inputP(1) = inputP(3)) Then
        
        'on convertie du repère unités HPG vers SVG
         coordonnees(0) = (coordonnees(0) - IIf(relatif, 0, 2.5)) / 0.025 'marge gauche
         coordonnees(1) = (IIf(relatif, 0, 297 - 10) - coordonnees(1)) / 0.025 'Hauteur Page - marge supérieure
         
    Else
        'conversion d'abord dans le repère des unités HPG
        Call changeRepereInverse(coordonnees, tableauVide, tableauVide, relatif)
        coordonnees(0) = (echelle(1) - echelle(0)) / (Max(inputP(0), inputP(2)) - Min(inputP(0), inputP(2))) * (coordonnees(0) - IIf(relatif, 0, Min(inputP(0), inputP(2)))) + echelle(0)
        'coordonnees(1) = (echelle(3) - echelle(2)) / (Max(inputP(1), inputP(3)) - Min(inputP(1), inputP(3))) * (coordonnees(1) - IIf(relatif, 0, Min(inputP(1), inputP(3)))) + echelle(2)
        coordonnees(1) = (coordonnees(1) - IIf(relatif, 0, Min(inputP(1), inputP(3)))) * ((echelle(3) - echelle(2)) / (Max(inputP(1), inputP(3)) - Min(inputP(1), inputP(3)))) + echelle(2)
        
        
    End If

End Sub


Function RemplaceCaractereSpeciaux(caractere)

Select Case caractere
    Case Chr(128)  
        RemplaceCaractereSpeciaux = "Ç"
    Case Chr(129)  
        RemplaceCaractereSpeciaux = "ü"
    Case Chr(130)  
        RemplaceCaractereSpeciaux = "é"
    Case Chr(131)  
        RemplaceCaractereSpeciaux = "â"
    Case Chr(132)  
        RemplaceCaractereSpeciaux = "Â"
    Case Chr(133)  
        RemplaceCaractereSpeciaux = "à"
    Case Chr(135)  
        RemplaceCaractereSpeciaux = "ç"
    Case Chr(136)  
        RemplaceCaractereSpeciaux = "ê"
    Case Chr(137)  
        RemplaceCaractereSpeciaux = "ë"
    Case Chr(138)  
        RemplaceCaractereSpeciaux = "è"
    Case Chr(139)  
        RemplaceCaractereSpeciaux = "ï"
    Case Chr(140)  
        RemplaceCaractereSpeciaux = "î"
    Case Chr(142)  
        RemplaceCaractereSpeciaux = "À"
    Case Chr(144)  
        RemplaceCaractereSpeciaux = "É"
    Case Chr(145)  
        RemplaceCaractereSpeciaux = "È"
    Case Chr(146)  
        RemplaceCaractereSpeciaux = "Ê"
    Case Chr(147)  
        RemplaceCaractereSpeciaux = "ô"
    Case Chr(148)  
        RemplaceCaractereSpeciaux = "Ë"
    Case Chr(149)  
        RemplaceCaractereSpeciaux = "Ï"
    Case Chr(150)  
        RemplaceCaractereSpeciaux = "û"
    Case Chr(151)  
        RemplaceCaractereSpeciaux = "ù"
    Case Chr(153)  
        RemplaceCaractereSpeciaux = "Ô"
    Case Chr(154)  
        RemplaceCaractereSpeciaux = "Ü"
    Case Chr(155)  
        RemplaceCaractereSpeciaux = "¢"
    Case Chr(156)  
        RemplaceCaractereSpeciaux = "£"
    Case Chr(157)  
        RemplaceCaractereSpeciaux = "Ù"
    Case Chr(158)  
        RemplaceCaractereSpeciaux = "Û"
    Case Chr(159)  
        RemplaceCaractereSpeciaux = "¿"
    Case "³" 
        RemplaceCaractereSpeciaux = "¦"
    Case "í"
        RemplaceCaractereSpeciaux = "Ø"
    Case Chr(248)  
        RemplaceCaractereSpeciaux = "°"
    Case Chr(0)
        RemplaceCaractereSpeciaux = ""
    Case Else
        RemplaceCaractereSpeciaux = caractere
    End Select
    
End Function

Function calcul_angle_trigo(tableau)
'tableau(0) = cos alpha
'tableau(1) = sin alpha

    Dim pi
    pi = 3.14159265 

    If tableau(0) > 0 Then
        
        calcul_angle_trigo = Atn(tableau(1) / tableau(0)) * 180 / pi
    ElseIf tableau(0) < 0 Then
        calcul_angle_trigo = Atn(tableau(1) / tableau(0)) * 180 / pi + IIf(tableau(1) < 0, -90, 90)
    Else
        If tableau(1) < 0 Then
            calcul_angle_trigo = -90
        ElseIf tableau(1) > 0 Then
            calcul_angle_trigo = 90
        Else
            calcul_angle_trigo = 0
        End If
    End If


End Function


Function IIf( expr, truepart, falsepart )
   IIf = falsepart
   If expr Then IIf = truepart
End Function


