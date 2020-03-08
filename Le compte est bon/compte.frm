VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} compte 
   Caption         =   "Le compte est bon par Hugues DUMONT"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8265
   OleObjectBlob   =   "compte.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "compte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'''Tableau contenant les possibilités du tirage des valeurs 2 * 1 à 10 + 2*(25;50;75;100)
Private possibilites(28) As Integer

'''Contient le dernier opérateur enregistré
Private operateur As String

'''Tableau contenant l'ensemble des opérations
Private operations As Collection

'''Variables nécessaires à la recherche de solutions
Private over As Boolean '''Indique si oui ou non la recherche de solutions est terminté
Private cible As Long '''nombre à atteindre
Private OperationsEnCours As Collection '''opérations effectuées pour arriver au résultat en cours
Private bestOperations As Collection '''opérations de la meilleure solution trouvée
Private bestDistance As Long '''distance entre la meilleure solution et l'objectif
Private bestNbOp As Long '''nombre d'opérations pour arriver à la meilleure solution
Private nbresAleas As Collection '''valeurs tirées au sort
Private operateurs As Collection '''Les opérateurs arithmiques + - * /
Private compteur As Long '''Nombre total d'opérations effectuées

'''Variable contenant le moment du début de la partie
Private debut As Single
'''Booléen pour arrêter la partie avant la fin du temps
Private finir As Boolean
'''Booléen pour n'afficher qu'une fois le message de fin partie et que la recherche est toujours en cours
Private displayOnce As Boolean

Private Sub UserForm_Initialize()
    Dim i As Integer

    For i = 1 To 6
        Me.Controls.Item("v" & i).Caption = ""
    Next i
    Me.obj.Value = ""
    Me.chrono.Caption = "40"
    
    Me.newGame.Enabled = True
    Me.EffDernier.Enabled = True
    Me.EffTout.Enabled = True
    If Not (operations Is Nothing) Then Call EffTout_Click
End Sub

Private Sub newGame_Click()
    Dim i As Integer
    'Dim debut As Single
    Dim usf As UserForm
    
    '''Initialisation des différents choix possibles pour le tirage
    For i = 1 To 10
        possibilites(i) = i
    Next i
    possibilites(11) = 25
    possibilites(12) = 50
    possibilites(13) = 75
    possibilites(14) = 100
    For i = 15 To 24
        possibilites(i) = i - 14
    Next i
    possibilites(25) = 25
    possibilites(26) = 50
    possibilites(27) = 75
    possibilites(28) = 100
    '''
    
    Set usf = finPartie
    Set operations = New Collection
    
    Call tirage
    
    Me.newGame.Enabled = False
    displayOnce = False
    debut = Timer
    finir = False
    Call ChargeSolution(usf)
    
    While Not over
        DoEvents
    Wend
    While Timer - debut < 40 And Not finir
        Me.chrono.Caption = CStr(Int(40 - Round((Timer - debut), 0)))
        DoEvents
    Wend
    If Not displayOnce Then
        MsgBox "La recherche de solutions n'est pas encore terminée." & _
            Chr(13) & "Merci de patient pendant ce temps.", vbOKOnly + vbInformation, _
            "Recherche de solutions en cours."
    End If
    
    Me.EffDernier.Enabled = False
    Me.EffTout.Enabled = False
    
    Call affichageJoueur(usf)
    finPartie.Show
    Set usf = Nothing
    Call UserForm_Initialize
End Sub

Private Sub ChargeSolution(ByRef usf As UserForm)
    Dim i As Integer
        
    For i = 1 To 6
        usf.Controls.Item("v" & i).Caption = Me.Controls.Item("v" & i).Caption
    Next i
    usf.Objectif.Caption = Me.obj.Value
    Call rechercheSolutions(usf)
    displayOnce = True
End Sub

Private Sub finChrono_Click()
    finir = True
End Sub

Private Sub tirage()
    Dim i As Integer, val As Integer, cmpt As Integer, j As Integer
    Randomize
    
    '''tirage aléatoire des nombres (en faisant attention à ce qu'un nombre n'apparaisse pas plus de 2 fois)
    i = 1
    While i < 7 '''Remplir les 6 valeurs (Hors objectif)
        j = 1
        val = Int(Rnd * 28) + 1 '''Tirer une nombre aléatoire parmis ceux possibles
        cmpt = 0
        While j < i '''Compter le nombre d'apparitions de ce nombre dans les
            If possibilites(val) = CInt(Me.Controls.Item("v" & j).Caption) Then
                cmpt = cmpt + 1
            End If
            j = j + 1
        Wend
        '''Si le nombre n'apparait pas encore 2 fois alors il est validé, sinon, on tire à nouveau un nombre
        If cmpt < 2 Then
            Me.Controls.Item("v" & i).Caption = CStr(possibilites(val))
            i = i + 1
        End If
    Wend
    
    '''Tirage aléatoire de l'objectif à atteindre
    Me.obj.Value = Int(Rnd * 899) + 100
    Me.obj.Locked = True
End Sub

Private Sub EffDernier_Click()
    If operations.Count > 0 Then
        With operations.Item(operations.Count)
            If .Nbr1 <> "" Then
                Me.Controls.Item(.Nbr1).Enabled = True
                Me.Controls.Item(.Nbr1).BackColor = RGB(255, 255, 255)
            End If
            If .Nbr2 <> "" Then
                Me.Controls.Item(.Nbr2).Enabled = True
                Me.Controls.Item(.Nbr2).BackColor = RGB(255, 255, 255)
            End If
            With Me.Controls
                .Item("calcul" & operations.Count).Caption = ""
                .Item("res" & operations.Count).Caption = ""
            End With
            operations.Remove operations.Count
        End With
    End If
End Sub

Private Sub EffTout_Click()
    While operations.Count > 0
        With operations.Item(operations.Count)
            If .Nbr1 <> "" Then
                Me.Controls.Item(.Nbr1).Enabled = True
                Me.Controls.Item(.Nbr1).BackColor = RGB(255, 255, 255)
            End If
            If .Nbr2 <> "" Then
                Me.Controls.Item(.Nbr2).Enabled = True
                Me.Controls.Item(.Nbr2).BackColor = RGB(255, 255, 255)
            End If
            With Me.Controls
                .Item("calcul" & operations.Count).Caption = ""
                .Item("res" & operations.Count).Enabled = True
                .Item("res" & operations.Count).Caption = ""
            End With
        End With
        operations.Remove operations.Count
    Wend
End Sub

Private Sub Quitter_click()
    Call userform_QueryClose(False, 1)
End Sub

Private Sub userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim vbRes As VbMsgBoxResult
    vbRes = MsgBox("Voulez-vous vraiment quitter ?", vbYesNo + vbQuestion, "Quitter ?")
    If vbRes = vbYes Then
        ThisWorkbook.Save
        Application.Quit
        End
    Else
        Cancel = True
    End If
End Sub

Private Sub operateurPlus_Click()
    Call majOperateur("+")
End Sub

Private Sub operateurMoins_Click()
    Call majOperateur("-")
End Sub

Private Sub operateurMult_Click()
    Call majOperateur("X")
End Sub

Private Sub operateurDiv_Click()
    Call majOperateur("/")
End Sub

Private Sub majOperateur(op As String)
    If operations.Count > 0 Then
        With operations.Item(operations.Count)
            If .Nbr2 = "" And .Nbr1 <> "" Then
                .optr = op
                With Me.Controls.Item("calcul" & operations.Count)
                    .Caption = .Caption & " " & operations.Item(operations.Count).optr
                End With
            End If
        End With
    End If
End Sub

Private Sub v1_Click()
    Call majNombre(Me.v1.Name)
End Sub

Private Sub v2_Click()
    Call majNombre(Me.v2.Name)
End Sub

Private Sub v3_Click()
    Call majNombre(Me.v3.Name)
End Sub

Private Sub v4_Click()
    Call majNombre(Me.v4.Name)
End Sub

Private Sub v5_Click()
    Call majNombre(Me.v5.Name)
End Sub

Private Sub v6_Click()
    Call majNombre(Me.v6.Name)
End Sub

Private Sub res1_Click()
    Call majNombre(Me.res1.Name)
End Sub

Private Sub res2_Click()
    Call majNombre(Me.res2.Name)
End Sub

Private Sub res3_Click()
    Call majNombre(Me.res3.Name)
End Sub

Private Sub res4_Click()
    Call majNombre(Me.res4.Name)
End Sub

'''Mettre à jour les opérations
Private Sub majNombre(val As String)
    Dim oper As New operation
    
    oper.Nbr1 = ""
    oper.Nbr2 = ""
    oper.optr = ""
    
    If operations.Count > 0 Then
        With operations.Item(operations.Count)
            If .Nbr2 <> "" Then
                oper.Nbr1 = val
                operations.Add oper
                With Me.Controls.Item("calcul" & operations.Count)
                    .Caption = Me.Controls.Item(operations.Item(operations.Count).Nbr1).Caption
                End With
            ElseIf .optr <> "" Then
                .Nbr2 = val
                If calculer Then
                    Me.Controls.Item(.Nbr1).Enabled = False
                    Me.Controls.Item(.Nbr1).BackColor = RGB(100, 100, 100)
                    Me.Controls.Item(.Nbr2).Enabled = False
                    Me.Controls.Item(.Nbr2).BackColor = RGB(100, 100, 100)
                    With Me.Controls.Item("calcul" & operations.Count)
                        .Caption = .Caption & " " & Me.Controls.Item(operations.Item(operations.Count).Nbr2).Caption
                    End With
                Else
                    .Nbr2 = ""
                End If
            End If
        End With
    Else
        oper.Nbr1 = val
        operations.Add oper
        With Me.Controls.Item("calcul" & operations.Count)
            .Caption = Me.Controls.Item(operations.Item(operations.Count).Nbr1).Caption
        End With
    End If
End Sub

'''Effectuer les calculs et véifier qu'ils sont autorisés
Private Function calculer() As Boolean
    Dim premier As Integer, deuxieme As Integer, resultat As Single
    With operations.Item(operations.Count)
        premier = CInt(Me.Controls.Item(.Nbr1).Caption)
        deuxieme = CInt(Me.Controls.Item(.Nbr2).Caption)
        If .optr = "+" Then
            resultat = premier + deuxieme
        ElseIf .optr = "-" Then
            resultat = premier - deuxieme
        ElseIf .optr = "X" Then
            resultat = premier * deuxieme
        Else
            resultat = premier / deuxieme
        End If
    End With
    
    If resultat < 0 Then
        MsgBox "Nombres négatifs non autorisés dans ce jeu !", vbOKOnly + vbCritical, "Résultat négatif"
        calculer = False
    ElseIf Int(resultat) <> resultat Then
        MsgBox "Division non entière !", vbOKOnly + vbCritical, "Division incorrect"
        calculer = False
    Else
        Me.Controls.Item("res" & operations.Count).Caption = CStr(Int(resultat))
        calculer = True
    End If
End Function

'''Recherche de solution
Private Sub rechercheSolutions(ByRef usf As UserForm)
    Dim i As Long, leMax As Long
    
    over = False

    ''''''Initialisation des variables
    compteur = 0 ''''''Nombre de parcours effectués lors de la recherche de solutions
    bestDistance = CLng(usf.Objectif.Caption) ''''''Meilleur résultat possible (objectif)
    cible = bestDistance ''''''Objectif à atteindre
    bestNbOp = 6 ''''''Nombre d'opérations mini à effectuer pour atteindre le résultat (le nombre de plaques initiales)
    
    ''''''Initialisation des collections
    Set nbresAleas = New Collection
    Set OperationsEnCours = New Collection
    Set bestOperations = New Collection
    Set operateurs = New Collection
    ''''''
    
    ''''''Récupération des valeurs des plaques
    For i = 1 To 6
        nbresAleas.Add CInt(compte.Controls.Item("v" & i).Caption)
    Next i
    ''''''
    
    Call initTablos ''''''Initialisation des opérandes
    Call aleasDecroissants ''''''Tri par ordre décroissants des plaques
    
    leMax = MAXI ''''''Récupération de la valeur maximale que l'on peut obtenir avec les plaques
    If leMax <= cible Then ''''''Si la valeur maximale est inférieure ou égale à l'objectif, alors c'est nécessairement le meilleur résultat possible
        compteur = 1 ''''''On a effectué un seul test
        If leMax = cible Then ''''''Si le compte est bon
            usf.SolutionJuste.Caption = "OUI"
        Else ''''''Sinon le compte n'est pas bon
            usf.SolutionJuste.Caption = "NON"
        End If
        usf.SolutionSolveur.Caption = CStr(leMax)
        usf.DistanceSolveur.Caption = CStr(CInt(usf.Objectif.Caption) - leMax)
    Else
        Call rechercheArbre(nbresAleas, usf) ''''''Recherche de la meilleure solution
        If bestDistance = 0 Then ''''''Si le plus petit écart de l'objectif est 0, alors on a trouvé au moins une solution exacte
            usf.SolutionJuste.Caption = "OUI"
        Else ''''''Sinon on n'a seulement une (ou des) valeurs approchantes
            usf.SolutionJuste.Caption = "NON"
        End If
    End If
    
    ''''''Affichage des opérations pour la meilleure solution trouvée
    Call affiche(compteur, bestOperations, usf)
    over = True
End Sub

Private Sub initTablos()
    Dim i As Long
    
    operateurs.Add "+"
    operateurs.Add "-"
    operateurs.Add "X"
    operateurs.Add "/"
End Sub

Private Sub aleasDecroissants()
    Dim i As Long, j As Long, tmp As Long
    Dim tablo(6) As Long
    
    For i = 6 To 1 Step -1
        tablo(i) = nbresAleas.Item(i)
        nbresAleas.Remove nbresAleas.Count
    Next i
    
    For i = 1 To 5
        For j = i + 1 To 6
            If tablo(j) > tablo(i) Then
                tmp = tablo(i)
                tablo(i) = tablo(j)
                tablo(j) = tmp
            End If
        Next j
    Next i
    
    For i = 1 To 6
        nbresAleas.Add tablo(i)
    Next i
End Sub

''''''Fonction pour déterminer la valeur maximale pouvant être obtenue avec l'ensemble des plaques
Private Function MAXI() As Long
    Dim i As Long, a As Long, b As Long, c As Long, d As Long, e As Long, f As Long
    
    If nbresAleas.Item(6) = 1 Then
        If nbresAleas.Item(5) = 1 Then
            ''''''Cas où 2 plaques valent 1 : additionner les plaquent valant 1 aux plus petites plaques disponibles (mais pas entre-elles)
            a = 1 + nbresAleas.Item(4) ''''''Addition du premier 1 avec la plus petite plaque ne valant pas 1
            bestOperations.Add "1 + " & nbresAleas.Item(4) & " = " & a
            If a < nbresAleas.Item(3) Then ''''''Cas où la plaque nouvellement créée est plus petite que les plaques déjà présentes, c'est avec elle qu'on additionne le deuxième 1
                b = 1 + a
                bestOperations.Add "1 + " & a & " = " & b
                c = b * nbresAleas.Item(3)
                bestOperations.Add b & " X " & nbresAleas.Item(3) & " = " & c
            Else
                b = 1 + nbresAleas.Item(3)
                bestOperations.Add "1 + " & nbresAleas.Item(3) & " = " & b
                c = a * b
                bestOperations.Add a & " X " & b & " = " & c
            End If
        Else
            ''''''Même cas, mais avec un seul 1, donc une seule addition à effectuer
            a = 1 + nbresAleas.Item(5)
            b = a * nbresAleas.Item(4)
            c = b * nbresAleas.Item(3)
            bestOperations.Add "1 + " & nbresAleas.Item(5) & " = " & a
            bestOperations.Add a & " X " & nbresAleas.Item(4) & " = " & b
            bestOperations.Add b & " X " & nbresAleas.Item(3) & " = " & c
        End If
    Else ''''''Cas sans 1, donc uniquement des multiplications consécutives
        a = nbresAleas.Item(6) * nbresAleas.Item(5)
        b = a * nbresAleas.Item(4)
        c = b * nbresAleas.Item(3)
        bestOperations.Add nbresAleas.Item(6) & " + " & nbresAleas.Item(5) & " = " & a
        bestOperations.Add a & " X " & nbresAleas.Item(4) & " = " & b
        bestOperations.Add b & " X " & nbresAleas.Item(3) & " = " & c
    End If
    
    ''''''Terminer les multiplications consécutives communes
    d = c * nbresAleas.Item(2)
    e = d * nbresAleas.Item(1)
    bestOperations.Add c & " X " & nbresAleas.Item(2) & " = " & d
    bestOperations.Add d & " X " & nbresAleas.Item(1) & " = " & e
    
    ''''''Retourner le maximum pouvant être obtenu avec l'ensemble des plaques
    MAXI = e
End Function

''''''Recherche de la meilleure solution possible
Private Sub rechercheArbre(tablo As Collection, ByRef usf As UserForm)
    Dim nbNombres As Long, res As Long
    Dim i As Long, j As Long, p As Long, k As Long
    Dim tab2 As Collection
    
    Set tab2 = New Collection ''''''Collection tampon pour les appels récursifs
    nbNombres = tablo.Count ''''''Nombre de plaques restantes à évaluer pour les opérations
    
    For i = 1 To nbNombres - 1
        For j = i + 1 To nbNombres
            For p = 1 To 4
                '''Chronomètre de la partie
                If Timer - debut < 40 And Not finir Then
                    Me.chrono.Caption = CStr(Int(40 - Round((Timer - debut), 0)))
                Else
                    If Not displayOnce Then
                        MsgBox "La recherche de solutions n'est pas encore terminée." & _
                            Chr(13) & "Merci de patient pendant ce temps.", vbOKOnly + vbInformation, _
                            "Recherche de solutions en cours."
                        displayOnce = True
                        Me.EffDernier.Enabled = False
                        Me.EffTout.Enabled = False
                        finir = True
                    End If
                End If
                ''''''Effectuer les opérations en associant chaque plaque
                res = calc(tablo.Item(i), tablo.Item(j), operateurs.Item(p)) ''''''Calcul entre deux plaques et un opérateur
                If res <> 0 Then ''''''Si le résultat des plaques est différent de 0, il est possible d'effectuer de nouvelles opérations avec la nouvelle plaque
                    ''''''Vider les plaques temporaires
                    For k = tab2.Count To 1 Step -1
                        tab2.Remove tab2.Count
                    Next k
                    ''''''
                    
                    ''''''Vérifier si la nouvelle plaque est la meilleure solution jusqu'ici
                    Call compare(Abs(res - cible), res, usf)
                    
                    ''''''ajouter la nouvelle plaque aux plaques temporaires
                    tab2.Add res
                    
                    ''''''Ainsi que les plaques non utilisées
                    For k = 1 To nbNombres
                        If k <> i And k <> j Then tab2.Add tablo.Item(k)
                        DoEvents
                    Next k
                    
                    DoEvents
                    If tab2.Count > 1 And OperationsEnCours.Count < bestNbOp - 1 Then Call rechercheArbre(tab2, usf)
                End If
                OperationsEnCours.Remove OperationsEnCours.Count
                DoEvents
            Next p
            DoEvents
        Next j
        DoEvents
    Next i
End Sub

Private Function calc(n1, n2, op) As Long
    Dim res As Long
    
    compteur = compteur + 1
    
    If op = "+" Then
        res = n1 + n2
    ElseIf op = "-" Then
        If n1 <= n2 Then
            res = n1
            n1 = n2
            n2 = res
        End If
        res = n1 - n2
    ElseIf op = "X" Then
        res = n1 * n2
    Else
        If n1 < n2 Then
            res = n1
            n1 = n2
            n2 = res
        End If
        res = n1 / n2
        If res <> Int(n1 / n2) Then res = 0
    End If
    
    OperationsEnCours.Add n1 & op & n2 & " = " & res
    calc = res
End Function

Private Sub compare(n As Long, r As Long, ByRef usf As UserForm)
    ''''''Si moins d'opérations ont été effectuées que pour le meilleur résultat précédent
    If OperationsEnCours.Count < bestNbOp Then
        If n = 0 Then ''''''alors si le compte est bon
            bestDistance = 0 ''''''L'écart avec l'objectif est 0
            bestNbOp = OperationsEnCours.Count ''''''On met à jour le nombre mini d'opérations à effectuer
            Call copieVersBestOperations ''''''On remplace les opérations précédentes par les nouvelles
            usf.SolutionSolveur.Caption = CStr(r)
            usf.DistanceSolveur.Caption = CStr(n)
        ElseIf bestDistance <> 0 And n < bestDistance And n <= 918 And r >= 81 Then ''''''Cas où l'on a pas encore trouvé le compte juste
            bestDistance = n ''''''Mise à jour de l'écart mini
            Call copieVersBestOperations ''''''Remplacement des opérations
            usf.SolutionSolveur.Caption = CStr(r)
            usf.DistanceSolveur.Caption = CStr(n)
        End If
    End If
End Sub

''''''Remplacement des précédentes opérations par les nouvelles pour le meilleur résultat possible
Private Sub copieVersBestOperations()
    Dim i As Long
    
    ''''''Suppression des précédentes meilleures opérations
    If bestOperations.Count > 0 Then
        For i = bestOperations.Count To 1 Step -1
            bestOperations.Remove bestOperations.Count
            DoEvents
        Next i
    End If
    
    ''''''Ajout des nouvelles
    For i = 1 To OperationsEnCours.Count
        bestOperations.Add OperationsEnCours.Item(i)
        DoEvents
    Next i
End Sub

''''''Affichage de la solution avec les différentes opérations
Private Sub affiche(nbOperations As Long, tabOperations As Collection, ByRef usf As UserForm)
    Dim i As Long
    
    For i = 1 To tabOperations.Count
        usf.Controls.Item("Ligne" & i).Caption = tabOperations.Item(i)
        DoEvents
    Next i
End Sub

Private Sub affichageJoueur(ByRef usf As UserForm)
    Dim i As Integer
    
    For i = 1 To 5
        If Me.Controls.Item("res" & i).Caption <> "" Then
            usf.ResultatJoueur.Caption = Me.Controls.Item("res" & i).Caption
        End If
    Next i
    
    If usf.ResultatJoueur.Caption = "" Then usf.ResultatJoueur.Caption = "0"
    
    If CInt(usf.ResultatJoueur.Caption) = CInt(usf.Objectif.Caption) Then
        usf.JusteJoueur.Caption = "OUI"
    Else
        usf.JusteJoueur.Caption = "NON"
    End If

    usf.DistanceJoueur.Caption = CStr(Abs(CInt(usf.Objectif.Caption) - CInt(usf.ResultatJoueur.Caption)))
    If usf.SolutionSolveur.Caption <> "" Then usf.DistanceMeilleur.Caption = CStr(Abs(CInt(usf.SolutionSolveur.Caption) - CInt(usf.ResultatJoueur.Caption)))
    
    For i = 1 To 5
        If Me.Controls.Item("calcul" & i).Caption <> "" Then
            usf.Controls.Item("Solut" & i).Caption = Me.Controls.Item("calcul" & i).Caption & _
                " = " & Me.Controls.Item("res" & i).Caption
        End If
    Next i
End Sub
