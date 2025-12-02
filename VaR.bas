Attribute VB_Name = "VaR"
Option Explicit

Function MoyenEcartVaR(vecIndicePrix As Range, MontantInvesti As Double, alpha As Double, nbJours As Integer)

    ' Vérification des entrées
    
    If vecIndicePrix.Columns.Count > 1 Then
        MoyenEcartVaR = "Erreur: plage doit être une seule colonne"
        Exit Function          ' le vecteur doit être une seule colonne
    End If
    
    If MontantInvesti <= 0 Then
        MoyenEcartVaR = "Erreur: montant investi doit être > 0"
        Exit Function
    End If
    
    If alpha <= 0 Or alpha >= 1 Then
        MoyenEcartVaR = "Erreur: alpha doit être entre 0 et 1"
        Exit Function
    End If
    
    If nbJours <= 0 Then
        MoyenEcartVaR = "Erreur: nbJours doit être >= 1"
        Exit Function
    End If
    
    Dim prix() As Variant
    Dim rendements() As Double
    Dim n As Long
    Dim i As Long
    
    prix = vecIndicePrix.Value
    n = vecIndicePrix.Rows.Count          ' nombre de prix
    
    ' Il faut au minimum 2 prix pour avoir 1 rendement
    If n < 2 Then
        MoyenEcartVaR = "Erreur: au moins 2 observations de prix"
        Exit Function
    End If
    
    ' On a n-1 rendements pour n prix
    ReDim rendements(1 To n - 1)
    
    For i = 1 To n - 1
        rendements(i) = (prix(i + 1, 1) - prix(i, 1)) / prix(i, 1)
    Next i
    
    Dim moyenneRend As Double
    Dim ecartTypeRend As Double
    Dim somme As Double
    Dim sommeCarres As Double
    
    somme = 0
    sommeCarres = 0
    
    ' On somme les rendements (il y en a n-1)
    For i = 1 To n - 1
        somme = somme + rendements(i)
        sommeCarres = sommeCarres + rendements(i) ^ 2
    Next i
    
    moyenneRend = somme / (n - 1)
    ecartTypeRend = Sqr((sommeCarres / (n - 1)) - (moyenneRend ^ 2))
    
    ' Calcul pour nbJours (profit sur l’horizon)
    Dim moyenneProf As Double
    Dim ecartTypeProf As Double
    Dim VaR As Double
    Dim z As Double
    
    ' Profit ˜ MontantInvesti * (nbJours * moyenneRend) avec variance nbJours * sigma^2
    moyenneProf = moyenneRend * MontantInvesti * nbJours
    ecartTypeProf = MontantInvesti * Sqr(nbJours) * ecartTypeRend
    
    ' VaR au seuil alpha de la distribution des pertes
    ' On travaille sur les PROFITS P ~ N(moyenneProf, ecartTypeProf^2)
    ' VaR(alpha) = - quantile_alpha(P)
    ' quantile_alpha(P) = moyenneProf + z_alpha * ecartTypeProf
    ' avec z_alpha = Norm_S_Inv(alpha) (négatif si alpha < 0.5)
    
    z = Application.WorksheetFunction.Norm_S_Inv(alpha)
    VaR = -(moyenneProf + z * ecartTypeProf)   ' VaR > 0, perte
    
    Dim result(1 To 3, 1 To 1) As Variant
    result(1, 1) = moyenneProf
    result(2, 1) = ecartTypeProf
    result(3, 1) = VaR
    
    MoyenEcartVaR = result

End Function





