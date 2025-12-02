Attribute VB_Name = "VaR_Hedged"
Option Explicit

Function MoyenEcartVaR_avecOptionVente( _
    S_0 As Double, montant_S As Double, montant_P As Double, K As Double, _
    interest As Double, mu As Double, sigma As Double, alpha As Double, _
    nbJours As Double, nbSim As Long)

    ' Fonction qui simule, pour une période de nbJours :
    ' - la moyenne du profit,
    ' - l'écart-type du profit,
    ' - la VaR au niveau alpha de la perte
    ' pour un portefeuille composé :
    '   - d'un montant_S $ investi dans l'actif S (S&P 500),
    '   - d'un montant_P $ investi dans une option de vente européenne sur S.
    '
    ' Variables d'entrée :
    '   S_0        = le prix de l'actif aujourd'hui (au temps zéro)
    '   montant_S  = le montant investi dans l'actif
    '   montant_P  = le montant investi dans l'option de vente
    '   K          = prix d'exercice de l'option de vente
    '   interest   = taux d'intérêt sans risque journalier
    '   mu         = moyenne des rendements journaliers de l'actif S
    '   sigma      = écart-type des rendements journaliers de l'actif S
    '   alpha      = niveau pour le VaR (entre 0 et 1, ex. 0.05)
    '   nbJours    = horizon en nombre de jours
    '   nbSim      = nombre de simulations Monte Carlo

    ' Nombre d'options achetées avec le montant_P :
    Dim nbOptions As Double
    nbOptions = montant_P / BSPut(S_0, K, nbJours, interest, sigma)

    Dim ProfitMoyen As Double
    Dim ProfitVariance As Double
    ProfitMoyen = 0
    ProfitVariance = 0

    ' Tableaux de profits et pertes simulés :
    Dim Profits_sim() As Double
    Dim Pertes_sim() As Double
    ReDim Profits_sim(1 To nbSim, 1 To 1)
    ReDim Pertes_sim(1 To nbSim, 1 To 1)

    Dim S_futur As Double
    Dim i As Long

    For i = 1 To nbSim
        ' Prix futur simulé (rendement normal agrégé sur nbJours)
        S_futur = Application.Norm_S_Inv(Rnd()) * sigma * S_0 * Sqr(nbJours) _
                  + S_0 + mu * S_0 * nbJours

        ' Profit du portefeuille : SP500 + payoff put - coût initial des puts
        Profits_sim(i, 1) = (S_futur - S_0) / S_0 * montant_S _
                            + nbOptions * ValeurFinalOptionVente(S_futur, K) _
                            - montant_P

        ' Perte = - profit
        Pertes_sim(i, 1) = -Profits_sim(i, 1)

        ' Accumulation pour moyenne et variance
        ProfitMoyen = ProfitMoyen + Profits_sim(i, 1)
        ProfitVariance = ProfitVariance + Profits_sim(i, 1) * Profits_sim(i, 1)
    Next i

    ' Moyenne et variance des profits
    ProfitMoyen = ProfitMoyen / nbSim
    ProfitVariance = ProfitVariance / nbSim - ProfitMoyen * ProfitMoyen

    ' Calcul de la VaR sur la distribution des pertes
    Dim indexVaR As Long
    Dim VaR As Double

    indexVaR = Application.WorksheetFunction.RoundUp((1 - alpha) * nbSim, 0)
    If indexVaR < 1 Then indexVaR = 1
    If indexVaR > nbSim Then indexVaR = nbSim

    ' On prend le (1 - alpha)-quantile des pertes (Small sur les pertes)
    VaR = Application.WorksheetFunction.Small(Pertes_sim, indexVaR)
    ' VaR est déjà positive (perte)

    ' Vecteur de sortie 3x1 :
    Dim vecSortie(1 To 3, 1 To 1) As Double
    vecSortie(1, 1) = ProfitMoyen
    vecSortie(2, 1) = Sqr(ProfitVariance)
    vecSortie(3, 1) = VaR

    MoyenEcartVaR_avecOptionVente = vecSortie

End Function


Function ValeurFinalOptionVente(S_final As Double, exercice As Double) As Double
    ' Valeur d'une option de vente à la maturité
    ValeurFinalOptionVente = Application.WorksheetFunction.Max(0, exercice - S_final)
End Function


