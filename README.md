# Portfolio Risk Analysis and Hedging with European Put Options (Excel VBA + Monte Carlo)

This project performs a complete risk analysis of a 1,000,000 $ portfolio replicating the S&P 500 index over a 10-day horizon. The goal is to measure portfolio risk and evaluate the effect of hedging with European put options. All computations ( analytical Value-at-Risk (VaR), Monte Carlo simulation, and Black-Scholes pricing ) are implemented in **Excel VBA**.

---

## Objectives

- Compute analytical Value-at-Risk (VaR) for an unhedged SP500 portfolio.
- Compute simulated VaR using Monte Carlo simulation.
- Price European put options using the Black-Scholes model.
- Compare hedged vs unhedged portfolios.
- Evaluate the impact of hedging on:
  - expected profit  
  - volatility  
  - 99% VaR (tail-risk)  
  - minimum value preserved  
- Recommend a strike price and hedge amount.

---

## Key Concepts

### 1. Daily Returns (Simple Returns)

Daily return:
```
r_t = (P_t - P_(t-1)) / P_(t-1)
```

Used to estimate:
- average daily return `mu`)
- daily volatility `sigma`)

Assume normal distribution of returns.

---

### 2. Analytical Value-at-Risk (VaR)

Multi-day scaling:
```
Mean_Profit = mu * T * invested_amount
StdDev_Profit = sigma * sqrt(T) * invested_amount
```

Analytical VaR at confidence `alpha`:
```
VaR(alpha) = - ( Mean_Profit + Z(alpha) * StdDev_Profit )
```

Where:
- `Z(alpha)` = standard normal quantile (ex: Z(99%) = 2.33)
- VaR is expressed as a positive loss

Interpretation:
> With probability (1 − alpha), the portfolio will not lose more than the VaR amount.

---

### 3. Monte Carlo Simulation

Simulated future SP500 price:
```
S_T = S_0 + mu * S_0 * T + sigma * S_0 * sqrt(T) * RandomNormal()
```

Portfolio components:
```
Underlying Profit = (S_T - S_0) / S_0 * montant_S
Put Payoff = max(K - S_T, 0)
Total Profit = Underlying Profit + nbOptions * Put Payoff - montant_P
```

Sorting simulated profits → extract empirical VaR.

---

### 4. Black-Scholes Pricing (European Put)

Inputs: S0, K, r, T, sigma
```
d1 = [ ln(S0/K) + (r + 0.5sigma^2)T ] / (sigma*sqrt(T))
d2 = d1 - sigma*sqrt(T)
```

Put price:
```
PutPrice = K * exp(-r*T) * N(-d2) - S0 * N(-d1)
```

Number of puts purchased:
```
nbOptions = montant_P / PutPrice
```

---

## VBA Implementation

### 1. VaR.bas — Analytical VaR
```vb
Function MoyenEcartVaR(vecPrix As Range, MontantInvesti As Double, alpha As Double, nbJours As Integer)
```

Features:
- Computes daily returns
- Estimates mean & volatility
- Scales them to multi-day horizon
- Computes analytical VaR

Returns: mean profit, standard deviation, VaR

---

### 2. VaR_hedged.bas — Monte Carlo VaR with Put Options
```vb
Function MoyenEcartVaR_avecOptionVente( _
    S0 As Double, montant_S As Double, montant_P As Double, K As Double, _
    r As Double, mu As Double, sigma As Double, alpha As Double, _
    nbJours As Double, nbSim As Long)
```

Features:
- Prices puts with Black-Scholes
- Simulates SP500 future prices
- Computes hedged portfolio profit
- Extracts simulated mean, volatility, and VaR

---

### 3. Bl_Sch_Benninga.bas — Black-Scholes Module

Contains:
- Normal CDF
- Computation of d1 and d2
- European put pricing formula

Used by the hedged VaR module
