# Portfolio Risk Analysis and Hedging with European Put Options (Excel VBA + Monte Carlo)

This project implements a complete risk analysis of a 1,000,000$ portfolio replicating 
the S&P 500 index over a 10-day horizon. To protect the investor against large market losses,
the analysis evaluates the effect of purchasing European put options.

All computations (analytical VaR, Monte Carlo simulation, and Black-Scholes pricing) 
are implemented in **Excel VBA**, making the project fully reproducible and transparent.

---

##  Objectives of the project

- Compute **Value-at-Risk (VaR)** of an unhedged SP500 portfolio using an **analytical normal model**.
- Compute **simulated VaR** using **Monte Carlo simulation**.
- Price European **put options** using the **Black-Scholes** formula.
- Evaluate the effect of purchasing puts on:
  - expected profit,
  - portfolio volatility,
  - tail-risk (VaR at 99%),
  - minimum value preserved in extreme events.
- Provide a recommendation for a reasonable strike price and hedge amount.

This is a practical exercise in **market risk management**, **derivative hedging**, and **quantitative modelling**.

---

# **Key Concepts Used**

Below are the main financial and mathematical concepts implemented in this project.

---

##  Daily Returns (Simple Returns)

Given price series \( P_t \), daily simple returns are:

\[
r_t = \frac{P_t - P_{t-1}}{P_{t-1}}
\]

These returns are used to estimate:
- empirical mean \( \mu \)
- empirical volatility \( \sigma \)

Assumption: returns follow a **normal distribution**, standard in basic VaR modelling.

---

##  Value-at-Risk (VaR)

The VaR at level \( \alpha \) is defined as the **maximum loss** over a horizon 
such that the probability of exceeding it is \( \alpha \).

\[
\text{VaR}_{\alpha} = - \left( \mu_P + \sigma_P \, \Phi^{-1}(\alpha) \right)
\]

where:
- \( \mu_P \) = expected 10-day profit,
- \( \sigma_P \) = volatility of 10-day profit,
- \( \Phi^{-1} \) = inverse standard normal CDF.

Interpretation:

> “With probability \( 1 - \alpha \), the portfolio will not lose more than VaR.”

We compute VaR **without puts** (unhedged portfolio) and **with puts** (hedged portfolio).

---

##  Monte Carlo Simulation

To evaluate the distribution of profits with options, we generate future prices:

\[
S_T = S_0 \left( 1 + \mu \, T + \sigma \sqrt{T} Z \right)
\]

where \( Z \sim N(0,1) \).

For each simulated price, we compute:

- SP500 profit
- put payoff \( \max(K - S_T, 0) \)
- total portfolio profit

This provides an **empirical distribution** of profits → we compute the simulated VaR.

---

##  Black-Scholes Pricing (European Put)

The project uses the standard **Black-Scholes formula** for the price of a European put:

\[
P = Ke^{-rT}\Phi(-d_2) - S_0\Phi(-d_1)
\]

This gives the price per option.  
We compute:

\[
\text{Number of options purchased} 
= \frac{\text{Amount invested in puts}}{\text{Option price}}
\]

---

#  **VBA Implementation**

The repo includes the following VBA modules:

### **1. `VaR.bas`**
Implements:

```vb
Function MoyenEcartVaR(vecPrix, montant, alpha, nbJours)

### **1. `VaR.bas`**
Implements:

Function MoyenEcartVaR_avecOptionVente(S0, montant_S, montant_P, K, r, mu, sigma, alpha, nbJours, nbSim)

