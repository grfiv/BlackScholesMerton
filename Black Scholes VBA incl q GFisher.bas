Attribute VB_Name = "BlackScholes"
Option Explicit

' Black Scholes Excel VBA George Fisher MIT Fall 2011
'      d1, d2
'      N  (en)     std normal CDF
'      N' (nprime) std normal PDF
'
'      Binary Options
'      Euro Call & Put; Black's Model Call & Put
'      Black's Formulation
'      Greeks
'      Implied Volatility

'      Put/Call Parity

'      American_Call_Dividend
'      American_Put_Binomial

'
' Also includes
'      risk-neutral prob
'      nodal value of a binimial tree
'      Monte Carlo terminal value of one stock path
'      Monte Carlo next-step along one stock path
'      forward prices & rates
'      CAGR
'      randn/randn_ssdt
'      cell-display functions
'      convert discrete to continuous interest rate


' Developer > Visual Basic (Alt_F11)
' Insert > Module


' Interest is
'       risk-free rate
'       domestic risk-free rate for currencies

' Yield is
'       dividend yield for stock
'       lease rate for commodities
'       foreign currency risk-free rate for currencies

' Sigma is the standard deviation of the underlying asset

' Time is a year fraction: for 3-months ... Time = 3/12

' Stock is S_0

' Exercise is K

' => Interest, Yield, Sigma, Time are all annual numbers
' => Time = 0 is the value at maturity
'        most of the functions accomodate this
'        for some, it's infinity or otherwise meaningless
' => Sigma = 0 is also accomodated in most functions

''
''   Utilities
''   ---------
''

' N: the standard-normal CDF

Function en(x)

    en = Application.NormSDist(x)

End Function

' N': the first derivative of N(x) ... the standard normal PDF

Function nprime(x As Double) As Double

    nprime = Exp(-0.5 * x * x) / Sqr(2 * 3.1415926)

End Function

' Random Normal (epsilon)

Function RandN()
' produces a standard normal random variable epsilon

    RandN = Application.NormSInv(Rnd())
    
End Function
Function RandN_ssdt(ssdt)
' produces a standard normal random variable epsilon times sigma*sqrt(deltaT)

    RandN_ssdt = Application.NormInv(Rnd(), 0, ssdt)
    
End Function

' Functions to display a formula in a cell (Benningna)

Function GetFormula(Rng As Range) As String

    Application.Volatile True
    GetFormula = "<-- " & Application.Text(Rng.FormulaLocal, "")
    
End Function

Function GgetFormula(Rng As Range) As String

    Application.Volatile True
    GgetFormula = Application.Text(Rng.FormulaLocal, "")
    
End Function

' binomial tree risk-neutral probability (Hull 7th edition Ch 19 P 409)

Function RiskNeutralProb(Interest, Yield, sigma, deltaT)

    Dim u As Double
    u = Exp(sigma * Sqr(deltaT))
    
    Dim d As Double
    d = Exp(-sigma * Sqr(deltaT))
    
    Dim a As Double
    a = Exp((Interest - Yield) * deltaT)
    
    Dim numerator As Double
    numerator = a - d
    
    Dim denominator As Double
    denominator = u - d
    
    RiskNeutralProb = numerator / denominator

End Function

' value of node j at time t in a binomial tree ***NOT DEBUGGED***
'   t starts at zero (on the left)
'   j starts at zero (at the bottom ... all downs)

Function BinomialValue(S_0, u, d, T, j)

    BinomialValue = S_0 * u ^ j * d ^ (T - j)

End Function

' terminal value of one stock path (one of many) for Monte Carlo simulation

Function MCterm(S_0, Interest, sigma, Time)
    
    MCterm = S_0 * Exp((Interest - 0.5 * sigma * sigma) * Time + sigma * RandN() * Sqr(Time))

End Function

' next step along a path for Monte Carlo simulation

Function MCnextStep(prior_value, Interest, sigma, deltaT)
    
    MCnextStep = prior_value + prior_value * (Interest * deltaT + sigma * RandN() * Sqr(deltaT))

End Function

' next step along a path for Monte Carlo simulation;
' Antithetic Variable ***NOT SURE ABOUT THIS***

Function MCnextStepAV(prior_value, Interest, sigma, deltaT)
    
    Dim f1, f2 As Double
    f1 = prior_value + prior_value * (Interest * deltaT + sigma * RandN() * Sqr(deltaT))
    f2 = prior_value + prior_value * (Interest * deltaT + sigma * -RandN() * Sqr(deltaT))
    
    MCnextStepAV = (f1 + f2) / 2

End Function


' Call & Put prices derived from put-call parity

Function CallParity(Stock, Exercise, Time, Interest, Yield, Put_price)

    CallParity = Put_price + _
                 Stock * Exp(-Yield * Time) - _
                 Exercise * Exp(-Interest * Time)
         
End Function

Function PutParity(Stock, Exercise, Time, Interest, Yield, Call_price)

    PutParity = Call_price + _
                Exercise * Exp(-Interest * Time) - _
                Stock * Exp(-Yield * Time)
         
End Function

' forward price

Function ForwardPrice(Spot, Time, Interest, Yield)

    ForwardPrice = Spot * Exp((Interest - Yield) * Time)

End Function

' forward rate from Time1 to Time2 (discrete compounding)

Function ForwardRate(SpotInterest1, Time1, SpotInterest2, Time2)

    Dim numerator As Double
    numerator = (1 + SpotInterest2) ^ Time2
    
    Dim denominator As Double
    denominator = (1 + SpotInterest1) ^ Time1
    
    ForwardRate = ((numerator / denominator) ^ (1 / (Time2 - Time1))) - 1

End Function

' CAGR

Function CAGR(Starting_value, Ending_Value, Number_of_years, Optional form)

    ' the default for is discrete; the alternative is continuous
    ' the parameter "form" is optional; leave out for discrete, put in a zero for continuous

    If IsMissing(form) Then
        CAGR = ((Ending_Value / Starting_value) ^ (1 / Number_of_years)) - 1
    Else
        CAGR = Log(Ending_Value / Starting_value) / Number_of_years
    End If
    

End Function

Function r_continuous(r_discrete, compounding_periods_per_year)

    Dim m As Double
    m = compounding_periods_per_year
    
    r_continuous = m * Log(1 + r_discrete / m)

End Function

' Convert TO discrete compounding FROM continuous
'
' t_discrete = m * (exp(r_continuous / m) - 1)
'
' where m is the number of compounding periods per year
'
Function r_discrete(r_continuous As Double, m As Integer) As Double

  r_discrete = m * (Exp(r_continuous / m) - 1)

End Function

' --------------------------------------------------------------------------------

''
''   Black Scholes
''   -------------
''

Function dOne(Stock, Exercise, Time, Interest, Yield, sigma)

    ' (365 * 24 * 60 * 60) = number of seconds in a year =  31,536,000
    ' 1 / (365 * 24 * 60 * 60) = 3.17097919837646E-08
    ' the year fraction of a second
    If Time <= 0 Then Time = 1E-20

    dOne = (Log(Stock / Exercise) + (Interest - Yield + 0.5 * sigma * sigma) * Time) _
            / (sigma * Sqr(Time))
            
End Function

Function dTwo(Stock, Exercise, Time, Interest, Yield, sigma)

    ' (365 * 24 * 60 * 60) = number of seconds in a year =  31,536,000
    ' 1 / (365 * 24 * 60 * 60) = 3.17097919837646E-08
    ' the year fraction of a second
    If Time <= 0 Then Time = 1E-20
    
    dTwo = (Log(Stock / Exercise) + (Interest - Yield - 0.5 * sigma * sigma) * Time) _
            / (sigma * Sqr(Time))
            
End Function

'
' Binary Options
'

' Digital: Cash or Nothing

Function CashCall(Stock, Exercise, Time, Interest, Yield, sigma)

    If Time < 0.000000005 Then
        If Stock >= Exercise Then
            CashCall = 1
            Exit Function
        Else
            CashCall = 0
            Exit Function
        End If
    End If
    
    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d2_, Nd2
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    Nd2 = Application.NormSDist(d2_)
    
    CashCall = Exp(-Interest * Time) * Nd2

End Function

Function CashPut(Stock, Exercise, Time, Interest, Yield, sigma)

    If Time < 0.000000005 Then
        If Stock >= Exercise Then
            CashPut = 0
            Exit Function
        Else
            CashPut = 1
            Exit Function
        End If
    End If
    
    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d2_, Nminusd2
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    Nminusd2 = Application.NormSDist(-d2_)
    
    CashPut = Exp(-Interest * Time) * Nminusd2

End Function

' Asset or Nothing

Function AssetCall(Stock, Exercise, Time, Interest, Yield, sigma)

    If Time < 0.000000005 Then
        If Stock >= Exercise Then
            AssetCall = Stock
            Exit Function
        Else
            AssetCall = 0
            Exit Function
        End If
    End If
    
    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_, Nd1
    If Exercise < 0 Then Exit Function
    If Exercise < 0.000000005 Then
        Nd1 = 1
    Else
        d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
        Nd1 = Application.NormSDist(d1_)
    End If

    AssetCall = Stock * Exp(-Yield * Time) * Nd1

End Function

Function AssetPut(Stock, Exercise, Time, Interest, Yield, sigma)

    If Time < 0.000000005 Then
       If Stock >= Exercise Then
            AssetPut = 0
            Exit Function
        Else
            AssetPut = Stock
            Exit Function
        End If
    End If
    
    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_, Nminusd1
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    Nminusd1 = Application.NormSDist(-d1_)
    
    AssetPut = Stock * Exp(-Yield * Time) * Nminusd1

End Function

'
'  Black's Formulation
'

Function BFormulationCall(Forward, Exercise, Time, Interest, Yield, sigma)

    Dim d1 As Double, d2 As Double
    d1 = (Log(Forward / Exercise) / (sigma * Sqr(Time))) + ((sigma * Sqr(Time)) / 2)
    d2 = (Log(Forward / Exercise) / (sigma * Sqr(Time))) - ((sigma * Sqr(Time)) / 2)
    
    BFormulationCall = Exp(-Interest * Time) * (Forward * en(d1) - Exercise * en(d2))
  
End Function

'
' European Call and Put
'

Function EuroCall(Stock, Exercise, Time, Interest, Yield, sigma)

    If Time = 0 Then
        EuroCall = Application.Max(0, Stock - Exercise)
        Exit Function
    End If
    
    If sigma = 0 Then
        EuroCall = Application.Max(0, Exp(-Yield * Time) * Stock - Exp(-Interest * Time) * Exercise)
        Exit Function
    End If
    
    Dim d1_ As Double, d2_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)

    EuroCall = Stock * Exp(-Yield * Time) * Application.NormSDist(d1_) _
               - Exercise * Exp(-Interest * Time) * Application.NormSDist(d2_)
     
End Function

Function EuroPut(Stock, Exercise, Time, Interest, Yield, sigma)

    If Time = 0 Then
        EuroPut = Application.Max(0, Exercise - Stock)
        Exit Function
    End If
    
    If sigma = 0 Then
        EuroPut = Application.Max(0, Exp(-Interest * Time) * Exercise - Exp(-Yield * Time) * Stock)
        Exit Function
    End If

    EuroPut = Exercise * Exp(-Interest * Time) * _
     Application.NormSDist(-dTwo(Stock, Exercise, Time, Interest, Yield, sigma)) - _
     Stock * Exp(-Yield * Time) * Application.NormSDist(-dOne(Stock, Exercise, _
        Time, Interest, Yield, sigma))
        
End Function

'
' Black's Model
'
Function VeronessiCaplet(n, Z, f, K, sigma, Tminus1)
    ' caplet for period T_i+1
    '
    ' n = time in years between periods (1 for a year; 0.25 for quarterly)
    ' Z = discount rate Z(0, T_i+1)
    ' f = forward rate  f(0, T_i, T_i+1) = (1/n) * [Z(0, T_i) / Z(0, T_i+1) -1]
    ' K = exercise rate/cap rate
    ' sigma = annual vol of f
    ' Tminus1 = T_i
    '
    
    Dim d1_ As Double, d2_ As Double
    d1_ = (1 / (sigma * Sqr(Tminus1))) * Log(f / K) + 0.5 * sigma * Sqr(Tminus1)
    d2_ = d1_ - sigma * Sqr(Tminus1)
    
    VeronessiCaplet = n * (Z * 100) * (f * Application.NormSDist(d1_) - K * Application.NormSDist(d2_))
    
End Function

Function BlacksModelCall(Forward, Exercise, Time, Spot_Yield_to_term_of_option, vol_of_forward)

    Dim Interest As Double, sigma As Double
    Interest = Spot_Yield_to_term_of_option
    sigma = vol_of_forward

    If Time = 0 Then
        BlacksModelCall = Application.Max(0, Forward - Exercise)
        Exit Function
    End If
    
    If sigma = 0 Then
        BlacksModelCall = Application.Max(0, Exp(-Interest * Time) * (Forward - Exercise))
        Exit Function
    End If
    
    Dim d1_ As Double, d2_ As Double
    d1_ = (Log(Forward / Exercise) + ((1 / 2) * sigma ^ 2) * Time) / (sigma * Sqr(Time))
    d2_ = (Log(Forward / Exercise) - ((1 / 2) * sigma ^ 2) * Time) / (sigma * Sqr(Time))

    BlacksModelCall = Exp(-Interest * Time) * (Forward * Application.NormSDist(d1_) _
                                             - Exercise * Application.NormSDist(d2_))
     
End Function
Function BlacksModelPut(Forward, Exercise, Time, Spot_Yield_to_term_of_option, vol_of_forward)

    Dim Interest As Double, sigma As Double
    Interest = Spot_Yield_to_term_of_option
    sigma = vol_of_forward

    If Time = 0 Then
        BlacksModelPut = Application.Max(0, Exercise - Forward)
        Exit Function
    End If
    
    If sigma = 0 Then
        BlacksModelPut = Application.Max(0, Exp(-Interest * Time) * (Exercise - Forward))
        Exit Function
    End If
    
    Dim d1_ As Double, d2_ As Double
    d1_ = (Log(Forward / Exercise) + ((1 / 2) * sigma ^ 2) * Time) / (sigma * Sqr(Time))
    d2_ = (Log(Forward / Exercise) - ((1 / 2) * sigma ^ 2) * Time) / (sigma * Sqr(Time))

    BlacksModelPut = Exp(-Interest * Time) * (Exercise * Application.NormSDist(-d2_) _
                                            - Forward * Application.NormSDist(-d1_))
     
End Function

'
' Per Kerry Back Chapt5.bas
'

Function American_Put_Binomial(S0, K, r, sigma, q, T, n)
'
' Inputs are S0 = initial stock price
'            K = strike price
'            r = risk-free rate
'            sigma = volatility
'            q = dividend yield
'            T = time to maturity
'            N = number of time periods
'
Dim dt, u, d, pu, dpu, dpd, u2, S, i, j
Dim PutV() As Double
ReDim PutV(n)
dt = T / n                              ' length of time period
u = Exp(sigma * Sqr(dt))                ' size of up step
d = 1 / u                               ' size of down step
pu = (Exp((r - q) * dt) - d) / (u - d)  ' probability of up step
dpu = Exp(-r * dt) * pu                 ' one-period discount x prob of up step
dpd = Exp(-r * dt) * (1 - pu)           ' one-period discount x prob of down step
u2 = u * u
S = S0 * d ^ n                          ' stock price at bottom node at last date
PutV(0) = Application.Max(K - S, 0)     ' put value at bottom node at last date
For j = 1 To n
    S = S * u2
    PutV(j) = Application.Max(K - S, 0)
Next j
For i = n - 1 To 0 Step -1              ' back up in time to date 0
    S = S0 * d ^ i                      ' stock price at bottom node at date i
    PutV(0) = Application.Max(K - S, dpd * PutV(0) + dpu * PutV(1))
    For j = 1 To i                      ' step up over nodes at date i
        S = S * u2
        PutV(j) = Application.Max(K - S, dpd * PutV(j) + dpu * PutV(j + 1))
    Next j
Next i
American_Put_Binomial = PutV(0)         ' put value at bottom node at date 0
End Function

'
' from Kerry Back Chapt8.bas
'
Function American_Call_Dividend(S, K, r, sigma, Div, TDiv, TCall)
'
' Inputs are S = initial stock price
'            K = strike price
'            r = risk-free rate
'            sigma = volatility
'            Div = cash dividend
'            TDiv = time until dividend payment
'            TCall = time until option matures >= TDiv
'
Dim LessDiv, upper, tol, lower, flower, fupper, guess, fguess
Dim LessDivStar, d1, d2, d1prime, d2prime, rho, N1, N2, M1, M2
LessDiv = S - Exp(-r * TDiv) * Div          ' stock value excluding dividend
If Div / K <= 1 - Exp(-r * (TCall - TDiv)) Then  ' early exercise cannot be optimal
    American_Call_Dividend = Black_Scholes_Call(LessDiv, K, r, sigma, 0, TCall)
    Exit Function
End If
'
' Now we find an upper bound for the bisection.
'
upper = K
Do While upper + Div - K < Black_Scholes_Call(upper, K, r, sigma, 0, TCall - TDiv)
   upper = 2 * upper
Loop
'
' Now we use bisection to compute Zstar = LessDivStar.
'
tol = 10 ^ -6
lower = 0
flower = Div - K
fupper = upper + Div - K - Black_Scholes_Call(upper, K, r, sigma, 0, TCall - TDiv)
guess = 0.5 * lower + 0.5 * upper
fguess = guess + Div - K - Black_Scholes_Call(guess, K, r, sigma, 0, TCall - TDiv)
Do While upper - lower > tol
    If fupper * fguess < 0 Then
        lower = guess
        flower = fguess
        guess = 0.5 * lower + 0.5 * upper
        fguess = guess + Div - K _
               - Black_Scholes_Call(guess, K, r, sigma, 0, TCall - TDiv)
    Else
        upper = guess
        fupper = fguess
        guess = 0.5 * lower + 0.5 * upper
        fguess = guess + Div - K _
               - Black_Scholes_Call(guess, K, r, sigma, 0, TCall - TDiv)
    End If
Loop
LessDivStar = guess
'
' Now we calculate the probabilities and the option value.
'
d1 = (Log(LessDiv / LessDivStar) _
   + (r + sigma ^ 2 / 2) * TDiv) / (sigma * Sqr(TDiv))
d2 = d1 - sigma * Sqr(TDiv)
d1prime = (Log(LessDiv / K) _
        + (r + sigma ^ 2 / 2) * TCall) / (sigma * Sqr(TCall))
d2prime = d1prime - sigma * Sqr(TCall)
rho = -Sqr(TDiv / TCall)
N1 = Application.NormSDist(d1)
N2 = Application.NormSDist(d2)
M1 = BiNormalProb(-d1, d1prime, rho)
M2 = BiNormalProb(-d2, d2prime, rho)
American_Call_Dividend = LessDiv * N1 + Exp(-r * TDiv) * (Div - K) * N2 _
                       + LessDiv * M1 - Exp(-r * TCall) * K * M2
End Function

'
' Greeks from Hull (Edition 7) Chapter 17 p378
' --------------------------------------------
'

' per the Black Scholes PDE for a portfolio of options
' on a single non-dividend-paying underlying stock
'
' Theta + Delta * S * r + Gamma * 0.5 * sigma^2 * S^2 = r * Portfolio_Value

' Per Hull: for large option portfolios, usually created by banks in the
' course of buying and selling OTC options to clients, the portfolio is
' Delta hedged every day and Gamma/Vega hedged as needed
'
'             Delta      Gamma      Vega
' Portfolio   P_delta    P_gamma    P_vega
' Option1     w1*1_delta w1*1_gamma w1*1_vega
' Option2     w2*2_delta w2*2_gamma w2*2_vega
'
' Set the columns equal to zero and solve the simultaneous equations

' Most OTC options are sold close to the money; high gamma and vega
' as (if) the price of the underlying move away gamma and vega decline

' Delta
' -----
'
' If a bank sells a call to a client (short a call)
'   ... it hedges itself with a synthetic long call
'
' Synthetic long call = Delta * Stock_price - bond
' ie., borrow the money to buy Delta shares of the stock
'

Function DeltaCall(Stock, Exercise, Time, Interest, Yield, sigma)

    If Time = 0 Then
        If Stock > Exercise Then
            DeltaCall = 1
            Exit Function
        Else
            DeltaCall = 0
            Exit Function
        End If
    End If
    
    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)

    DeltaCall = Exp(-Yield * Time) * Application.NormSDist(d1_)
    
End Function

Function DeltaPut(Stock, Exercise, Time, Interest, Yield, sigma)

    If Time = 0 Then
        If Stock < Exercise Then
            DeltaPut = -1
            Exit Function
        Else
            DeltaPut = 0
            Exit Function
        End If
    End If
    
    If sigma = 0 Then
        sigma = 0.0000000001
    End If
    
    Dim d1_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)

    DeltaPut = Exp(-Yield * Time) * (Application.NormSDist(d1_) - 1)
        
End Function

'
' Gamma the convexity
' -----
'

Function OptionGamma(Stock, Exercise, Time, Interest, Yield, sigma)

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)

    OptionGamma = nprime(d1_) * Exp(-Yield * Time) _
        / (Stock * sigma * Sqr(Time))

End Function

'
' Theta the decay in the value of an option/portfolio of options as time passes
' -----
'
' divide by 365 for "per calendar day"; 252 for "per trading day"
'
' In a delta-neutral portfolio, Theta is a proxy for Gamma
'

Function ThetaCall(Stock, Exercise, Time, Interest, Yield, sigma)

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    Dim d2_ As Double
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nd1_ As Double
    Nd1_ = Application.NormSDist(d1_)
    Dim Nd2_ As Double
    Nd2_ = Application.NormSDist(d2_)
    
    ThetaCall = -Stock * nprime(d1_) * sigma * Exp(-Yield * Time) / (2 * Sqr(Time)) _
        + Yield * Stock * Nd1_ * Exp(-Yield * Time) _
        - Interest * Exercise * Exp(-Interest * Time) * Nd2_
    
End Function

Function ThetaPut(Stock, Exercise, Time, Interest, Yield, sigma)

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    Dim d2_ As Double
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nminusd1_ As Double
    Nminusd1_ = Application.NormSDist(-d1_)
    Dim Nminusd2_ As Double
    Nminusd2_ = Application.NormSDist(-d2_)
    
    ThetaPut = -Stock * nprime(d1_) * sigma * Exp(-Yield * Time) / (2 * Sqr(Time)) _
        - Yield * Stock * Nminusd1_ * Exp(-Yield * Time) _
        + Interest * Exercise * Exp(-Interest * Time) * Nminusd2_
    
End Function

'
' Vega the sensitivity to changes in the volatility of the underlying
' ----
'
Function Vega(Stock, Exercise, Time, Interest, Yield, sigma)

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Vega = Stock * Sqr(Time) * nprime(d1_) * Exp(-Yield * Time)
    
End Function

'
' Rho the sensitivity to changes in the interest rate
' ---
'

'
' Note the various Rho calculations see Hull 7th Edition Ch 17 P378
'

Function RhoFuturesCall(Stock, Exercise, Time, Interest, Yield, sigma)

    RhoFuturesCall = -EuroCall(Stock, Exercise, Time, Interest, Yield, sigma) * Time
    
End Function
Function RhoFuturesPut(Stock, Exercise, Time, Interest, Yield, sigma)

    RhoFuturesPut = -EuroPut(Stock, Exercise, Time, Interest, Yield, sigma) * Time
    
End Function

'
' The Rho corresponding to the domestic interest rate is RhoCall/Put, below
'                              foreign  interest rate is RhoFXCall/Put, shown here
'
Function RhoFXCall(Stock, Exercise, Time, Interest, Yield, sigma)

    Dim d1_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nd1_ As Double
    Nd1_ = Application.NormSDist(d1_)
    
    RhoFXCall = -Time * Exp(-Yield * Time) * Stock * Nd1_
    
End Function
Function RhoFXPut(Stock, Exercise, Time, Interest, Yield, sigma)

    Dim d1_ As Double
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nminusd1_ As Double
    Nminusd1_ = Application.NormSDist(-d1_)
    
    RhoFXPut = Time * Exp(-Yield * Time) * Stock * Nminusd1_
    
End Function

'
' "Standard" Rhos
'

Function RhoCall(Stock, Exercise, Time, Interest, Yield, sigma)

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d2_ As Double
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nd2_ As Double
    Nd2_ = Application.NormSDist(d2_)
    
    RhoCall = Exercise * Time * Exp(-Interest * Time) * Nd2_
    
End Function

Function RhoPut(Stock, Exercise, Time, Interest, Yield, sigma)

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d2_ As Double
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nminusd2_ As Double
    Nminusd2_ = Application.NormSDist(-d2_)
    
    RhoPut = -Exercise * Time * Exp(-Interest * Time) * Nminusd2_
    
End Function

'
' Since Bennigna and Back produce identical numbers
' and MATLAB produced numbers that are +/- 2%, I'm
' inclined to go with these numbers
'

'
' Implied Volatility from Benningna
' ---------------------------------
'
Function EuroCallVol(Stock, Exercise, Time, Interest, Yield, Call_price)

    Dim High, Low As Double
    High = 2
    Low = 0
    Do While (High - Low) > 0.000001
    If EuroCall(Stock, Exercise, Time, Interest, Yield, (High + Low) / 2) > _
        Call_price Then
             High = (High + Low) / 2
             Else: Low = (High + Low) / 2
    End If
    Loop
    EuroCallVol = (High + Low) / 2
    
End Function

Function EuroPutVol(Stock, Exercise, Time, Interest, Yield, Put_price)

    Dim High, Low As Double
    High = 2
    Low = 0
    Do While (High - Low) > 0.000001
    If EuroPut(Stock, Exercise, Time, Interest, Yield, (High + Low) / 2) > _
        Put_price Then
             High = (High + Low) / 2
             Else: Low = (High + Low) / 2
    End If
    Loop
    EuroPutVol = (High + Low) / 2
    
End Function

'
' Implied Volatility from Kerry Back p64
' Chapt3.bas Newton Raphson technique
' Answer IDENTICAL to Bennigna (EuroCallVol)
'
Function Black_Scholes_Call(S, K, r, sigma, q, T)

    Black_Scholes_Call = EuroCall(S, K, T, r, q, sigma)

End Function
Function Black_Scholes_Call_Implied_Vol(S, K, r, q, T, CallPrice)
'
' Inputs are S = initial stock price
'            K = strike price
'            r = risk-free rate
'            q = dividend yield
'            T = time to maturity
'            CallPrice = call price
'
Dim tol, lower, flower, upper, fupper, guess, fguess
If CallPrice < Exp(-q * T) * S - Exp(-r * T) * K Then
    MsgBox ("Option price violates the arbitrage bound.")
    Exit Function
End If
tol = 10 ^ -6
lower = 0
flower = Black_Scholes_Call(S, K, r, lower, q, T) - CallPrice
upper = 1
fupper = Black_Scholes_Call(S, K, r, upper, q, T) - CallPrice
Do While fupper < 0                   ' double upper until it is an upper bound
    upper = 2 * upper
    fupper = Black_Scholes_Call(S, K, r, upper, q, T) - CallPrice
Loop
guess = 0.5 * lower + 0.5 * upper
fguess = Black_Scholes_Call(S, K, r, guess, q, T) - CallPrice
Do While upper - lower > tol               ' until root is bracketed within tol
    If fupper * fguess < 0 Then            ' root is between guess and upper
        lower = guess                      ' make guess the new lower bound
        flower = fguess
        guess = 0.5 * lower + 0.5 * upper  ' new guess = bi-section
        fguess = Black_Scholes_Call(S, K, r, guess, q, T) - CallPrice
    Else                                   ' root is between lower and guess
        upper = guess                      ' make guess the new upper bound
        fupper = fguess
        guess = 0.5 * lower + 0.5 * upper  ' new guess = bi-section
        fguess = Black_Scholes_Call(S, K, r, guess, q, T) - CallPrice
    End If
Loop
Black_Scholes_Call_Implied_Vol = guess
End Function

'
' Implied Volatility from Wilmott Into Ch 8 p192 Newton Raphson***NOT DEBUGGED***
'
Function ImpVolCall(Stock, Exercise, Time, Interest, Yield, Call_price)

    Volatility = 0.2
    epsilon = 0.0001
    dv = epsilon + 1
    
    While Abs(dv) > epsilon
        PriceError = EuroCall(Stock, Exercise, Time, Interest, Yield, Volatility) - Call_price
        dv = PriceError / Vega(Stock, Exercise, Time, Interest, Yield, Volatility)
        Volatility = Volatility - dv
    Wend
    
    ImpVolCall = Volatility

End Function
