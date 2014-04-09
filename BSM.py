# Black Scholes Merton for Python (from Excel VBA)
# George Fisher MIT Spring 2012
#
# I created BSM routines for Excel first and then converted to Python
# I drew upon published work by Back, Benninga, Hull and Wilmott but the majority is my own original work
#
#      d1, d2
#      N  (en/phi) std normal CDF
#      N' (nprime) std normal PDF
#
#      Binary Options
#      Euro Call & Put
#      Greeks
#      Implied Volatility

#      Put/Call Parity

#      American_Call_Dividend
#      American_Put_Binomial

#
# Also includes
#      risk-neutral prob
#      forward prices & rates
#      CAGR
#      randn/randn_ssdt
#      convert discrete to continuous interest rate

# Interest is
#       risk-free rate
#       domestic risk-free rate for currencies

# Yield is
#       dividend yield for stock
#       lease rate for commodities
#       foreign currency risk-free rate for currencies

# Sigma is the standard deviation of the underlying asset

# Time is a year fraction: for 3-months ... Time = 3/12

# Stock is S_0

# Exercise is K

# => Interest, Yield, Sigma, Time are all annual numbers
# => Time = 0 is the value at maturity
#        most of the functions accomodate this
#        for some, it's infinity or otherwise meaningless
# => Sigma = 0 is also accomodated in most functions

##
##   Utilities
##   ---------
##

# N: the standard-normal CDF

def en(x):
    return phi(x)
    

def phi(x):
    import math
    # constants
    a1 =  0.254829592
    a2 = -0.284496736
    a3 =  1.421413741
    a4 = -1.453152027
    a5 =  1.061405429
    p  =  0.3275911

    # Save the sign of x
    sign = 1
    if x < 0:
        sign = -1
    x = abs(x)/math.sqrt(2.0)

    # A&S formula 7.1.26
    t = 1.0/(1.0 + p*x)
    y = 1.0 - (((((a5*t + a4)*t) + a3)*t + a2)*t + a1)*t*math.exp(-x*x)

    return 0.5*(1.0 + sign*y)


# N': the first derivative of N(x) ... the standard normal PDF

def nprime(x):
    import math
    return math.exp(-0.5 * x * x) / math.sqrt(2 * 3.1415926)

# Random Normal (epsilon)

def RandN():
    # produces a standard normal random variable epsilon
    import random
    return random.gauss(0,1)

def RandN_ssdt(ssdt):
    # produces a standard normal random variable epsilon times sigma*sqrt(deltaT)
    import random
    return random.gauss(0,ssdt)

# binomial tree risk-neutral probability (Hull 7th edition Ch 19 P 409)

def RiskNeutralProb(Interest, Yield, sigma, deltaT):
    import math

    u = math.exp( sigma * math.sqrt(deltaT))
    d = math.exp(-sigma * math.sqrt(deltaT))
    
    a = math.exp((Interest - Yield) * deltaT)
    
    numerator = a - d
    denominator = u - d
    
    return numerator / denominator

# Call & Put prices derived from put-call parity
#                                ---------------

def CallParity(Stock, Exercise, Time, Interest, Yield, Put_price):

    import math
    return Put_price + Stock * math.exp(-Yield * Time) - Exercise * math.exp(-Interest * Time)

def PutParity(Stock, Exercise, Time, Interest, Yield, Call_price):

    import math
    return Call_price + Exercise * math.exp(-Interest * Time) - Stock * math.exp(-Yield * Time)

# forward price

def ForwardPrice(Spot, Time, Interest, Yield):

    import math
    return Spot * math.exp((Interest - Yield) * Time)

# forward rate from Time1 to Time2 (discrete compounding)

def ForwardRate(SpotInterest1, Time1, SpotInterest2, Time2):

    numerator   = (1 + SpotInterest2) ** Time2
    denominator = (1 + SpotInterest1) ** Time1
    
    return ((numerator / denominator) ** (1 / (Time2 - Time1))) - 1

# CAGR

def CAGRd(Starting_value, Ending_Value, Number_of_years):

    # discrete CAGR

    return ((Ending_Value / Starting_value) ** (1 / Number_of_years)) - 1

# Convert TO continuous compounding FROM discrete

def r_continuous(r_discrete, compounding_periods_per_year):

    import math
    m = compounding_periods_per_year
    return m * math.log(1 + r_discrete / m)

# Convert TO discrete compounding FROM continuous
#
# t_discrete = m * (exp(r_continuous / m) - 1)
#
# where m is the number of compounding periods per year
#
def r_discrete(r_continuous, compounding_periods_per_year):

  import math
  m = compounding_periods_per_year
  return m * (math.exp(r_continuous / m) - 1)

# --------------------------------------------------------------------------------

##
##   Black Scholes
##   -------------
##

def dOne(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    return (math.log(Stock / Exercise) + (Interest - Yield + 0.5 * sigma * sigma) * Time) / (sigma * math.sqrt(Time))

def dTwo(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    return (math.log(Stock / Exercise) + (Interest - Yield - 0.5 * sigma * sigma) * Time) / (sigma * math.sqrt(Time))

#
# Binary Options
#

# Digital: Cash or Nothing

def CashCall(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    if Time < 0.000000005:
        if Stock >= Exercise:
            return 1
        else:
            return 0
    
    if sigma == 0:
        sigma = 0.0000000001
    
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    Nd2 = phi(d2_)
    
    return math.exp(-Interest * Time) * Nd2

def CashPut(Stock, Exercise, Time, Interest, Yield, sigma):

    if Time < 0.000000005:
        if Stock >= Exercise:
            return 0
        else:
            return 1
    
    if sigma == 0:
        sigma = 0.0000000001
    
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    Nminusd2 = phi(-d2_)
    
    return math.exp(-Interest * Time) * Nminusd2

# Asset or Nothing

def AssetCall(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    if Time < 0.000000005:
        if Stock >= Exercise:
            return Stock
        else:
            return 0   
    
    if sigma == 0:
        sigma = 0.0000000001
    
    if Exercise < 0.000000005:
        Nd1 = 1
    else:
        d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
        Nd1 = phi(d1_)
    
    return Stock * math.exp(-Yield * Time) * Nd1

def AssetPut(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    if Time < 0.000000005:
        if Stock >= Exercise:
            return 0
        else:
            return Stock
    
    if sigma == 0:
        sigma = 0.0000000001
    
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    Nminusd1 = phi(-d1_)
    
    return Stock * math.exp(-Yield * Time) * Nminusd1

#
# European Call and Put
# ---------------------
#

def EuroCall(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    if Time == 0:
        return max(0, Stock - Exercise)
    
    if sigma == 0:
        return max(0, math.exp(-Yield * Time) * Stock - math.exp(-Interest * Time) * Exercise)
    
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)

    return Stock * math.exp(-Yield * Time) * phi(d1_) - Exercise * math.exp(-Interest * Time) * phi(d2_)

def EuroPut(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    if Time == 0:
        return max(0, Exercise - Stock)
    
    if sigma == 0:
        return max(0, math.exp(-Interest * Time) * Exercise - math.exp(-Yield * Time) * Stock)
        
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)

    return Exercise * math.exp(-Interest * Time) * phi(-d2_) - Stock * math.exp(-Yield * Time) * phi(-d1_)


#
# American Put
# ------------
# Per Kerry Back Chapt5.bas
#

def American_Put_Binomial(S0, K, r, sigma, q, T, N):
    import math
    """
    #
    # Inputs are S0 = initial stock price
    #            K = strike price
    #            r = risk-free rate
    #            sigma = volatility
    #            q = dividend yield
    #            T = time to maturity
    #            N = number of time periods
    #
    """
    dt = T / N                                   # length of time period
    u = math.exp(sigma * math.sqrt(dt))          # size of up step
    d = 1 / u                                    # size of down step
    pu = (math.exp((r - q) * dt) - d) / (u - d)  # probability of up step
    dpu = math.exp(-r * dt) * pu                 # one-period discount x prob of up step
    dpd = math.exp(-r * dt) * (1 - pu)           # one-period discount x prob of down step
    u2 = u * u
    S = S0 * d ** N                              # stock price at bottom node at last date
    PutV[0] = max(K - S, 0)                      # put value at bottom node at last date
    for j in range(1,N+1):
        S = S * u2
        PutV[j] = max(K - S, 0)

    for i in range(N - 1, 0, -1):                # back up in time to date 0
        S = S0 * d ** i                          # stock price at bottom node at date i
        PutV[0] = max(K - S, dpd * PutV(0) + dpu * PutV(1))
        for j in range(1,i+1):                   # step up over nodes at date i
            S = S * u2
            PutV[j] = max(K - S, dpd * PutV(j) + dpu * PutV(j + 1))

    return PutV[0]                               # put value at bottom node at date 0

#
# Greeks from Hull (Edition 7) Chapter 17 p378
# --------------------------------------------
#

# per the Black Scholes PDE for a portfolio of options
# on a single non-dividend-paying underlying stock
#
# Theta + Delta * S * r + Gamma * 0.5 * sigma**2 * S**2 = r * Portfolio_Value

# Per Hull: for large option portfolios, usually created by banks in the
# course of buying and selling OTC options to clients, the portfolio is
# Delta hedged every day and Gamma/Vega hedged as needed
#
#             Delta      Gamma      Vega
# Portfolio   P_delta    P_gamma    P_vega
# Option1     w1*1_delta w1*1_gamma w1*1_vega
# Option2     w2*2_delta w2*2_gamma w2*2_vega
#
# Set the columns equal to zero and solve the simultaneous equations

# Most OTC options are sold close to the money; high gamma and vega
# as (if) the price of the underlying move away gamma and vega decline

# Delta
# -----
#
# If a bank sells a call to a client (short a call)
#   ... it hedges itself with a synthetic long call
#
# Synthetic long call = Delta * Stock_price - bond
# ie., borrow the money to buy Delta shares of the stock
#

def DeltaCall(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    if Time == 0:
        if Stock > Exercise:
            return 1
        else:
            return 0
    
    if sigma == 0:
        sigma = 0.0000000001

    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)

    return math.exp(-Yield * Time) * phi(d1_)

def DeltaPut(Stock, Exercise, Time, Interest, Yield, sigma):

    import math
    if Time == 0:
        if Stock < Exercise:
            return -1
        else:
            return 0
    
    if sigma == 0:
        sigma = 0.0000000001
    
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)

    return math.exp(-Yield * Time) * (phi(d1_) - 1)


#
# Gamma the convexity
# -----
#

def OptionGamma(Stock, Exercise, Time, Interest, Yield, sigma):

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)

    OptionGamma = nprime(d1_) * math.exp(-Yield * Time) _
        / (Stock * sigma * math.sqrt(Time))

#
# Theta the decay in the value of an option/portfolio of options as time passes
# -----
#
# divide by 365 for "per calendar day"; 252 for "per trading day"
#
# In a delta-neutral portfolio, Theta is a proxy for Gamma
#

def ThetaCall(Stock, Exercise, Time, Interest, Yield, sigma):

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    Dim d2_
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nd1_
    Nd1_ = phi(d1_)
    Dim Nd2_
    Nd2_ = phi(d2_)
    
    ThetaCall = -Stock * nprime(d1_) * sigma * math.exp(-Yield * Time) / (2 * math.sqrt(Time)) _
        + Yield * Stock * Nd1_ * math.exp(-Yield * Time) _
        - Interest * Exercise * math.exp(-Interest * Time) * Nd2_

def ThetaPut(Stock, Exercise, Time, Interest, Yield, sigma):

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    Dim d2_
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nminusd1_
    Nminusd1_ = phi(-d1_)
    Dim Nminusd2_
    Nminusd2_ = phi(-d2_)
    
    ThetaPut = -Stock * nprime(d1_) * sigma * math.exp(-Yield * Time) / (2 * math.sqrt(Time)) _
        - Yield * Stock * Nminusd1_ * math.exp(-Yield * Time) _
        + Interest * Exercise * math.exp(-Interest * Time) * Nminusd2_

#
# Vega the sensitivity to changes in the volatility of the underlying
# ----
#
def Vega(Stock, Exercise, Time, Interest, Yield, sigma):

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d1_
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Vega = Stock * math.sqrt(Time) * nprime(d1_) * math.exp(-Yield * Time)

#
# Rho the sensitivity to changes in the interest rate
# ---
#

#
# Note the various Rho calculations see Hull 7th Edition Ch 17 P378
#

def RhoFuturesCall(Stock, Exercise, Time, Interest, Yield, sigma):

    RhoFuturesCall = -EuroCall(Stock, Exercise, Time, Interest, Yield, sigma) * Time

def RhoFuturesPut(Stock, Exercise, Time, Interest, Yield, sigma):

    RhoFuturesPut = -EuroPut(Stock, Exercise, Time, Interest, Yield, sigma) * Time

#
# The Rho corresponding to the domestic interest rate is RhoCall/Put, below
#                              foreign  interest rate is RhoFXCall/Put, shown here
#
def RhoFXCall(Stock, Exercise, Time, Interest, Yield, sigma):

    Dim d1_
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nd1_
    Nd1_ = phi(d1_)
    
    RhoFXCall = -Time * math.exp(-Yield * Time) * Stock * Nd1_

def RhoFXPut(Stock, Exercise, Time, Interest, Yield, sigma):

    Dim d1_
    d1_ = dOne(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nminusd1_
    Nminusd1_ = phi(-d1_)
    
    RhoFXPut = Time * math.exp(-Yield * Time) * Stock * Nminusd1_

#
# "Standard" Rhos
#

def RhoCall(Stock, Exercise, Time, Interest, Yield, sigma):

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d2_
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nd2_
    Nd2_ = phi(d2_)
    
    RhoCall = Exercise * Time * math.exp(-Interest * Time) * Nd2_

def RhoPut(Stock, Exercise, Time, Interest, Yield, sigma):

    If sigma = 0 Then
        sigma = 0.0000000001
    End If

    Dim d2_
    d2_ = dTwo(Stock, Exercise, Time, Interest, Yield, sigma)
    
    Dim Nminusd2_
    Nminusd2_ = phi(-d2_)
    
    RhoPut = -Exercise * Time * math.exp(-Interest * Time) * Nminusd2_

#
# Since Bennigna and Back produce identical numbers
# and MATLAB produced numbers that are +/- 2%, I'm
# inclined to go with these numbers
#

#
# Implied Volatility from Benningna
# ---------------------------------
#
def EuroCallVol(Stock, Exercise, Time, Interest, Yield, Call_price):

    Dim High, Low
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

def EuroPutVol(Stock, Exercise, Time, Interest, Yield, Put_price):

    Dim High, Low
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

#
# Implied Volatility from Kerry Back p64
# Chapt3.bas Newton Raphson technique
# Answer IDENTICAL to Bennigna (EuroCallVol)
#
def Black_Scholes_Call(S, K, r, sigma, q, T):

    Black_Scholes_Call = EuroCall(S, K, T, r, q, sigma)

def Black_Scholes_Call_Implied_Vol(S, K, r, q, T, CallPrice):
#
# Inputs are S = initial stock price
#            K = strike price
#            r = risk-free rate
#            q = dividend yield
#            T = time to maturity
#            CallPrice = call price
#
Dim tol, lower, flower, upper, fupper, guess, fguess
If CallPrice < math.exp(-q * T) * S - math.exp(-r * T) * K Then
    MsgBox ("Option price violates the arbitrage bound.")
    Exit Function
End If
tol = 10 ^ -6
lower = 0
flower = Black_Scholes_Call(S, K, r, lower, q, T) - CallPrice
upper = 1
fupper = Black_Scholes_Call(S, K, r, upper, q, T) - CallPrice
Do While fupper < 0                   # double upper until it is an upper bound
    upper = 2 * upper
    fupper = Black_Scholes_Call(S, K, r, upper, q, T) - CallPrice
Loop
guess = 0.5 * lower + 0.5 * upper
fguess = Black_Scholes_Call(S, K, r, guess, q, T) - CallPrice
Do While upper - lower > tol               # until root is bracketed within tol
    If fupper * fguess < 0 Then            # root is between guess and upper
        lower = guess                      # make guess the new lower bound
        flower = fguess
        guess = 0.5 * lower + 0.5 * upper  # new guess = bi-section
        fguess = Black_Scholes_Call(S, K, r, guess, q, T) - CallPrice
    Else                                   # root is between lower and guess
        upper = guess                      # make guess the new upper bound
        fupper = fguess
        guess = 0.5 * lower + 0.5 * upper  # new guess = bi-section
        fguess = Black_Scholes_Call(S, K, r, guess, q, T) - CallPrice
    End If
Loop
Black_Scholes_Call_Implied_Vol = guess

#
# Implied Volatility from Wilmott Into Ch 8 p192 Newton Raphson***NOT DEBUGGED***
#
def ImpVolCall(Stock, Exercise, Time, Interest, Yield, Call_price):

    Volatility = 0.2
    epsilon = 0.0001
    dv = epsilon + 1
    
    While Abs(dv) > epsilon
        PriceError = EuroCall(Stock, Exercise, Time, Interest, Yield, Volatility) - Call_price
        dv = PriceError / Vega(Stock, Exercise, Time, Interest, Yield, Volatility)
        Volatility = Volatility - dv
    Wend
    
    ImpVolCall = Volatility



#
# from Kerry Back Chapt8.bas ... need Python's "BiNormalProb"
#
def American_Call_Dividend(S, K, r, sigma, Div, TDiv, TCall):
    import math
    """
#
# Inputs are S = initial stock price
#            K = strike price
#            r = risk-free rate
#            sigma = volatility
#            Div = cash dividend
#            TDiv = time until dividend payment
#            TCall = time until option matures >= TDiv
#
"""
    LessDiv = S - math.exp(-r * TDiv) * Div               # stock value excluding dividend
    If Div / K <= 1 - math.exp(-r * (TCall - TDiv)):      # early exercise cannot be optimal
        return Black_Scholes_Call(LessDiv, K, r, sigma, 0, TCall)
    #
    # Now we find an upper bound for the bisection.
    #
    upper = K
    while upper + Div - K < Black_Scholes_Call(upper, K, r, sigma, 0, TCall - TDiv):
        upper = 2 * upper
    #
    # Now we use bisection to compute Zstar = LessDivStar.
    #
    tol = 10 ** -6
    lower = 0
    flower = Div - K
    fupper = upper + Div - K - Black_Scholes_Call(upper, K, r, sigma, 0, TCall - TDiv)
    guess = 0.5 * lower + 0.5 * upper
    fguess = guess + Div - K - Black_Scholes_Call(guess, K, r, sigma, 0, TCall - TDiv)
    
    while upper - lower > tol:
        if fupper * fguess < 0:
            lower = guess
            flower = fguess
            guess = 0.5 * lower + 0.5 * upper
            fguess = guess + Div - K - Black_Scholes_Call(guess, K, r, sigma, 0, TCall - TDiv)
        else:
            upper = guess
            fupper = fguess
            guess = 0.5 * lower + 0.5 * upper
            fguess = guess + Div - K - Black_Scholes_Call(guess, K, r, sigma, 0, TCall - TDiv)

    LessDivStar = guess
    #
    # Now we calculate the probabilities and the option value.
    #
    d1 = (math.log(LessDiv / LessDivStar) + (r + sigma ** 2 / 2) * TDiv) / (sigma * math.sqrt(TDiv))
    d2 = d1 - sigma * math.sqrt(TDiv)
    d1prime = (math.log(LessDiv / K) + (r + sigma ** 2 / 2) * TCall) / (sigma * math.sqrt(TCall))
    d2prime = d1prime - sigma * math.sqrt(TCall)
    rho = -math.sqrt(TDiv / TCall)
    N1 = phi(d1)
    N2 = phi(d2)
    M1 = BiNormalProb(-d1, d1prime, rho)
    M2 = BiNormalProb(-d2, d2prime, rho)
    return LessDiv * N1 + math.exp(-r * TDiv) * (Div - K) * N2 + LessDiv * M1 - math.exp(-r * TCall) * K * M2