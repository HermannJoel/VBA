Attribute VB_Name = "BsOption"
'''1.
BSOptionValue = iopt*(S*eqt*NDOne-X*ert*NDTwo)

Function BSOptionValue(iopt, S, X, r, q, tyr, sigma)
' returns the black-Scholes value(iopt=1 for C, -1 for P; q=div yield)
' uses BSDOne fn
' uses BSDTwo fn
    Dim eqt, ert, NDOne, NDTwo
    eqt=Exp(-q*tyr)
    ert=Exp(-r*tyr)
    if S > 0 and X > 0 and tyr > 0 And sigma > 0 Then
        NDOne=Application.NormSDist(iopt*BSDOne(S, X, r, q, tyr, sigma))
        NDTwo=Application.NormSDist(iopt*BSDTwo(S, X, r, q, tyr, sigma))
        BSOptionValue = iopt*(S*eqt*NDOne-X*ert*NDTwo)
    Else
        BSOptionValue=-1
    End if
End Function

'''2.Options on forward and futures
Function BlackOptionValue(iopt, F, X, r, rfgn, tyr, sigma)
’ returns Black option value for forwards
’ uses BSOptionValue fn
Dim S
S=F Ł Exp((rfgn - r) Ł tyr)
BlackOptionValue = BSOptionValue(iopt, S, X, r, rfgn, tyr, sigma)
End Function


'''3.Option greeks

Function BSOptionGreeks(igreek, iopt, S, X, r, q, tyr, sigma)
’ returns BS option greeks (depends on value of igreek)
’ returns delta(1), gamma(2), rho(3), theta(4) or vega(5)
’ iopt=1 for call, -1 for put; q=div yld
’ uses BSOptionValue fn
’ uses BSDOne fn
’ uses BSDTwo fn
’ uses BSNdashDOne fn
Dim eqt, c, c1, c1d, c2, d, g, v
eqt = Exp(-q Ł tyr)
c = BSOptionValue(iopt, S, X, r, q, tyr, sigma)
c1 = Application.NormSDist(iopt Ł BSDOne(S, X, r, q, tyr, sigma))
c1d = BSNdashDOne(S, X, r, q, tyr, sigma)
c2 = Application.NormSDist(iopt Ł BSDTwo(S, X, r, q, tyr, sigma))
d = iopt Ł eqt Ł c1
g = c1d Ł eqt / (S Ł sigma Ł Sqr(tyr))
v = -1
If igreek = 1 Then v = d
If igreek = 2 Then v = g
If igreek = 3 Then v = iopt Ł X Ł tyr Ł Exp(-r Ł tyr) Ł c2
If igreek = 4 Then v = r Ł c-(r - q) Ł S Ł d - 0.5 Ł (sigma Ł S)O2 Ł g
If igreek = 5 Then v = S Ł Sqr(tyr) Ł c1d Ł eqt
BSOptionGreeks = v
End Function