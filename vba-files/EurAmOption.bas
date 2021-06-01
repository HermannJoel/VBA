Attribute VB_Name = "EurAmOption"

'EUROPEAN CALL AND PUT PRICING MODEL
'1.WITHOUT DIVIDEND

'fivemonth option with strike price $30 on a non-dividend-paying stock with spot price
'$30 and annual volatility 30 percent, when the risk-free rate is 5 percent. Hence,
'S = 30,K = 30, rf = 0.05, T = 5/12, and σ = 0.3

Function Gauss(X)
    Gauss = Application.NormSDist(X)

End Function

Function BS_call(S, K, rf, T, v)
    d = (Log(S / K) + T * (rf + 0.5 * v ^ 2)) / (v * Sqr(T))
    BS_call = S * Gauss(d) - Exp(-rf * T) * K * Gauss(d - v * Sqr(T))
End Function

Function BS_put(S, K, rf, T, v)
    BS_put = BS_call(S, K, rf, T, v) -S + K * Exp(-rf * T)
End Function

'2.VBA function to price European call options on a dividend-paying stock.
Function BS_div_call(S, K, rf, T, v, Div)
    Divnum = Application.Count(Div) / 3
    PVD = 0
        For i = 1 To Divnum
            PVD = PVD + Div(i, 2) * Exp(-Div(i, 1) * Div(i, 3))
        Next i
    Smod = S - PVD
    BS_div_call = BS_call(Smod, K, rf, T, v)
End Function

'3.American Option with dividend
Function BS_div_amer_call(S, K, rf, T, v, div)
    Dim allCall() As Double
    divnum = Application.Count(div) / 3
    ReDim allCall(divnum + 1) As Double
        For j = 1 To divnum
            PVD = 0
            For i = 1 To j
                If (i < j) Then
                    PVD = PVD + div(i, 2) * Exp(-div(i, 1) * div(i, 3))
                End If
            Next i
            Smod = S - PVD
            allCall(j) = BS_call(Smod, K, rf, div(j, 1), v)
        Next j
    allCall(divnum + 1) = BS_div_call(S, K, rf, T, v, div)
    BS_div_amer_call = Application.Max(allCall)
End Function

'4. Implied volatility
Function ImpliedVolatility(S, K, rf, q, T, v, CallPrice)
    ImpliedVolatility = (CallPrice - BS_call(S, K, rf, q, T, v)) ^ 2
End Function

'5.
Function NewtRaph(S, K, rf, q, T, x_guess, CallPrice)
' More VBA statements
    fx = Run("ImpliedVolatility", S, K, rf, q, T, cur_x, CallPrice)
    cur_x_delta = cur_x - delta_x
    fx_delta = Run("ImpliedVolatility", S, K, rf, q, T, cur_x_delta, CallPrice)
    dx = ((fx - fx_delta) / delta_x)
' More VBA statements
End Function

'6. Greeks

Function Fz(x)
    Fz = Exp(-x ^ 2 / 2) / Sqr(2 * Application.Pi())
End Function

Function BSDelta(S, K, T, r, v, PutCall As String)
    d = (Log(S / K) + T * (r + 0.5 * v ^ 2)) / (v * Sqr(T))
    Select Case PutCall
        Case "Call": BSDelta = Gauss(d)
        Case "Put": BSDelta = Gauss(d) - 1
    End Select
End Function

Function BSGamma(S, K, T, r, v)
    d = (Log(S / K) + T * (r + 0.5 * v ^ 2)) / (v * Sqr(T))
    BSGamma = Fz(d) / S / v / Sqr(T)
End Function

Function BSVega(S, K, T, r, v)
    d = (Log(S / K) + T * (r + 0.5 * v ^ 2)) / (v * Sqr(T))
    BSVega = S * Fz(d) * Sqr(T)
End Function

Function BSRho(S, K, T, r, v, PutCall As String)
    d = (Log(S / K) + T * (r + 0.5 * v ^ 2)) / (v * Sqr(T))
    Select Case PutCall
        Case "Call": BSRho = T * K * Exp(-r * T) * Gauss(d - v * Sqr(T))
        Case "Put": BSRho = -T * K * Exp(-r * T) * Gauss(v * Sqr(T) - d)
    End Select
End Function

Function BSTheta(S, K, T, r, v, PutCall As String)
    d = (Log(S / K) + T * (r + 0.5 * v ^ 2)) / (v * Sqr(T))
    Select Case PutCall
        Case "Call": BSTheta = -S * Fz(d) * v / 2 / Sqr(T) - r * K * Exp(-r * T) * Gauss(d - v * Sqr(T))
        Case "Put": BSTheta = -S * Fz(d) * v / 2 / Sqr(T) + r * K * Exp(-r * T) * Gauss(v * Sqr(T) - d)
    End Select
End Function





