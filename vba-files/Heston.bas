Attribute VB_Name = "Heston"
'S = 100,K =100, r = 0, τ = 180/365 = 0.5, v = 0.01, ρ = 0, κ = 2, θ = 0.01, λ = 0, and σ = 0.1.
'Stochastic Vol stock path Heston stochastic volatility price process,

For daycnt = 1 To daynum
    e = Application.NormSInv(Rnd)
    eS = Application.NormSInv(Rnd)
    ev = rho * eS + Sqr(1 - rho ^ 2) * e
    lnSt = lnSt + (r - 0.5 * curv) * deltat + Sqr(curv) * Sqr(deltat) * eS
    curS = Exp(lnSt)
    lnvt = lnvt + (kappa * (theta - curv) - lambda * _
        curv - 0.5 * sigmav) * deltat + sigmav * _
        (1 / Sqr(curv)) * Sqr(deltat) * ev
    curv = Exp(lnvt)
    allS(daycnt) = curS
Next daycnt


Function simPath(kappa, theta, lambda, rho, sigmav, daynum, startS, r, startv, K)
    For itcount = 1 To ITER
    lnSt = Log(startS): lnvt = Log(startv)
    curv = startv: curS = startS
        For daycnt = 1 To daynum

' Stock Path Generating VBA Code ...
        Next daycnt
    Next itcount
    imPath = Application.Transpose(allS)
End Function

'call price by Monte Carlo.
Function HestonMC(kappa, theta, lambda, rho, sigmav, daynum, startS, r, startv, K, ITER)
'More VBA Code
    For itcount = 1 To ITER
        For daycnt = 1 To daynum
        'Stock Path Generating VBA Code ...
        Next daycnt
    simPath = simPath + Exp((-daynum / 365) * r) * _
        Application.Max(allS(daynum) - K, 0)
    Next itcount
        HestonMC = simPath / ITER
End Function

'The closed-form Heston
Function HestonP1(phi, kappa, theta, lambda, rho, sigma, tau, K, S, r, v)
    mu1 = 0.5
    b1 = set_cNum(kappa + lambda - rho * sigma, 0)
    d1 = cNumSqrt(cNumSub(cNumSq(cNumSub(set_cNum(0, rho * _ 
        sigma * phi), b1)), cNumSub(set_cNum(0, sigma ^ 2 * _
        2 * mu1 * phi), set_cNum(sigma ^ 2 * phi ^ 2, 0))))
    g1 = cNumDiv(cNumAdd(cNumSub(b1, set_cNum(0, rho * _
        sigma * phi)), d1), cNumSub(cNumSub(b1, set_cNum(0, rho * sigma * phi)), d1)) 
    DD1_1 = cNumDiv(cNumAdd(cNumSub(b1, set_cNum(0, rho * _
        sigma * phi)), d1), set_cNum(sigma ^ 2, 0))
    DD1_2 = cNumSub(set_cNum(1, 0), cNumExp(cNumProd(d1, set_cNum(tau, 0))))
    DD1_3 = cNumSub(set_cNum(1, 0), cNumProd(g1, cNumExp(cNumProd(d1, set_cNum(tau, 0)))))
    DD1 = cNumProd(DD1_1, cNumDiv(DD1_2, DD1_3))
    CC1_1 = set_cNum(0, r * phi * tau)
    CC1_2 = set_cNum((kappa * theta) / (sigma ^ 2), 0)
    CC1_3 = cNumProd(cNumAdd(cNumSub(b1, set_cNum(0, rho * sigma * phi)), d1), set_cNum(tau, 0))
    CC1_4 = cNumProd(set_cNum(2, 0), cNumLn(cNumDiv
        (cNumSub(set_cNum(1, 0), cNumProd(g1, _
        cNumExp(cNumProd(d1, set_cNum(tau, 0))))), cNumSub(set_cNum(1, 0), g1))))
    cc1 = cNumAdd(CC1_1, cNumProd(CC1_2, cNumSub(CC1_3, CC1_4)))
    f1 = cNumExp(cNumAdd(cNumAdd(cc1, cNumProd(DD1, set_cNum(v, 0))), set_cNum(0,
    phi * Application.Ln(S))))
    HestonP1 = cNumReal(cNumDiv(cNumProd(cNumExp( set_cNum(0, -phi * Application.Ln(K))), f1), set_cNum(0, phi)))
End Function

Function Heston(PutCall As String, kappa, theta, lambda, rho, sigma, tau, K, S, r, v)
    Dim P1_int(1001) As Double, P2_int(1001) As Double, phi_int(1001) As Double
    Dim p1 As Double, p2 As Double, phi As Double, xg(16) As Double, wg(16) As Double
    cnt = 1
    For phi = 0.0001 To 100.0001 Step 0.1
        phi_int(cnt) = phi
        P1_int(cnt) = HestonP1(phi, kappa, theta, lambda, rho, sigma, tau, K, S, r, v)
        P2_int(cnt) = HestonP2(phi, kappa, theta, lambda, rho, sigma, tau, K, S, r, v)
    cnt = cnt + 1
    Next phi
        p1 = 0.5 + (1 / thePI) * TRAPnumint(phi_int, P1_int)
        p2 = 0.5 + (1 / thePI) * TRAPnumint(phi_int, P2_int)
    If p1 < 0 Then p1 = 0
    If p1 > 1 Then p1 = 1
    If p2 < 0 Then p2 = 0
    If p2 > 1 Then p2 = 1

    HestonC = S * p1 - K * Exp(-r * tau) * p2
    If PutCall = "Call" Then
        Heston = HestonC
    ElseIf PutCall = "Put" Then
        Heston = HestonC + K * Exp(-r * tau) - S
    End If
End Function



