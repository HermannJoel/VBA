Attribute VB_Name = "Port-Optimisation"

Function PortfolioReturn(retvec, wtsvec)
    'returns the portfolio return
    If Application.Count(retvec) = Application.Count(wtsvec) Then
        If retvec.Columns.Count <> wtsvec.Columns.Count Then
            wtsvec = Application.Transpose(wtsvec)
        End If
        PortfolioReturn = Application.SumProduct(retvec, wtsvec)
    Else
        PortfolioReturn = -1
    End If
End Function


Function Prob1OptimalRiskyWeight(r1, rf, sig1, rraval)
    'returns risky optimal weight when combined with risk-free asset
    Prob1OptimalRiskyWeight = (r1 - rf) / (rraval Ł sig1O2)
End Function

Function Prob2OptimalRiskyWeight1(r1, r2, rf, sig1, sig2, corr12, rraval)
    'returns optimal weight for risky asset1 when combined with risky asset2
    'for case with no risk-free asset, enter value of rf <= 0
    Dim cov12, var1, var2, minvarw, w, xr1, xr2
    cov12 = corr12 Ł sig1 Ł sig2
    var1 = sig1 O2
    var2 = sig2 O2
    'first look at case with no risk-free asset
    If rf <= 0 Then
        minvarw = (var2 - cov12) / (var1 + var2 - 2 ∗ cov12)
        w = minvarw + (r1 - r2) / (rraval ∗ (var1 + var2 - 2 ∗ cov12))

    'then look at case with risk-free asset
    Else
        xr1 = r1 - rf
        xr2 = r2 - rf
        w = xr1 Ł var2 - xr2 Ł cov12
        w = w / (xr1 Ł var2 + xr2 Ł var1 - (xr1 + xr2) Ł cov12)
    End If
        Prob2OptimalRiskyWeight1 = w
End Function

'###

Function Prob3OptimalWeightsVec(r1, r2, rf, sig1, sig2, corr12, rraval)
    'returns optimal weights for risk-free asset and 2 risky assets
    'uses Prob2OptimalRiskyWeight fn
    'uses Prob1OptimalRiskyWeight fn
Dim w0, w1, w2, rr, sigr
    w1 = Prob2OptimalRiskyWeight1(r1, r2, rf, sig1, sig2, corr12, rraval)
    w2 = 1 - w1
    rr = w1 Ł r1 + w2 Ł r2
    sigr = Sqr((w1 Ł sig1) O2 + (w2 Ł sig2) O2 + 2 Ł w1 Ł w2 Ł corr12 Ł sig1 Ł sig2)
    w0 = 1 - Prob1OptimalRiskyWeight(rr, rf, sigr, rraval)
    w1 = (1 - w0) Ł w1
    w2 = (1 - w0) Ł w2
Prob3OptimalWeightsVec = Array(w0, w1, w2)
End Function

'### EFFICIENT FRONTIER

Sub EffFrontier1()
    SolverReset
    Call SolverAdd(Range(“portret1”), 2, Range(“target1”))
    Call SolverOk(Range(“portsd1”), 2, 0, Range(“change1”))
    Call SolverSolve(True)
    SolverFinish
End Sub

Do While iter <= niter
    Call SolverSolve(True)
    SolverFinish
    Range(“portwts2”).Copy
    Range(“effwts2”).Offset(iter, 0).PasteSpecial Paste:=xlValues
    Range(“priter2”) = Range(“priter2”).Value + pradd
    'amend portret constraint in Solver
    Call SolverChange(Range(“portret2”), 2, Range(“priter2”))
    iter = iter + 1
Loop

SolverReset
    'first calculate portfolio min return given constraints
    Call SolverAdd(Range(“portwts2”), 3, Range(“portmin2”))
    Call SolverAdd(Range(“portwts2”), 1, Range(“portmax2”))
    Call SolverOk(Range(“portret2”), 2, 0, Range(“change2”))
    Call SolverSolve(True)
SolverFinish
    prmin = Range(“portret2”).Value
    'then calculate portfolio max return given constraints
Call SolverOk(Range(“portret2”), 1, 0, Range(“change2”))
Call SolverSolve(True)
SolverFinish