Attribute VB_Name = "ExoticOptions"

'DI + DO = Plain Vanilla
'UI + UO = Plain Vanilla

Function NewBarrier(Spot, Bar, T, v, M1, M2 As String)
    If Bar > Spot Then
        Sign = 1
    Else
        Sign = -1
    End If
    Select Case M1
        Case Is <> 0
            Bar = Bar * Exp(Sign * 0.5826 * v * Sqr(T / M1))
        Case Is = 0
            Select Case M2
                Case "H": Bar = Bar * Exp(Sign * 0.5826 * v * Sqr(1 / 24 / 365))
                Case "D": Bar = Bar * Exp(Sign * 0.5826 * v * Sqr(1 / 365))
                Case "W": Bar = Bar * Exp(Sign * 0.5826 * v * Sqr(1 / 52))
                Case "M": Bar = Bar * Exp(Sign * 0.5826 * v * Sqr(1 / 12))
            End Select
    End Select
    NewBarrier = Bar
End Function

'Pricing Barrier Options with Binomial Tree

Function FBBarrierC(Spot, K, H, T, r, sigma, n, OpType As String, PutCall
As String)
    dt = T / n: u0 = Exp(sigma * Sqr(dt)): d0 = 1 / u0
    p0 = (Exp(r * dt) - d0) / (u0 - d0)
    N0 = Round(Log(H / Spot) / Log(d0))
    L = (Log(H / Spot) + N0 * sigma * Sqr(dt)) / N0 / sigma ^ 2 / dt
    u = Exp(sigma * Sqr(dt) + L * sigma ^ 2 * dt)
    d = Exp(-sigma * Sqr(dt) + L * sigma ^ 2 * dt)
    p = (Exp(r * dt) - d) / (u - d)
    Dim S() As Double, Op() As Double
    ReDim S(n + 1, n + 1) As Double, Op(n + 1, n + 1) As Double
    S(1, 1) = Spot
    For j = 1 To n + 1
        For i = 1 To j
            If j <= N0 + 1 Then
                S(i, j) = S(1, 1) * u ^ (j - i) * d ^ (i - 1)
            Else
                S(i, j) = S(1, 1) * d ^ (N0) * u0 ^ (j - i) * d0 ^ (i - N0 - 1)
            End If
                S(i, j) = Round(S(i, j) * 10000, 5) / 10000
        Next i
    Next j

    For i = 1 To n + 1
    If (S(i, n + 1) <= H And (OpType = "DI" Or OpType = "DO")) Or (S(i, n + 1)
                    >= H And (OpType = "UI" Or OpType = "UO")) Then
        Op(i, n + 1) = 0
    Else
    Select Case PutCall
        Case "Call"
            Op(i, n + 1) = Application.Max(S(i, n + 1) - K, 0)
        Case "Put"
            Op(i, n + 1) = Application.Max(K - S(i, n + 1), 0)
    End Select
    End If
    Next i


    For j = n To 1 Step -1
    For i = 1 To j
        If (S(i, j) <= H And (OpType = "DI" Or OpType = "DO")) Or
            (S(i, j) >= H And (OpType = "UI" Or OpType = "UO")
        Then
            Op(i, j) = 0
        Else
            If j <= N0 Then
                Prob = p
            Else
                Prob = p0
            End If
        Op(i, j) = Exp(-r * dt) * (Prob * Op(i, j + 1) + (1 - Prob) * Op(i + 1, j + 1))
        End If
    Next i
    Next j     


    If OpType = "DO" Or OpType = "UO" Then
        FBBarrierC = Op(1, 1)
    Else
        FBBarrierC = Binomial(Spot, K, T, r, sigma, n, PutCall) - Op(1, 1)
    End If
End Function  

'Pricing Barrier Options with Binomial Tree

Function BarrierBin(Spot, Strike, Bar, T, r, v, old_n, PutCall As String,
                    EuroAmer As String, BarType As String, M1, M2 As String)
    Bar = NewBarrier(Spot, Bar, T, v, M1, M2)
    If (BarType = "DO" Or BarType = "DI") And Bar > Spot Then
        MsgBox "Error: Barrier Must be Below Spot Price for a Down-and-Out or Down-and-In Option"
    ElseIf (BarType = "UO" Or BarType = "UI") And Bar < Spot Then 
        MsgBox"Error: Barrier Must be Above Spot Price for a Up-and-Out or Up-and-In Option"
    Else
    For m = 1 To 100
        F(m) = m ^ 2 * v ^ 2 * T / (Log(Spot / Bar)) ^ 2
    Next m
    If old_n < F(1) Then
        MsgBox ("Increase Number Steps to at Least " & Application.Floor(F(1) + 1, 1))
    Else
    For i = 1 To 99
        If (F(i) < old_n) And (old_n < F(i + 1)) Then
            n = Application.Floor(F(i + 1), 1)
        Exit For
        End If
    Next i
    End If

'2 step CRR parameters
    dt = T / n: u = Exp(v * Sqr(dt))
    d = 1 / u: p = (Exp(r * dt) - d) / (u - d)
    exp_rT = Exp(-r * dt)
    For i = 1 To n + 1
    AssetPrice = Spot * u ^ (n + 1 - i) * d ^ (i - 1)
        If  ((BarType = "DO" Or BarType = "DI") And _
            AssetPrice <= Bar) Or _
            ((BarType = "UO" Or BarType = "UI") And _
            AssetPrice >= Bar) Then
            Op(i, n + 1) = 0
        ElseIf PutCall = "Call" Then
            Op(i, n + 1) = Application.Max(AssetPrice - Strike, 0)
        End If
    Next i
    For j = n To 1 Step -1
    For i = 1 To j
    AssetPrice = Spot * u ^ (j - i) * d ^ (i - 1)
        If ((BarType = "DO" Or BarType = "DI") And _
            AssetPrice <= Bar) Or _
            ((BarType = "UO" Or BarType = "UI") And _
            AssetPrice >= Bar) Then 
            Op(i, j) = 0
        Else
            Op(i, j) = exp_rT * (p * Op(i, j + 1) + _
            (1 - p) * Op(i + 1, j + 1))
        End If
    Next i
    Next j


    If BarType = "DO" Or BarType = "UO" Then 
        output(1, 1) = Op(1, 1)

    Else
    Select Case EuroAmer
    Case "Euro"
        If PutCall = "Call" Then
            output(1, 1) = fbinomial(Spot, Strike, r, v, T, n, "Call", "Euro") - Op(1, 1)

        ElseIf PutCall = "Put" Then
            output(1, 1) = fbinomial(Spot, Strike, r, v, T, n, "Put", "Euro") - Op(1, 1)
        End If
    Case "Amer"
        If PutCall = "Call" And BarType = "DI" Then
            output(1, 1) = (Spot / Bar) ^ (1 - 2 * r / v ^ 2) * _
            fbinomial(Bar ^ 2 / Spot, Strike, r, v, T, n, "Call", "Amer")
        ElseIf PutCall = "Put" And BarType = "UI" Then 
            output(1, 1) = (Spot / Bar) ^ (1 - 2 * r / v ^ 2) * _
            fbinomial(Bar ^ 2 / Spot, Strike, r, v, T, n, "Put","Amer")
        ElseIf (PutCall = "Put" And BarType = "DI") Then
            MsgBox "You Cannot Price an American Down-and-In Put _
                    Using This Algorithm"
        ElseIf (PutCall = "Call" And BarType = "UI") Then
            MsgBox "You Cannot Price an American Up-and-In Call _
                    Using This Algorithm"
        End If
    End Select
    End If
    output(2, 1) = n
    BarrierBin = output
End Function
