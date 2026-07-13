Attribute VB_Name = "CFH_Accounting_revised"
Option Explicit

Private gSteps As Long

Private sA_FV()       As Double
Private sA_Intr()     As Double
Private sA_TV()       As Double
Private sA_CFHR()     As Double
Private sA_COH()      As Double
Private sA_COHamort() As Double
Private sA_Ineff()    As Double
Private sA_Alive()    As Double

Private qA_FV()       As Double
Private qA_CFHR()     As Double

Private sB_FVw()      As Double
Private sB_FVx()      As Double
Private sB_CFHRw()    As Double
Private sB_CFHRx()    As Double
Private sB_Ineffw()   As Double
Private sB_Ineffx()   As Double
Private sB_FVTPL()    As Double
Private qB_FVx()      As Double
Private qB_CFHRx()    As Double

Private repSurv As Long, repKO As Long
Private detS_A() As Double, detS_Bw() As Double, detS_Bx() As Double
Private detK_A() As Double, detK_Bw() As Double, detK_Bx() As Double

Private tA_FVpeak()  As Double
Private tA_CFHRend() As Double
Private tA_IneffTot() As Double
Private tB_FVgross() As Double
Private tB_CFHRend() As Double
Private tB_IneffTot() As Double
Private tB_postKO()  As Double
Private tKOflag()    As Long

Private AP_S1() As Double
Private AP_S2() As Double
Private AP_KO() As Long

Public Sub Run_CFH_Accounting_Engine()

    Dim wsL As Worksheet, wsE As Worksheet
    On Error Resume Next
    Set wsL = Sheets("LSMC"): Set wsE = Sheets("Encoding")
    On Error GoTo 0
    If wsL Is Nothing Or wsE Is Nothing Then
        MsgBox "LSMC and/or Encoding sheet not found.", vbCritical: Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail

    Dim Lambda   As Double: Lambda = wsL.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsL.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsL.Range("B3").Value
    Dim Drift1P  As Double: Drift1P = wsL.Range("B4").Value
    Dim Drift2P  As Double: Drift2P = wsL.Range("B5").Value
    Dim KO_up    As Double: KO_up = wsL.Range("B6").Value
    Dim KO_dn    As Double: KO_dn = wsL.Range("B7").Value
    Dim S1_0     As Double: S1_0 = wsL.Range("B8").Value
    Dim S2_0     As Double: S2_0 = wsL.Range("B9").Value
    Dim vol1     As Double: vol1 = wsL.Range("B10").Value
    Dim vol2     As Double: vol2 = wsL.Range("B11").Value
    Dim corr     As Double: corr = wsL.Range("B12").Value

    Dim r_US  As Double: r_US = wsE.Range("B4").Value
    Dim r_KRW As Double: r_KRW = wsE.Range("B5").Value
    Dim barrels As Double: barrels = wsE.Range("B9").Value
    Dim WACC  As Double: WACC = wsE.Range("B12").Value

    Dim T_WTI As Double: T_WTI = wsE.Range("B15").Value
    Dim T_FX  As Double: T_FX = wsE.Range("B16").Value
    Dim T_total As Double: T_total = WorksheetFunction.Max(T_WTI, T_FX)
    If T_total <= 0 Then T_total = wsL.Range("B14").Value

    Dim K As Double: K = S1_0

    Dim wsIn As Worksheet: Set wsIn = EnsureSheet("CFH_Inputs")
    SeedCFHInputs wsIn
    Dim n_acct As Long: n_acct = CLng(wsIn.Range("B2").Value)
    If n_acct < 1 Then n_acct = 20000

    Dim hedge_steps As Long: hedge_steps = CLng(T_total * STEPS_PER_YEAR)
    Dim dt As Double: dt = T_total / hedge_steps
    gSteps = hedge_steps

    Dim primed As Double
    If Beta_Mat_Steps <> hedge_steps Then
        primed = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1P, Drift2P, _
                    KO_up, KO_dn, S1_0, S2_0, vol1, vol2, corr, _
                    hedge_steps, T_total, 50000, K, True)
    End If
    Dim Base_Premium As Double: Base_Premium = wsL.Range("J2").Value
    If Base_Premium <= 0# And primed > 0# Then Base_Premium = primed

    Call Build_AcctPaths(n_acct, hedge_steps, S1_0, S2_0, vol1, vol2, corr, _
            Drift1P, Drift2P, Lambda, JumpMean, JumpVol, KO_up, KO_dn, T_total)

    Dim F0_WTI As Double: F0_WTI = S1_0 * Exp(r_US * T_total)
    Dim G0_FX  As Double: G0_FX = S2_0 * Exp((r_KRW - r_US) * T_total)
    Dim N_FX   As Double: N_FX = barrels * F0_WTI

    Dim TV0 As Double: TV0 = Base_Premium * barrels

    repSurv = 0: repKO = 0
    Dim s As Long
    For s = 1 To n_acct
        If repSurv = 0 And AP_KO(s) = 0 Then repSurv = s
        If repKO = 0 And AP_KO(s) > 0 Then repKO = s
        If repSurv > 0 And repKO > 0 Then Exit For
    Next s
    If repSurv = 0 Then repSurv = 1
    If repKO = 0 Then repKO = 1

    AllocAccumulators hedge_steps, n_acct

    Dim sim As Long
    For sim = 1 To n_acct
        WalkPath sim, hedge_steps, dt, S1_0, S2_0, K, barrels, _
                 r_US, r_KRW, F0_WTI, G0_FX, N_FX, TV0, _
                 (sim = repSurv), (sim = repKO)
    Next sim

    WriteLedgerA hedge_steps, n_acct, dt
    WriteLedgerB hedge_steps, n_acct, dt
    WriteSFPProforma hedge_steps, n_acct, dt, barrels, Base_Premium
    WriteComparison hedge_steps, n_acct, Base_Premium, barrels, TV0, N_FX

    Call RestoreProductionBetaMat

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "CFH Accounting Engine complete." & vbCrLf & vbCrLf & _
           "Paths: " & n_acct & "   Daily steps: " & hedge_steps & vbCrLf & _
           "Base Premium (unit): " & Format(Base_Premium, "#,##0.00") & " KRW" & vbCrLf & _
           "Firm notional (barrels): " & Format(barrels, "#,##0") & vbCrLf & _
           "F0_WTI: " & Format(F0_WTI, "0.00") & "   N_FX: " & Format(N_FX, "#,##0") & vbCrLf & vbCrLf & _
           "Sheets written: CFH_A_Ledger, CFH_B_Ledger_WTI, CFH_B_Ledger_FX," & vbCrLf & _
           "CFH_SFP_Proforma, CFH_Comparison." & vbCrLf & _
           "Beta_Mat restored to production state.", vbInformation, "CFH Accounting"
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_CFH_Accounting_Engine failed: " & Err.Description, vbCritical
End Sub

Private Sub WalkPath(ByVal sim As Long, ByVal Steps As Long, ByVal dt As Double, _
                     ByVal S1_0 As Double, ByVal S2_0 As Double, ByVal K As Double, _
                     ByVal barrels As Double, ByVal r_US As Double, ByVal r_KRW As Double, _
                     ByVal F0_WTI As Double, ByVal G0_FX As Double, ByVal N_FX As Double, _
                     ByVal TV0 As Double, ByVal isSurv As Boolean, ByVal isKOrep As Boolean)

    Dim T_total As Double: T_total = Steps * dt
    Dim koStep As Long: koStep = AP_KO(sim)

    Dim aCFHR As Double: aCFHR = 0#
    Dim aIneff As Double: aIneff = 0#
    Dim aIntrPrev As Double: aIntrPrev = 0#
    Dim aCOHamort As Double: aCOHamort = 0#
    Dim coh_step As Double: coh_step = TV0 / Steps
    Dim aDead As Boolean: aDead = False

    Dim bCFHRw As Double: bCFHRw = 0#
    Dim bCFHRx As Double: bCFHRx = 0#
    Dim bIneffw As Double: bIneffw = 0#
    Dim bIneffx As Double: bIneffx = 0#
    Dim bFVTPL As Double: bFVTPL = 0#
    Dim cumHIw_prev As Double: cumHIw_prev = 0#
    Dim cumHDw_prev As Double: cumHDw_prev = 0#
    Dim cumHIx_prev As Double: cumHIx_prev = 0#
    Dim cumHDx_prev As Double: cumHDx_prev = 0#
    Dim effW_prev As Double: effW_prev = 0#
    Dim effX_prev As Double: effX_prev = 0#
    Dim fvWprev As Double: fvWprev = 0#
    Dim fvXprev As Double: fvXprev = 0#
    Dim bDiscont As Boolean: bDiscont = False

    Dim stp As Long
    Dim S1 As Double, S2 As Double, tt As Double, tau As Double
    Dim Alive As Boolean

    For stp = 0 To Steps

        If stp = 0 Then
            S1 = S1_0: S2 = S2_0
        Else
            S1 = AP_S1(sim, stp): S2 = AP_S2(sim, stp)
        End If
        tt = stp * dt
        tau = T_total - tt

        Alive = (koStep = 0) Or (stp < koStep)
        Dim koHere As Boolean: koHere = (koStep > 0 And stp = koStep)

        Dim aFV As Double, aIntr As Double, aTV As Double
        Dim vUnit As Double, intrUnit As Double
        If aDead Then
            aFV = 0#: aIntr = 0#: aTV = 0#
        ElseIf koHere Then

            aFV = 0#: aIntr = 0#: aTV = 0#
            aDead = True
        Else
            intrUnit = WorksheetFunction.Max(S1 - K, 0#) * S2
            If stp = 0 Then
                vUnit = TV0 / barrels
            Else
                vUnit = QuantoFV_FromBeta(S1, S2, S1_0, S2_0, stp)
                If vUnit < intrUnit Then vUnit = intrUnit
            End If
            aFV = vUnit * barrels
            aIntr = intrUnit * barrels
            aTV = aFV - aIntr
            If aTV < 0# Then aTV = 0#
        End If

        If stp > 0 And Not aDead Then
            Dim dIntr As Double: dIntr = aIntr - aIntrPrev
            aCFHR = aCFHR + dIntr

            aCOHamort = aCOHamort + coh_step
        ElseIf stp > 0 And koHere Then
            aCOHamort = aCOHamort + coh_step
        End If
        aIntrPrev = aIntr

        Dim aCOH As Double
        aCOH = aTV - aCOHamort
        If aCOH < 0# Then aCOH = 0#

        Dim Fw As Double: Fw = ForwardWTI(S1, tau, r_US)
        Dim fvWTI As Double: fvWTI = (Fw - F0_WTI) * barrels * Exp(-r_US * tau) * S2
        Dim cumHIw As Double: cumHIw = fvWTI
        Dim cumHDw As Double: cumHDw = fvWTI

        Dim Gx As Double: Gx = ForwardFX(S2, tau, r_US, r_KRW)
        Dim cumHIx As Double: cumHIx = N_FX * (Gx - G0_FX) * Exp(-r_KRW * tau)
        Dim N_hypo As Double: N_hypo = barrels * Fw
        Dim cumHDx As Double: cumHDx = N_hypo * (Gx - G0_FX) * Exp(-r_KRW * tau)

        If stp > 0 Then
            If Not bDiscont And koHere Then bDiscont = True

            If Not bDiscont Then

                Dim effW As Double: effW = SignedLowerOf(cumHIw, cumHDw)
                bCFHRw = bCFHRw + (effW - effW_prev)
                bIneffw = bIneffw + ((cumHIw - cumHIw_prev) - (effW - effW_prev))
                effW_prev = effW

                Dim effX As Double: effX = SignedLowerOf(cumHIx, cumHDx)
                bCFHRx = bCFHRx + (effX - effX_prev)
                bIneffx = bIneffx + ((cumHIx - cumHIx_prev) - (effX - effX_prev))
                effX_prev = effX
            Else

                bFVTPL = bFVTPL + (fvWTI - fvWprev) + (cumHIx - fvXprev)
            End If
        End If
        cumHIw_prev = cumHIw: cumHDw_prev = cumHDw
        cumHIx_prev = cumHIx: cumHDx_prev = cumHDx
        fvWprev = fvWTI: fvXprev = cumHIx

        sA_FV(stp) = sA_FV(stp) + aFV:   qA_FV(stp) = qA_FV(stp) + aFV * aFV
        sA_Intr(stp) = sA_Intr(stp) + aIntr
        sA_TV(stp) = sA_TV(stp) + aTV
        sA_CFHR(stp) = sA_CFHR(stp) + aCFHR: qA_CFHR(stp) = qA_CFHR(stp) + aCFHR * aCFHR
        sA_COH(stp) = sA_COH(stp) + aCOH
        sA_COHamort(stp) = sA_COHamort(stp) + aCOHamort
        sA_Ineff(stp) = sA_Ineff(stp) + aIneff
        If Alive Or koHere Then sA_Alive(stp) = sA_Alive(stp) + 1#

        sB_FVw(stp) = sB_FVw(stp) + fvWTI
        sB_FVx(stp) = sB_FVx(stp) + cumHIx: qB_FVx(stp) = qB_FVx(stp) + cumHIx * cumHIx
        sB_CFHRw(stp) = sB_CFHRw(stp) + bCFHRw
        sB_CFHRx(stp) = sB_CFHRx(stp) + bCFHRx: qB_CFHRx(stp) = qB_CFHRx(stp) + bCFHRx * bCFHRx
        sB_Ineffw(stp) = sB_Ineffw(stp) + bIneffw
        sB_Ineffx(stp) = sB_Ineffx(stp) + bIneffx
        sB_FVTPL(stp) = sB_FVTPL(stp) + bFVTPL

        If isSurv Then
            detS_A(stp, 0) = aFV: detS_A(stp, 1) = aIntr: detS_A(stp, 2) = aTV
            detS_A(stp, 3) = aCFHR: detS_A(stp, 4) = aCOH: detS_A(stp, 5) = aIneff
            detS_A(stp, 6) = S1: detS_A(stp, 7) = S2
            detS_Bw(stp, 0) = fvWTI: detS_Bw(stp, 1) = bCFHRw: detS_Bw(stp, 2) = bIneffw
            detS_Bx(stp, 0) = cumHIx: detS_Bx(stp, 1) = bCFHRx: detS_Bx(stp, 2) = bIneffx
            detS_Bx(stp, 3) = bFVTPL
        End If
        If isKOrep Then
            detK_A(stp, 0) = aFV: detK_A(stp, 1) = aIntr: detK_A(stp, 2) = aTV
            detK_A(stp, 3) = aCFHR: detK_A(stp, 4) = aCOH: detK_A(stp, 5) = aIneff
            detK_A(stp, 6) = S1: detK_A(stp, 7) = S2
            detK_Bw(stp, 0) = fvWTI: detK_Bw(stp, 1) = bCFHRw: detK_Bw(stp, 2) = bIneffw
            detK_Bx(stp, 0) = cumHIx: detK_Bx(stp, 1) = bCFHRx: detK_Bx(stp, 2) = bIneffx
            detK_Bx(stp, 3) = bFVTPL
        End If

        If aFV > tA_FVpeak(sim) Then tA_FVpeak(sim) = aFV
    Next stp

    tA_CFHRend(sim) = aCFHR
    tA_IneffTot(sim) = aIneff
    tB_FVgross(sim) = Abs(fvWprev) + Abs(fvXprev)
    tB_CFHRend(sim) = bCFHRw + bCFHRx
    tB_IneffTot(sim) = bIneffw + bIneffx
    tB_postKO(sim) = bFVTPL
    tKOflag(sim) = IIf(koStep > 0, 1, 0)
End Sub

Private Function QuantoFV_FromBeta(ByVal S1 As Double, ByVal S2 As Double, _
                                   ByVal S1_0 As Double, ByVal S2_0 As Double, _
                                   ByVal stp As Long) As Double
    Dim v1 As Double: v1 = S1 / S1_0
    Dim v2 As Double: v2 = S2 / S2_0
    Dim vb As Double
    vb = Beta_Mat(stp, 0) + Beta_Mat(stp, 1) * v1 + Beta_Mat(stp, 2) * v2 _
       + Beta_Mat(stp, 3) * v1 ^ 2 + Beta_Mat(stp, 4) * v2 ^ 2 _
       + Beta_Mat(stp, 5) * v1 * v2
    If vb < 0# Then vb = 0#
    QuantoFV_FromBeta = vb
End Function

Private Sub Build_AcctPaths(ByVal n As Long, ByVal Steps As Long, _
        ByVal S0_WTI As Double, ByVal S0_FX As Double, _
        ByVal vol1 As Double, ByVal vol2 As Double, ByVal corr As Double, _
        ByVal Drift1 As Double, ByVal Drift2 As Double, _
        ByVal Lambda As Double, ByVal JumpMean As Double, ByVal JumpVol As Double, _
        ByVal KO_up As Double, ByVal KO_dn As Double, ByVal T_total As Double)

    Dim dt    As Double: dt = T_total / Steps
    Dim kappa As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt
    Dim sqdt  As Double: sqdt = Sqr(dt)

    ReDim AP_S1(1 To n, 0 To Steps)
    ReDim AP_S2(1 To n, 0 To Steps)
    ReDim AP_KO(1 To n)

    Randomize 20240101

    Dim sim As Long, stp As Long
    Dim S1 As Double, S2 As Double, S1_prev As Double
    Dim z1 As Double, z2 As Double, e2 As Double, nJ As Long
    Dim ko_step As Long, p_up As Double, p_dn As Double

    For sim = 1 To n
        S1 = S0_WTI: S2 = S0_FX: S1_prev = S0_WTI: ko_step = 0
        AP_S1(sim, 0) = S0_WTI: AP_S2(sim, 0) = S0_FX

        For stp = 1 To Steps
            z1 = GetNormal()
            z2 = GetNormal()
            e2 = corr * z1 + Sqr(1# - corr ^ 2) * z2
            nJ = GetPoisson(Lambda * dt)

            S1 = S1 * Exp((Drift1 - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt _
                         + vol1 * sqdt * z1 + JumpSum(nJ, JumpMean, JumpVol))
            S2 = S2 * Exp((Drift2 - 0.5 * vol2 ^ 2) * dt + vol2 * sqdt * e2)

            AP_S1(sim, stp) = S1
            AP_S2(sim, stp) = S2

            If ko_step = 0 Then
                If S1 >= KO_up Or S1 <= KO_dn Then
                    ko_step = stp
                ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then
                    p_up = Exp(-2# * Log(KO_up / S1_prev) * Log(KO_up / S1) / lnVar)
                    p_dn = Exp(-2# * Log(S1_prev / KO_dn) * Log(S1 / KO_dn) / lnVar)
                    If Rnd() < p_up Then ko_step = stp
                    If ko_step = 0 And Rnd() < p_dn Then ko_step = stp
                End If
            End If
            S1_prev = S1
        Next stp

        AP_KO(sim) = ko_step
    Next sim
End Sub

Private Function ForwardWTI(ByVal S1 As Double, ByVal tau As Double, ByVal r_US As Double) As Double
    ForwardWTI = S1 * Exp(r_US * tau)
End Function

Private Function ForwardFX(ByVal S2 As Double, ByVal tau As Double, _
                           ByVal r_US As Double, ByVal r_KRW As Double) As Double
    ForwardFX = S2 * Exp((r_KRW - r_US) * tau)
End Function

Private Function SignedLowerOf(ByVal cumHI As Double, ByVal cumHD As Double) As Double
    Dim m As Double
    m = WorksheetFunction.Min(Abs(cumHI), Abs(cumHD))
    SignedLowerOf = IIf(cumHI < 0#, -m, m)
End Function

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = nm
    End If
    Set EnsureSheet = ws
End Function

Private Sub SeedCFHInputs(ByVal ws As Worksheet)
    If Len(CStr(ws.Range("A1").Value)) = 0 Then
        ws.Range("A1").Value = "CFH Accounting Inputs":  ws.Range("A1").Font.Bold = True
        ws.Range("A2").Value = "n_acct (realised paths)":          ws.Range("B2").Value = 20000
        ws.Range("A3").Value = "n_sens (FV-validation reprice)":   ws.Range("B3").Value = 1000
        ws.Range("A4").Value = "FX notional basis (1=S0,2=F0,3=budget)": ws.Range("B4").Value = 2
        ws.Range("A5").Value = "COH amortisation (1=straight-line)":     ws.Range("B5").Value = 1
        ws.Range("A6").Value = "Reporting cadence (1=daily)":            ws.Range("B6").Value = 1
        ws.Range("A7").Value = "KO forecast highly-probable (1=yes)":    ws.Range("B7").Value = 1
        ws.Columns("A:B").AutoFit
    End If
End Sub

Private Sub AllocAccumulators(ByVal Steps As Long, ByVal n As Long)
    ReDim sA_FV(0 To Steps): ReDim sA_Intr(0 To Steps): ReDim sA_TV(0 To Steps)
    ReDim sA_CFHR(0 To Steps): ReDim sA_COH(0 To Steps): ReDim sA_COHamort(0 To Steps)
    ReDim sA_Ineff(0 To Steps): ReDim sA_Alive(0 To Steps)
    ReDim qA_FV(0 To Steps): ReDim qA_CFHR(0 To Steps)
    ReDim sB_FVw(0 To Steps): ReDim sB_FVx(0 To Steps)
    ReDim sB_CFHRw(0 To Steps): ReDim sB_CFHRx(0 To Steps)
    ReDim sB_Ineffw(0 To Steps): ReDim sB_Ineffx(0 To Steps): ReDim sB_FVTPL(0 To Steps)
    ReDim qB_FVx(0 To Steps): ReDim qB_CFHRx(0 To Steps)
    ReDim detS_A(0 To Steps, 0 To 7): ReDim detS_Bw(0 To Steps, 0 To 2): ReDim detS_Bx(0 To Steps, 0 To 3)
    ReDim detK_A(0 To Steps, 0 To 7): ReDim detK_Bw(0 To Steps, 0 To 2): ReDim detK_Bx(0 To Steps, 0 To 3)
    ReDim tA_FVpeak(1 To n): ReDim tA_CFHRend(1 To n): ReDim tA_IneffTot(1 To n)
    ReDim tB_FVgross(1 To n): ReDim tB_CFHRend(1 To n): ReDim tB_IneffTot(1 To n)
    ReDim tB_postKO(1 To n): ReDim tKOflag(1 To n)
End Sub

Private Sub WriteLedgerA(ByVal Steps As Long, ByVal n As Long, ByVal dt As Double)
    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_A_Ledger")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Structure A -- Single Combined Quanto KO CFH (firm-scale KRW)"
    ws.Range("A1").Font.Bold = True
    Dim hdr As Variant
    hdr = Array("Step", "t(yr)", "MeanFV", "StdFV", "MeanIntrinsic", "MeanTimeValue", _
                "MeanCFHR(OCI)", "StdCFHR", "MeanCOH(OCI)", "CumCOHamort(P&L)", _
                "CumIneff(P&L)", "AliveFrac", _
                "repSurv_FV", "repSurv_CFHR", "repSurv_COH", "repSurv_S1", _
                "repKO_FV", "repKO_CFHR", "repKO_S1")
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Value = hdr
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Font.Bold = True

    Dim out() As Variant: ReDim out(0 To Steps, 0 To UBound(hdr))
    Dim i As Long
    For i = 0 To Steps
        out(i, 0) = i: out(i, 1) = i * dt
        out(i, 2) = sA_FV(i) / n: out(i, 3) = StdFrom(sA_FV(i), qA_FV(i), n)
        out(i, 4) = sA_Intr(i) / n: out(i, 5) = sA_TV(i) / n
        out(i, 6) = sA_CFHR(i) / n: out(i, 7) = StdFrom(sA_CFHR(i), qA_CFHR(i), n)
        out(i, 8) = sA_COH(i) / n: out(i, 9) = sA_COHamort(i) / n
        out(i, 10) = sA_Ineff(i) / n: out(i, 11) = sA_Alive(i) / n
        out(i, 12) = detS_A(i, 0): out(i, 13) = detS_A(i, 3): out(i, 14) = detS_A(i, 4): out(i, 15) = detS_A(i, 6)
        out(i, 16) = detK_A(i, 0): out(i, 17) = detK_A(i, 3): out(i, 18) = detK_A(i, 6)
    Next i
    ws.Range("A4").Resize(Steps + 1, UBound(hdr) + 1).Value = out
    ws.Columns.AutoFit
End Sub

Private Sub WriteLedgerB(ByVal Steps As Long, ByVal n As Long, ByVal dt As Double)

    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_B_Ledger_WTI")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Structure B Layer 1 -- WTI forward CFH (firm-scale KRW)"
    ws.Range("A1").Font.Bold = True
    Dim h1 As Variant
    h1 = Array("Step", "t(yr)", "MeanFV", "MeanCFHR(OCI)", "CumIneff(P&L)", _
               "repSurv_FV", "repSurv_CFHR", "repKO_FV", "repKO_CFHR")
    ws.Range("A3").Resize(1, UBound(h1) + 1).Value = h1
    ws.Range("A3").Resize(1, UBound(h1) + 1).Font.Bold = True
    Dim o1() As Variant: ReDim o1(0 To Steps, 0 To UBound(h1))
    Dim i As Long
    For i = 0 To Steps
        o1(i, 0) = i: o1(i, 1) = i * dt
        o1(i, 2) = sB_FVw(i) / n: o1(i, 3) = sB_CFHRw(i) / n: o1(i, 4) = sB_Ineffw(i) / n
        o1(i, 5) = detS_Bw(i, 0): o1(i, 6) = detS_Bw(i, 1)
        o1(i, 7) = detK_Bw(i, 0): o1(i, 8) = detK_Bw(i, 1)
    Next i
    ws.Range("A4").Resize(Steps + 1, UBound(h1) + 1).Value = o1
    ws.Columns.AutoFit

    Dim ws2 As Worksheet: Set ws2 = EnsureSheet("CFH_B_Ledger_FX")
    ws2.Cells.ClearContents
    ws2.Range("A1").Value = "Structure B Layer 2 -- FX forward CFH (FIXED notional; mismatch leg)"
    ws2.Range("A1").Font.Bold = True
    Dim h2 As Variant
    h2 = Array("Step", "t(yr)", "MeanFV", "StdFV", "MeanCFHR(OCI)", "StdCFHR", _
               "CumIneff(P&L)", "CumPostKO_FVTPL(P&L)", _
               "repSurv_FV", "repSurv_CFHR", "repSurv_Ineff", _
               "repKO_FV", "repKO_CFHR", "repKO_FVTPL")
    ws2.Range("A3").Resize(1, UBound(h2) + 1).Value = h2
    ws2.Range("A3").Resize(1, UBound(h2) + 1).Font.Bold = True
    Dim o2() As Variant: ReDim o2(0 To Steps, 0 To UBound(h2))
    For i = 0 To Steps
        o2(i, 0) = i: o2(i, 1) = i * dt
        o2(i, 2) = sB_FVx(i) / n: o2(i, 3) = StdFrom(sB_FVx(i), qB_FVx(i), n)
        o2(i, 4) = sB_CFHRx(i) / n: o2(i, 5) = StdFrom(sB_CFHRx(i), qB_CFHRx(i), n)
        o2(i, 6) = sB_Ineffx(i) / n: o2(i, 7) = sB_FVTPL(i) / n
        o2(i, 8) = detS_Bx(i, 0): o2(i, 9) = detS_Bx(i, 1): o2(i, 10) = detS_Bx(i, 2)
        o2(i, 11) = detK_Bx(i, 0): o2(i, 12) = detK_Bx(i, 1): o2(i, 13) = detK_Bx(i, 3)
    Next i
    ws2.Range("A4").Resize(Steps + 1, UBound(h2) + 1).Value = o2
    ws2.Columns.AutoFit
End Sub

Private Sub WriteSFPProforma(ByVal Steps As Long, ByVal n As Long, ByVal dt As Double, _
                             ByVal barrels As Double, ByVal Base_Premium As Double)
    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_SFP_Proforma")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "SFP Pro-forma (cross-path mean, firm-scale KRW): Structure A vs Structure B"
    ws.Range("A1").Font.Bold = True
    Dim hdr As Variant
    hdr = Array("Step", "t(yr)", _
                "A_Derivative", "A_OCI(CFHR+COH)", "A_RetainedEarnings", "A_#DerivLines", _
                "B_Deriv_WTI", "B_Deriv_FX", "B_OCI(CFHR_W+CFHR_X)", "B_RetainedEarnings", "B_#DerivLines")
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Value = hdr
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Font.Bold = True
    Dim out() As Variant: ReDim out(0 To Steps, 0 To UBound(hdr))
    Dim i As Long
    For i = 0 To Steps
        out(i, 0) = i: out(i, 1) = i * dt
        out(i, 2) = sA_FV(i) / n
        out(i, 3) = (sA_CFHR(i) + sA_COH(i)) / n
        out(i, 4) = -(sA_COHamort(i) + sA_Ineff(i)) / n
        out(i, 5) = 1
        out(i, 6) = sB_FVw(i) / n
        out(i, 7) = sB_FVx(i) / n
        out(i, 8) = (sB_CFHRw(i) + sB_CFHRx(i)) / n
        out(i, 9) = -(sB_Ineffw(i) + sB_Ineffx(i) + sB_FVTPL(i)) / n
        out(i, 10) = 2
    Next i
    ws.Range("A4").Resize(Steps + 1, UBound(hdr) + 1).Value = out
    ws.Columns.AutoFit
End Sub

Private Sub WriteComparison(ByVal Steps As Long, ByVal n As Long, _
                            ByVal Base_Premium As Double, ByVal barrels As Double, _
                            ByVal TV0 As Double, ByVal N_FX As Double)
    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_Comparison")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "CFH Comparison -- Hypotheses H1-H4 (cross-path, n=" & n & ")"
    ws.Range("A1").Font.Bold = True

    Dim r As Long: r = 3
    ws.Cells(r, 1) = "t0 sanity": ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2) = "Base_Premium x barrels": ws.Cells(r, 3) = TV0
    ws.Cells(r, 4) = "A_FV(0)": ws.Cells(r, 5) = sA_FV(0) / n
    ws.Cells(r, 6) = "A_COH(0)": ws.Cells(r, 7) = sA_COH(0) / n
    r = r + 2

    ws.Cells(r, 1) = "H1: Derivative SFP volatility": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    ws.Cells(r, 1) = "A peak-FV std (across paths)": ws.Cells(r, 2) = StdArr(tA_FVpeak, n): r = r + 1
    ws.Cells(r, 1) = "B gross-FV std (across paths)": ws.Cells(r, 2) = StdArr(tB_FVgross, n): r = r + 1
    ws.Cells(r, 1) = "A # derivative lines on SFP": ws.Cells(r, 2) = 1: r = r + 1
    ws.Cells(r, 1) = "B # derivative lines on SFP (gross)": ws.Cells(r, 2) = 2: r = r + 2

    ws.Cells(r, 1) = "H2: OCI reserve stability & ineffectiveness": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    ws.Cells(r, 1) = "A end-CFHR std": ws.Cells(r, 2) = StdArr(tA_CFHRend, n): r = r + 1
    ws.Cells(r, 1) = "B end-CFHR std": ws.Cells(r, 2) = StdArr(tB_CFHRend, n): r = r + 1
    ws.Cells(r, 1) = "A mean |cum ineffectiveness|": ws.Cells(r, 2) = MeanAbsArr(tA_IneffTot, n): r = r + 1
    ws.Cells(r, 1) = "B mean |cum ineffectiveness|": ws.Cells(r, 2) = MeanAbsArr(tB_IneffTot, n): r = r + 1
    ws.Cells(r, 1) = "  (B FX leg is the structural-mismatch driver)": r = r + 2

    ws.Cells(r, 1) = "H3: KO-event SFP impact (KO paths only)": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    Dim koCount As Long, sumPostKO As Double, sumPostKO2 As Double
    Dim i As Long
    For i = 1 To n
        If tKOflag(i) = 1 Then
            koCount = koCount + 1
            sumPostKO = sumPostKO + tB_postKO(i)
            sumPostKO2 = sumPostKO2 + tB_postKO(i) * tB_postKO(i)
        End If
    Next i
    ws.Cells(r, 1) = "KO rate": ws.Cells(r, 2) = koCount / n: r = r + 1
    ws.Cells(r, 1) = "A post-KO derivative on SFP": ws.Cells(r, 2) = 0: ws.Cells(r, 3) = "(extinguished; CFHR frozen->COGS recycle)": r = r + 1
    ws.Cells(r, 1) = "B mean post-KO FVTPL P&L (KO paths)"
    ws.Cells(r, 2) = IIf(koCount > 0, sumPostKO / koCount, 0#): r = r + 1
    ws.Cells(r, 1) = "B std post-KO FVTPL P&L (KO paths)"
    ws.Cells(r, 2) = IIf(koCount > 0, Sqr(WorksheetFunction.Max(0#, sumPostKO2 / koCount - (sumPostKO / koCount) ^ 2)), 0#): r = r + 2

    ws.Cells(r, 1) = "H4: Disclosure / compliance burden": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    ws.Cells(r, 1) = "Hedging relationships": ws.Cells(r, 2) = 1: ws.Cells(r, 3) = 2: r = r + 1
    ws.Cells(r, 1) = "Effectiveness tests / period": ws.Cells(r, 2) = 1: ws.Cells(r, 3) = 2: r = r + 1
    ws.Cells(r, 1) = "Hedge-documentation sets": ws.Cells(r, 2) = 1: ws.Cells(r, 3) = 2: r = r + 1
    ws.Cells(r, 1) = "  (col B = Structure A, col C = Structure B)": r = r + 1

    ws.Columns("A:G").AutoFit
End Sub

Private Function StdFrom(ByVal sumv As Double, ByVal sumsq As Double, ByVal n As Long) As Double
    If n < 2 Then StdFrom = 0#: Exit Function
    Dim m As Double: m = sumv / n
    Dim v As Double: v = sumsq / n - m * m
    If v < 0# Then v = 0#
    StdFrom = Sqr(v)
End Function

Private Function StdArr(ByRef a() As Double, ByVal n As Long) As Double
    Dim s As Double, S2 As Double, i As Long
    For i = 1 To n
        s = s + a(i): S2 = S2 + a(i) * a(i)
    Next i
    StdArr = StdFrom(s, S2, n)
End Function

Private Function MeanAbsArr(ByRef a() As Double, ByVal n As Long) As Double
    Dim s As Double, i As Long
    For i = 1 To n
        s = s + Abs(a(i))
    Next i
    MeanAbsArr = IIf(n > 0, s / n, 0#)
End Function

Public Sub Run_CFH_FV_Validation()
    Dim wsL As Worksheet: Set wsL = Sheets("LSMC")
    Dim wsE As Worksheet: Set wsE = Sheets("Encoding")

    Dim Lambda As Double: Lambda = wsL.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsL.Range("B2").Value
    Dim JumpVol As Double: JumpVol = wsL.Range("B3").Value
    Dim KO_up As Double: KO_up = wsL.Range("B6").Value
    Dim KO_dn As Double: KO_dn = wsL.Range("B7").Value
    Dim S1_0 As Double: S1_0 = wsL.Range("B8").Value
    Dim S2_0 As Double: S2_0 = wsL.Range("B9").Value
    Dim vol1 As Double: vol1 = wsL.Range("B10").Value
    Dim vol2 As Double: vol2 = wsL.Range("B11").Value
    Dim corr As Double: corr = wsL.Range("B12").Value
    Dim r_US As Double: r_US = wsE.Range("B4").Value
    Dim r_KRW As Double: r_KRW = wsE.Range("B5").Value
    Dim T_total As Double: T_total = WorksheetFunction.Max(wsE.Range("B15").Value, wsE.Range("B16").Value)
    Dim K As Double: K = S1_0

    Dim wsIn As Worksheet: Set wsIn = EnsureSheet("CFH_Inputs"): SeedCFHInputs wsIn
    Dim n_sens As Long: n_sens = CLng(wsIn.Range("B3").Value)
    If n_sens < 200 Then n_sens = 1000

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_Comparison")
    Dim r As Long: r = 40
    ws.Cells(r, 1) = "FV Validation (Beta_Mat surface vs reduced-n LSMC reprice)"
    ws.Cells(r, 1).Font.Bold = True: r = r + 1
    ws.Cells(r, 1) = "tau/T": ws.Cells(r, 2) = "S1": ws.Cells(r, 3) = "S2"
    ws.Cells(r, 4) = "Beta_Mat FV(unit)": ws.Cells(r, 5) = "LSMC reprice(unit)"
    ws.Cells(r, 6) = "Abs diff": ws.Cells(r, 7) = "Pct of LSMC"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Font.Bold = True: r = r + 1

    Dim Steps As Long: Steps = CLng(T_total * STEPS_PER_YEAR)
    Dim dmy As Double
    If Beta_Mat_Steps <> Steps Then
        dmy = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, r_US, r_KRW - r_US, _
                KO_up, KO_dn, S1_0, S2_0, vol1, vol2, corr, Steps, T_total, 50000, K, True)
    End If

    Dim fracs As Variant: fracs = Array(0#, 0.5, 0.9)
    Dim j As Long
    For j = LBound(fracs) To UBound(fracs)
        Dim tfrac As Double: tfrac = fracs(j)
        Dim stp As Long: stp = CLng(tfrac * Steps)
        If stp < 1 Then stp = 1
        If stp > Steps Then stp = Steps
        Dim tau As Double: tau = T_total - stp * (T_total / Steps)
        If tau <= 0.001 Then tau = 0.001

        Dim betaFV As Double: betaFV = QuantoFV_FromBeta(S1_0, S2_0, S1_0, S2_0, stp)
        Dim intr As Double: intr = WorksheetFunction.Max(S1_0 - K, 0#) * S2_0
        If betaFV < intr Then betaFV = intr

        Dim subSteps As Long: subSteps = CLng(tau * STEPS_PER_YEAR)
        If subSteps < 1 Then subSteps = 1
        Dim lsmcFV As Double
        lsmcFV = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, r_US, r_KRW - r_US, _
                    KO_up, KO_dn, S1_0, S2_0, vol1, vol2, corr, subSteps, tau, n_sens, K, False)

        ws.Cells(r, 1) = tfrac: ws.Cells(r, 2) = S1_0: ws.Cells(r, 3) = S2_0
        ws.Cells(r, 4) = betaFV: ws.Cells(r, 5) = lsmcFV
        ws.Cells(r, 6) = Abs(betaFV - lsmcFV)
        ws.Cells(r, 7) = IIf(lsmcFV <> 0#, Abs(betaFV - lsmcFV) / Abs(lsmcFV), 0#)
        r = r + 1
    Next j

    Call RestoreProductionBetaMat
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "FV validation complete -- see CFH_Comparison (row 40+).", vbInformation
    Exit Sub
CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Run_CFH_FV_Validation failed: " & Err.Description, vbCritical
End Sub
