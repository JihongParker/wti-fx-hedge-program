Attribute VB_Name = "CFH_Accounting_revised"
Option Explicit

' =====================================================================
'  VBA CFH ACCOUNTING ENGINE (PRODUCTION v3.1)
'  Refactored to resolve structural timeline defects and matrix symmetry
' =====================================================================

Private Const STEPS_PER_YEAR As Double = 260#

' --- Global Parameters ---
Private gSteps   As Long
Private gVol1    As Double
Private gVol2    As Double
Private gCorr    As Double
Private gHStar0  As Double
Private gFvW0    As Double
Private gFvX0    As Double
Private gT_WTI   As Double
Private gT_FX    As Double

' --- Structure A Ledger Arrays ---
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

' --- Structure B Ledger Arrays ---
Private sB_FVw()       As Double
Private sB_FVx()       As Double
Private sB_CFHRw()     As Double
Private sB_CFHRx()     As Double
Private sB_Ineffw()    As Double
Private sB_Ineffx()    As Double
Private sB_COHamortw() As Double
Private sB_COHamortx() As Double
Private sB_FVTPL()     As Double
Private qB_FVx()       As Double
Private qB_CFHRx()     As Double

' --- Economic Reality Layer Arrays ---
Private sE_EconA()    As Double
Private sE_EconB()    As Double
Private qE_EconA()    As Double
Private qE_EconB()    As Double
Private tA_NakedMax() As Double
Private tB_NakedMax() As Double
Private tEconPathA()  As Double
Private tEconPathB()  As Double

Private Beta_Mat_WTI() As Double
Private Beta_Mat_WTI_Steps As Long

' --- Diagnostics & Analytics ---
Private repSurv As Long, repKO As Long
Private detS_A() As Double, detS_Bw() As Double, detS_Bx() As Double
Private detK_A() As Double, detK_Bw() As Double, detK_Bx() As Double

Private tA_FVpeak()   As Double
Private tA_CFHRend()  As Double
Private tA_IneffTot() As Double
Private tB_FVgross()  As Double
Private tB_CFHRend()  As Double
Private tB_IneffTot() As Double
Private tB_postKO()   As Double
Private tKOflag()     As Long

Private AP_S1() As Double
Private AP_S2() As Double
Private AP_KO() As Long

' =====================================================================
'  1. Pure VBA Mathematical Math Kernels
' =====================================================================

Private Function VMax(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then VMax = a Else VMax = b
End Function

Private Function VMin(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then VMin = a Else VMin = b
End Function

Private Function PureVBA_NormSDist(ByVal x As Double) As Double
    Dim b1 As Double: b1 = 0.31938153
    Dim b2 As Double: b2 = -0.356563782
    Dim b3 As Double: b3 = 1.781477937
    Dim b4 As Double: b4 = -1.821255978
    Dim b5 As Double: b5 = 1.330274429
    Dim p  As Double: p = 0.2316419
    
    Dim absX As Double: absX = Abs(x)
    Dim t As Double, z As Double, res As Double
    
    t = 1# / (1# + p * absX)
    z = (1# / Sqr(2# * 3.14159265358979)) * Exp(-absX * absX / 2#)
    res = 1# - z * (b1 * t + b2 * t ^ 2 + b3 * t ^ 3 + b4 * t ^ 4 + b5 * t ^ 5)
    
    res = VMax(VMin(res, 1#), 0#)
    
    If x >= 0# Then
        PureVBA_NormSDist = res
    Else
        PureVBA_NormSDist = 1# - res
    End If
End Function

Private Function GKCall(ByVal s2 As Double, ByVal K_FX As Double, ByVal vol2 As Double, _
                         ByVal tau As Double, ByVal r_for As Double, ByVal r_dom As Double, ByVal notional As Double) As Double
    If tau <= 1E-06 Or vol2 <= 0# Then
        GKCall = VMax(s2 - K_FX, 0#) * notional: Exit Function
    End If
    Dim d1 As Double: d1 = (Log(s2 / K_FX) + (r_dom - r_for + 0.5 * vol2 ^ 2) * tau) / (vol2 * Sqr(tau))
    Dim d2 As Double: d2 = d1 - vol2 * Sqr(tau)
    GKCall = (s2 * Exp(-r_for * tau) * PureVBA_NormSDist(d1) - _
              K_FX * Exp(-r_dom * tau) * PureVBA_NormSDist(d2)) * notional
End Function

Private Function SignedLowerOf(ByVal cumHI As Double, ByVal cumHD As Double) As Double
    If (cumHI > 0# And cumHD < 0#) Or (cumHI < 0# And cumHD > 0#) Then
        SignedLowerOf = 0#
        Exit Function
    End If
    Dim m As Double: m = Abs(cumHI)
    If Abs(cumHD) < m Then m = Abs(cumHD)
    SignedLowerOf = IIf(cumHI < 0#, -m, m)
End Function

Private Function HStar(ByVal S1 As Double, ByVal s2 As Double, ByVal tau As Double, _
                       ByVal barrels As Double, ByVal rho As Double, ByVal sigma1 As Double, ByVal sigma2 As Double) As Double
    HStar = barrels * S1 * s2 * Exp(rho * sigma1 * sigma2 * tau)
End Function

' =====================================================================
'  2. Main Ledger Orchestrator
' =====================================================================
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

    Dim Lambda As Double: Lambda = wsL.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsL.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsL.Range("B3").Value
    Dim Drift1P  As Double: Drift1P = wsL.Range("B4").Value
    Dim Drift2P  As Double: Drift2P = wsL.Range("B5").Value
    Dim KO_up    As Double: KO_up = wsL.Range("B6").Value
    Dim KO_dn    As Double: KO_dn = wsL.Range("B7").Value
    Dim S1_0     As Double: S1_0 = wsL.Range("B8").Value
    Dim S2_0     As Double: S2_0 = wsL.Range("B9").Value

    gVol1 = wsL.Range("B10").Value
    gVol2 = wsL.Range("B11").Value
    gCorr = wsL.Range("B12").Value

    Dim r_US     As Double: r_US = wsE.Range("B4").Value
    Dim r_KRW    As Double: r_KRW = wsE.Range("B5").Value
    Dim barrels  As Double: barrels = wsE.Range("B9").Value

    gT_WTI = wsE.Range("B15").Value
    gT_FX = wsE.Range("B16").Value
    Dim T_total  As Double: T_total = VMax(gT_WTI, gT_FX)
    If T_total <= 0 Then T_total = wsL.Range("B14").Value

    Dim k        As Double: k = S1_0

    Dim wsIn As Worksheet: Set wsIn = EnsureSheet("CFH_Inputs")
    SeedCFHInputs wsIn
    Dim n_acct As Long: n_acct = CLng(wsIn.Range("B2").Value)
    If n_acct < 1 Then n_acct = 20000

    Dim hedge_steps As Long: hedge_steps = CLng(T_total * STEPS_PER_YEAR)
    Dim dt As Double: dt = T_total / hedge_steps
    gSteps = hedge_steps

    Dim primed As Double
    If Beta_Mat_Steps <> hedge_steps Then
        primed = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, r_US, r_KRW - r_US, _
                    KO_up, KO_dn, S1_0, S2_0, gVol1, gVol2, gCorr, _
                    hedge_steps, T_total, 50000, k, True)
    End If
    Dim Base_Premium As Double: Base_Premium = wsL.Range("J2").Value
    If Base_Premium <= 0# And primed > 0# Then Base_Premium = primed

    Dim TV0 As Double: TV0 = Base_Premium * barrels
    If TV0 <= 0# Then GoTo CleanFail

    Dim Steps_WTI As Long: Steps_WTI = CLng(gT_WTI * STEPS_PER_YEAR)
    If Steps_WTI < 1 Then Steps_WTI = 1
    gFvW0 = Calc_WTI_KO_LSMC_Price(Lambda, JumpMean, JumpVol, r_US, KO_up, KO_dn, _
                                    S1_0, S2_0, gVol1, Steps_WTI, gT_WTI, 50000, k, True) * barrels

    Call Build_AcctPaths(n_acct, hedge_steps, S1_0, S2_0, gVol1, gVol2, gCorr, _
            Drift1P, Drift2P, Lambda, JumpMean, JumpVol, KO_up, KO_dn, T_total)

    Dim F0_WTI As Double: F0_WTI = S1_0 * Exp(r_US * gT_WTI)
    Dim G0_FX  As Double: G0_FX = S2_0 * Exp((r_KRW - r_US) * gT_FX)
    Dim N_FX   As Double: N_FX = barrels * F0_WTI

    gHStar0 = HStar(S1_0, S2_0, T_total, barrels, gCorr, gVol1, gVol2)
    gFvX0 = GKCall(S2_0, G0_FX, gVol2, gT_FX, r_US, r_KRW, N_FX)

    repSurv = 0: repKO = 0
    Dim s As Long
    For s = 1 To n_acct
        If repSurv = 0 And AP_KO(s) = 0 Then repSurv = s
        If repKO = 0 And AP_KO(s) > 0 Then repKO = s
    Next s
    If repSurv = 0 Then repSurv = 1
    If repKO = 0 Then repKO = 1

    AllocAccumulators hedge_steps, n_acct

    Dim sim As Long
    For sim = 1 To n_acct
        WalkPath sim, hedge_steps, dt, S1_0, S2_0, k, barrels, _
                 r_US, r_KRW, F0_WTI, G0_FX, N_FX, TV0, _
                 (sim = repSurv), (sim = repKO)
    Next sim

    WriteLedgerA hedge_steps, n_acct, dt
    WriteLedgerB hedge_steps, n_acct, dt
    WriteSFPProforma hedge_steps, n_acct, dt, barrels, Base_Premium
    WriteEconomicLedger hedge_steps, n_acct, dt
    WriteComparison hedge_steps, n_acct, Base_Premium, barrels, TV0, N_FX, S1_0, S2_0, k

    Call RestoreProductionBetaMat

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "IFRS 9 CFH Engine v3.1 Run Completed Successfully.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Engine Crash: " & Err.Description, vbCritical
End Sub

' =====================================================================
'  3. Path-by-Path Evolution and Accounting Module
' =====================================================================
Private Sub WalkPath(ByVal sim As Long, ByVal Steps As Long, ByVal dt As Double, _
                     ByVal S1_0 As Double, ByVal S2_0 As Double, ByVal k As Double, _
                     ByVal barrels As Double, ByVal r_US As Double, ByVal r_KRW As Double, _
                     ByVal F0_WTI As Double, ByVal G0_FX As Double, ByVal N_FX As Double, _
                     ByVal TV0 As Double, ByVal isSurv As Boolean, ByVal isKOrep As Boolean)

    Dim wtiCFHR_Frz As Double, wtiIneff_Frz As Double
    Dim fxCFHR_Frz  As Double, fxIneff_Frz  As Double
    Dim bCFHRw_KO   As Double, bCFHRx_KO    As Double

    Dim T_total As Double: T_total = Steps * dt
    Dim koStep As Long: koStep = AP_KO(sim)

    Dim aIntr_0 As Double: aIntr_0 = VMax(S1_0 - k, 0#) * S2_0 * barrels
    Dim Intr_WTI_0 As Double: Intr_WTI_0 = VMax(S1_0 - k, 0#) * S2_0 * barrels
    Dim Intr_FX_0 As Double: Intr_FX_0 = VMax(S2_0 - G0_FX, 0#) * N_FX

    ' Structure A states
    Dim aCFHR As Double: aCFHR = 0#
    Dim aIneff As Double: aIneff = 0#
    Dim aCOHamort As Double: aCOHamort = 0#
    Dim coh_step As Double: coh_step = TV0 / Steps
    Dim aCFHR_frozen As Double, aCOH_frozen As Double
    Dim aDead As Boolean: aDead = False

    ' Structure B states
    Dim bCFHRw As Double: bCFHRw = 0#
    Dim bIneffw As Double: bIneffw = 0#
    Dim bCOHamortw As Double: bCOHamortw = 0#
    Dim TV_W0 As Double: TV_W0 = VMax(gFvW0 - Intr_WTI_0, 0#)
    Dim coh_step_w As Double: coh_step_w = TV_W0 / Steps
    Dim bCOH_frozen_w As Double

    Dim bCFHRx As Double: bCFHRx = 0#
    Dim bIneffx As Double: bIneffx = 0#
    Dim bCOHamortx As Double: bCOHamortx = 0#
    Dim TV_X0 As Double: TV_X0 = VMax(gFvX0 - Intr_FX_0, 0#)
    Dim coh_step_x As Double: coh_step_x = TV_X0 / Steps
    Dim bCOH_frozen_x As Double

    Dim bFVTPL As Double: bFVTPL = 0#
    Dim bDiscont As Boolean: bDiscont = False

    Dim wtiExpiredPrev As Boolean: wtiExpiredPrev = False
    Dim wtiFV_Frozen As Double
    Dim fxExpiredPrev As Boolean: fxExpiredPrev = False
    Dim fxFV_Frozen As Double

    Dim aFV_prev As Double: aFV_prev = TV0
    Dim bFVw_prev As Double: bFVw_prev = gFvW0
    Dim bFVx_prev As Double: bFVx_prev = gFvX0
    Dim Phys_prev As Double: Phys_prev = barrels * S1_0 * S2_0

    Dim fvWTI As Double, cumHIx As Double
    Dim fvWprev As Double: fvWprev = gFvW0
    Dim fvXprev As Double: fvXprev = gFvX0

    Dim pathCumEconA As Double: pathCumEconA = 0#
    Dim pathCumEconB As Double: pathCumEconB = 0#

    Dim stp As Long, S1 As Double, s2 As Double, tau_WTI As Double, tau_FX As Double
    Dim tt As Double, alive As Boolean, koHere As Boolean
    Dim wtiExpiredCur As Boolean, fxExpiredCur As Boolean

    For stp = 0 To Steps
        If stp = 0 Then
            S1 = S1_0: s2 = S2_0
        Else
            S1 = AP_S1(sim, stp): s2 = AP_S2(sim, stp)
        End If
        tt = stp * dt
        tau_WTI = VMax(0#, gT_WTI - tt)
        tau_FX = VMax(0#, gT_FX - tt)
        wtiExpiredCur = (tau_WTI <= 1E-12)
        fxExpiredCur = (tau_FX <= 1E-12)
        alive = (koStep = 0) Or (stp < koStep)
        koHere = (koStep > 0 And stp = koStep)

        ' Global Freeze Layer (Path-independent structural tracking)
        If fxExpiredCur Then
            If Not fxExpiredPrev Then
                fxFV_Frozen = VMax(s2 - G0_FX, 0#) * N_FX
                fxExpiredPrev = True
            End If
            cumHIx = fxFV_Frozen
        Else
            cumHIx = GKCall(s2, G0_FX, gVol2, tau_FX, r_US, r_KRW, N_FX)
        End If

        If wtiExpiredCur Then
            If Not wtiExpiredPrev Then
                Dim wtiFinalBeta As Double
                wtiFinalBeta = WTI_KO_Call_BetaFV(S1, S1_0, Beta_Mat_WTI_Steps, barrels, koStep)
                If wtiFinalBeta < VMax(S1 - k, 0#) * S2_0 * barrels Then
                    wtiFinalBeta = VMax(S1 - k, 0#) * S2_0 * barrels
                End If
                wtiFV_Frozen = wtiFinalBeta
                wtiExpiredPrev = True
            End If
        End If

        ' --- 3.1 Structure A Layer ---
        Dim aFV As Double, aIntr As Double, aTV As Double
        
        If koHere And Not aDead Then
            Dim intrUnit_ko As Double: intrUnit_ko = VMax(S1 - k, 0#) * s2
            Dim vUnit_ko As Double: vUnit_ko = QuantoFV_FromBeta(S1, s2, S1_0, S2_0, stp)
            If vUnit_ko < intrUnit_ko Then vUnit_ko = intrUnit_ko
            
            aFV = vUnit_ko * barrels
            aIntr = intrUnit_ko * barrels
            aTV = aFV - aIntr
            
            aCFHR_frozen = aCFHR
            Dim a_residual_TV As Double: a_residual_TV = VMax(TV0 - aCOHamort, 0#)
            
            aIneff = (aIntr - aIntr_0) - aCFHR - a_residual_TV
            aCOHamort = aCOHamort + a_residual_TV
            aCOH_frozen = 0#
            
            aDead = True
        ElseIf aDead Then
            aFV = 0#: aIntr = 0#: aTV = 0#
        Else
            Dim intrUnit As Double: intrUnit = VMax(S1 - k, 0#) * s2
            Dim vUnit As Double
            If stp = 0 Then
                vUnit = TV0 / barrels
            Else
                vUnit = QuantoFV_FromBeta(S1, s2, S1_0, S2_0, stp)
            End If
            If vUnit < intrUnit Then vUnit = intrUnit
            aFV = vUnit * barrels: aIntr = intrUnit * barrels: aTV = aFV - aIntr
        End If

        If stp > 0 And Not koHere Then
            If Not aDead Then
                Dim HStar_t As Double: HStar_t = HStar(S1, s2, T_total - tt, barrels, gCorr, gVol1, gVol2)
                Dim v1n As Double: v1n = S1 / S1_0
                Dim v2n As Double: v2n = s2 / S2_0
                Dim rawH As Double: rawH = ((Beta_Mat(stp, 1) + 2# * Beta_Mat(stp, 3) * v1n + Beta_Mat(stp, 5) * v2n) / S1_0) / s2
                Dim h_t As Double
             If rawH < 0# Then
    h_t = 0#
ElseIf rawH > 1# Then
    h_t = 1#
Else
    h_t = rawH
End If

                aCFHR = SignedLowerOf(aIntr - aIntr_0, h_t * (HStar_t - gHStar0))
                aIneff = (aIntr - aIntr_0) - aCFHR
                aCOHamort = aCOHamort + coh_step
            Else
                aCFHR = aCFHR_frozen
            End If
        End If
        
        Dim aCOH As Double: aCOH = IIf(Not aDead, VMax(aTV - (TV0 - aCOHamort), 0#), aCOH_frozen)

        ' --- 3.2 Structure B Layer ---
        Dim Intr_WTI As Double: Intr_WTI = VMax(S1 - k, 0#) * S2_0 * barrels
        Dim Intr_FX As Double: Intr_FX = VMax(s2 - G0_FX, 0#) * N_FX

        If Not bDiscont Then
            If koHere Then
                bDiscont = True
                bCFHRw_KO = bCFHRw
                bCFHRx_KO = bCFHRx

                Dim residual_TV_W As Double: residual_TV_W = VMax(TV_W0 - bCOHamortw, 0#)
                bIneffw = bIneffw - residual_TV_W
                bCOHamortw = bCOHamortw + residual_TV_W
                bCOH_frozen_w = 0#
                bCOH_frozen_x = VMax((cumHIx - Intr_FX) - (TV_X0 - bCOHamortx), 0#)
            Else
                If wtiExpiredCur Then
                    bCFHRw = wtiCFHR_Frz: bIneffw = wtiIneff_Frz
                Else
                    bCFHRw = SignedLowerOf(Intr_WTI - Intr_WTI_0, barrels * (S1 - S1_0) * S2_0)
                    bIneffw = (Intr_WTI - Intr_WTI_0) - bCFHRw
                    bCOHamortw = bCOHamortw + coh_step_w
                    If wtiExpiredCur Then wtiCFHR_Frz = bCFHRw: wtiIneff_Frz = bIneffw
                End If

                If fxExpiredCur Then
                    bCFHRx = fxCFHR_Frz: bIneffx = fxIneff_Frz
                Else
                    Dim fxHyp As Double: fxHyp = VMax(s2 - G0_FX, 0#) * barrels * S1 * Exp(r_US * tau_WTI)
                    bCFHRx = SignedLowerOf(Intr_FX - Intr_FX_0, fxHyp - Intr_FX_0)
                    bIneffx = (Intr_FX - Intr_FX_0) - bCFHRx
                    bCOHamortx = bCOHamortx + coh_step_x
                    If fxExpiredCur Then fxCFHR_Frz = bCFHRx: fxIneff_Frz = bIneffx
                End If
            End If
        Else
            bFVTPL = bFVTPL + (fvWTI - fvWprev) + (cumHIx - fvXprev)
            bCFHRw = bCFHRw_KO
            bCFHRx = bCFHRx_KO
        End If

        If bDiscont Then
            fvWTI = 0#
        ElseIf wtiExpiredCur Then
            fvWTI = wtiFV_Frozen
        Else
            Dim wtiStp As Long
            If gT_WTI > 0 And Beta_Mat_WTI_Steps > 0 Then
                wtiStp = CLng((tt / gT_WTI) * CDbl(Beta_Mat_WTI_Steps))
                If wtiStp < 1 Then wtiStp = 1
                If wtiStp > Beta_Mat_WTI_Steps Then wtiStp = Beta_Mat_WTI_Steps
            Else
                wtiStp = stp
            End If
            fvWTI = WTI_KO_Call_BetaFV(S1, S1_0, wtiStp, barrels, koStep)
            If fvWTI < Intr_WTI Then fvWTI = Intr_WTI
        End If

        Dim bCOHw As Double: bCOHw = IIf(bDiscont, bCOH_frozen_w, VMax((fvWTI - Intr_WTI) - (TV_W0 - bCOHamortw), 0#))
        Dim bCOHx As Double: bCOHx = IIf(bDiscont, bCOH_frozen_x, VMax((cumHIx - Intr_FX) - (TV_X0 - bCOHamortx), 0#))

        ' --- 3.3 Economic Integration Layer ---
        Dim Phys As Double: Phys = barrels * S1 * s2
        If stp > 0 Then
            Dim dPhys As Double: dPhys = Phys - Phys_prev
            Dim dV_A_step As Double: dV_A_step = aFV - aFV_prev
            Dim dV_B_step As Double: dV_B_step = (fvWTI + cumHIx) - (bFVw_prev + bFVx_prev)

            Dim dEconA As Double: dEconA = dV_A_step - dPhys
            Dim dEconB As Double: dEconB = dV_B_step - dPhys

            pathCumEconA = pathCumEconA + dEconA
            pathCumEconB = pathCumEconB + dEconB

            sE_EconA(stp) = sE_EconA(stp) + dEconA: qE_EconA(stp) = qE_EconA(stp) + dEconA * dEconA
            sE_EconB(stp) = sE_EconB(stp) + dEconB: qE_EconB(stp) = qE_EconB(stp) + dEconB * dEconB

            If bDiscont Or koHere Then
                Dim NakedA As Double: NakedA = Phys
                Dim NakedB As Double: NakedB = VMax(Phys - cumHIx, 0#)
                If NakedA > tA_NakedMax(sim) Then tA_NakedMax(sim) = NakedA
                If NakedB > tB_NakedMax(sim) Then tB_NakedMax(sim) = NakedB
            End If
        End If

        aFV_prev = aFV: bFVw_prev = fvWTI: bFVx_prev = cumHIx: Phys_prev = Phys
        fvWprev = fvWTI: fvXprev = cumHIx

        ' Ledger Accumulation
        sA_FV(stp) = sA_FV(stp) + aFV:   qA_FV(stp) = qA_FV(stp) + aFV * aFV
        sA_Intr(stp) = sA_Intr(stp) + aIntr: sA_TV(stp) = sA_TV(stp) + aTV
        sA_CFHR(stp) = sA_CFHR(stp) + aCFHR: qA_CFHR(stp) = qA_CFHR(stp) + aCFHR * aCFHR
        sA_COH(stp) = sA_COH(stp) + aCOH: sA_COHamort(stp) = sA_COHamort(stp) + aCOHamort
        sA_Ineff(stp) = sA_Ineff(stp) + aIneff
        If alive Or koHere Then sA_Alive(stp) = sA_Alive(stp) + 1#
        
        sB_FVw(stp) = sB_FVw(stp) + fvWTI
        sB_FVx(stp) = sB_FVx(stp) + cumHIx: qB_FVx(stp) = qB_FVx(stp) + cumHIx * cumHIx
        sB_CFHRw(stp) = sB_CFHRw(stp) + bCFHRw + bCOHw
        sB_CFHRx(stp) = sB_CFHRx(stp) + bCFHRx + bCOHx
        sB_Ineffw(stp) = sB_Ineffw(stp) + bIneffw: sB_Ineffx(stp) = sB_Ineffx(stp) + bIneffx
        sB_COHamortw(stp) = sB_COHamortw(stp) + bCOHamortw
        sB_COHamortx(stp) = sB_COHamortx(stp) + bCOHamortx
        sB_FVTPL(stp) = sB_FVTPL(stp) + bFVTPL
        
        If isSurv Then
            detS_A(stp, 0) = aFV: detS_A(stp, 3) = aCFHR: detS_A(stp, 5) = aIneff
            detS_Bw(stp, 0) = fvWTI: detS_Bw(stp, 1) = bCFHRw: detS_Bw(stp, 2) = bIneffw
            detS_Bx(stp, 0) = cumHIx: detS_Bx(stp, 1) = bCFHRx: detS_Bx(stp, 2) = bIneffx
        End If
        If isKOrep Then
            detK_A(stp, 0) = aFV: detK_A(stp, 3) = aCFHR: detK_A(stp, 5) = aIneff
            detK_Bw(stp, 0) = fvWTI: detK_Bw(stp, 1) = bCFHRw: detK_Bw(stp, 2) = bIneffw
            detK_Bx(stp, 0) = cumHIx: detK_Bx(stp, 1) = bCFHRx: detK_Bx(stp, 2) = bIneffx
        End If
        If aFV > tA_FVpeak(sim) Then tA_FVpeak(sim) = aFV
    Next stp

    tEconPathA(sim) = pathCumEconA
    tEconPathB(sim) = pathCumEconB

    tA_CFHRend(sim) = aCFHR: tA_IneffTot(sim) = aIneff
    tB_FVgross(sim) = Abs(fvWprev) + Abs(fvXprev)
    tB_CFHRend(sim) = bCFHRw + bCFHRx: tB_IneffTot(sim) = bIneffw + bIneffx
    tB_postKO(sim) = bFVTPL: tKOflag(sim) = IIf(koStep > 0, 1, 0)
End Sub

' =====================================================================
'  4. Pricing and Calibration Surfaces
' =====================================================================
Private Function QuantoFV_FromBeta(ByVal S1 As Double, ByVal s2 As Double, _
                                   ByVal S1_0 As Double, ByVal S2_0 As Double, _
                                   ByVal stp As Long) As Double
    Dim v1 As Double: v1 = S1 / S1_0
    Dim v2 As Double: v2 = s2 / S2_0
    Dim vb As Double
    vb = Beta_Mat(stp, 0) + Beta_Mat(stp, 1) * v1 + Beta_Mat(stp, 2) * v2 _
       + Beta_Mat(stp, 3) * v1 ^ 2 + Beta_Mat(stp, 4) * v2 ^ 2 _
       + Beta_Mat(stp, 5) * v1 * v2
    If vb < 0# Then vb = 0#
    QuantoFV_FromBeta = vb
End Function

Private Function WTI_KO_Call_BetaFV(ByVal S1 As Double, ByVal S1_0 As Double, _
                                    ByVal stp As Long, ByVal barrels As Double, ByVal koStep As Long) As Double
    If stp = 0 Then
        WTI_KO_Call_BetaFV = gFvW0
        Exit Function
    End If
    If (koStep > 0 And stp >= koStep) Or stp < 1 Or stp > Beta_Mat_WTI_Steps Then
        WTI_KO_Call_BetaFV = 0#
        Exit Function
    End If
    Dim v1 As Double: v1 = S1 / S1_0
    Dim vb As Double
    vb = Beta_Mat_WTI(stp, 0) + Beta_Mat_WTI(stp, 1) * v1 + Beta_Mat_WTI(stp, 3) * v1 ^ 2
    If vb < 0# Then vb = 0#
    WTI_KO_Call_BetaFV = vb * barrels
End Function

Private Function Calc_WTI_KO_LSMC_Price(ByVal Lambda As Double, ByVal JumpMean As Double, ByVal JumpVol As Double, _
                                         ByVal r_US As Double, _
                                         ByVal KOUpper As Double, ByVal KOLower As Double, _
                                         ByVal S1_init As Double, ByVal S2_init As Double, _
                                         ByVal vol1 As Double, _
                                         ByVal Steps As Long, ByVal t As Double, _
                                         ByVal n_paths As Long, ByVal k As Double, _
                                         Optional ByVal bStoreBeta As Boolean = False) As Double

    Dim dt As Double: dt = t / Steps
    Dim sqdt As Double: sqdt = Sqr(dt)
    Dim kappa As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt
    Dim disc As Double: disc = Exp(-r_US * dt)

    Dim S1_path() As Double, alive() As Boolean
    ReDim S1_path(1 To n_paths, 0 To Steps)
    ReDim alive(1 To n_paths, 0 To Steps)

    If bStoreBeta Then
        ReDim Beta_Mat_WTI(1 To Steps, 0 To 5)
        Beta_Mat_WTI_Steps = Steps
    End If

    Rnd -1
    Randomize 20240101 + 1

    Dim p As Long, i As Long
    For p = 1 To n_paths
        S1_path(p, 0) = S1_init
        alive(p, 0) = True
        Dim S1 As Double: S1 = S1_init
        Dim isAlive As Boolean: isAlive = True
        Dim S1_prev As Double: S1_prev = S1_init

        For i = 1 To Steps
            If isAlive Then
                S1_prev = S1
                Dim z1 As Double: z1 = GetNormal()
                Dim nJ As Long: nJ = GetPoisson(Lambda * dt)
                Dim jSum As Double: jSum = JumpSum(nJ, JumpMean, JumpVol)

                S1 = S1 * Exp((r_US - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt + vol1 * sqdt * z1 + jSum)

                If S1 >= KOUpper Or S1 <= KOLower Then
                    isAlive = False
                ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then
                    Dim p_up As Double, p_dn As Double
                    p_up = Exp(-2# * Log(KOUpper / S1_prev) * Log(KOUpper / S1) / lnVar)
                    p_dn = Exp(-2# * Log(S1_prev / KOLower) * Log(S1 / KOLower) / lnVar)
                    If GetUniform() < p_up Then isAlive = False
                    If isAlive And GetUniform() < p_dn Then isAlive = False
                End If
            End If
            S1_path(p, i) = S1
            alive(p, i) = isAlive
        Next i
    Next p

    Dim CF() As Double: ReDim CF(1 To n_paths)
    For p = 1 To n_paths
        If alive(p, Steps) Then
            CF(p) = VMax(S1_path(p, Steps) - k, 0) * S2_init
        Else
            CF(p) = 0
        End If
    Next p

    For i = Steps - 1 To 1 Step -1
        For p = 1 To n_paths
            CF(p) = CF(p) * disc
        Next p

        Dim itmCount As Long: itmCount = 0
        For p = 1 To n_paths
            If alive(p, i) And (S1_path(p, i) - k) > 0 Then itmCount = itmCount + 1
        Next p

        If itmCount > 10 Then
            Dim X_Mat() As Double, Y_Vec() As Double
            ReDim X_Mat(1 To itmCount, 1 To 2)
            ReDim Y_Vec(1 To itmCount, 1 To 1)

            Dim itmRow As Long: itmRow = 0
            For p = 1 To n_paths
                If alive(p, i) And (S1_path(p, i) - k) > 0 Then
                    itmRow = itmRow + 1
                    Dim v1 As Double: v1 = S1_path(p, i) / k
                    X_Mat(itmRow, 1) = v1
                    X_Mat(itmRow, 2) = v1 ^ 2
                    Y_Vec(itmRow, 1) = CF(p)
                End If
            Next p

            Dim rawOut As Variant
            On Error Resume Next
            rawOut = Application.WorksheetFunction.LinEst(Y_Vec, X_Mat, True, False)
            Dim errNum As Long: errNum = Err.Number
            On Error GoTo 0

            Dim Coeff(0 To 2) As Double
            If errNum = 0 And IsArray(rawOut) Then
                On Error Resume Next
                Coeff(2) = rawOut(1)
                Coeff(1) = rawOut(2)
                Coeff(0) = rawOut(3)
                
                If Err.Number <> 0 Then
                    Err.Clear
                    Coeff(2) = rawOut(1, 1)
                    Coeff(1) = rawOut(1, 2)
                    Coeff(0) = rawOut(1, 3)
                End If
                On Error GoTo 0
            End If

            If bStoreBeta Then
                Beta_Mat_WTI(i, 0) = Coeff(0)
                Beta_Mat_WTI(i, 1) = Coeff(1)
                Beta_Mat_WTI(i, 2) = 0#
                Beta_Mat_WTI(i, 3) = Coeff(2)
                Beta_Mat_WTI(i, 4) = 0#
                Beta_Mat_WTI(i, 5) = 0#
            End If

            For p = 1 To n_paths
                If alive(p, i) And (S1_path(p, i) - k) > 0 Then
                    Dim curS1 As Double: curS1 = S1_path(p, i) / k
                    Dim intrinsic As Double: intrinsic = (S1_path(p, i) - k) * S2_init
                    Dim cont_val As Double: cont_val = Coeff(0) + Coeff(1) * curS1 + Coeff(2) * (curS1 ^ 2)
                    If intrinsic > cont_val Then CF(p) = intrinsic
                End If
            Next p
        Else
            If bStoreBeta And i < Steps - 1 Then
                Beta_Mat_WTI(i, 0) = Beta_Mat_WTI(i + 1, 0)
                Beta_Mat_WTI(i, 1) = Beta_Mat_WTI(i + 1, 1)
                Beta_Mat_WTI(i, 2) = 0#
                Beta_Mat_WTI(i, 3) = Beta_Mat_WTI(i + 1, 3)
                Beta_Mat_WTI(i, 4) = 0#
                Beta_Mat_WTI(i, 5) = 0#
            End If
        End If
    Next i

    Dim totalVal As Double
    For p = 1 To n_paths
        totalVal = totalVal + CF(p) * disc
    Next p

    Calc_WTI_KO_LSMC_Price = totalVal / n_paths
End Function

Private Function GetUniform() As Double
    GetUniform = Rnd
End Function

' =====================================================================
'  5. Output Reporting Modules
' =====================================================================
Private Sub WriteLedgerA(ByVal Steps As Long, ByVal n As Long, ByVal dt As Double)
    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_A_Ledger")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Structure A -- Single Combined Quanto KO CFH (firm-scale KRW)"
    ws.Range("A1").Font.Bold = True
    Dim hdr As Variant
    hdr = Array("Step", "t(yr)", "MeanFV", "StdFV", "MeanIntrinsic", "MeanTimeValue", _
                "MeanCFHR(OCI)", "StdCFHR", "MeanCOH(OCI)", "CumCOHamort(P&L)", "CumIneff(P&L)", "AliveFrac")
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Value = hdr: ws.Range("A3").Resize(1, UBound(hdr) + 1).Font.Bold = True

    Dim out() As Variant: ReDim out(0 To Steps, 0 To UBound(hdr))
    Dim i As Long
    For i = 0 To Steps
        out(i, 0) = i: out(i, 1) = i * dt
        out(i, 2) = sA_FV(i) / n: out(i, 3) = StdFrom(sA_FV(i), qA_FV(i), n)
        out(i, 4) = sA_Intr(i) / n: out(i, 5) = sA_TV(i) / n
        out(i, 6) = sA_CFHR(i) / n: out(i, 7) = StdFrom(sA_CFHR(i), qA_CFHR(i), n)
        out(i, 8) = sA_COH(i) / n: out(i, 9) = sA_COHamort(i) / n
        out(i, 10) = sA_Ineff(i) / n: out(i, 11) = sA_Alive(i) / n
    Next i
    ws.Range("A4").Resize(Steps + 1, UBound(hdr) + 1).Value = out: ws.Columns.AutoFit
End Sub

Private Sub WriteLedgerB(ByVal Steps As Long, ByVal n As Long, ByVal dt As Double)
    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_B_Ledger_WTI")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Structure B Layer 1 -- WTI KO Option CFH (FX-fixed at S2_0)"
    ws.Range("A1").Font.Bold = True
    Dim h1 As Variant: h1 = Array("Step", "t(yr)", "MeanFV", "MeanCFHR(OCI)", "CumIneff(P&L)")
    ws.Range("A3").Resize(1, UBound(h1) + 1).Value = h1: ws.Range("A3").Resize(1, UBound(h1) + 1).Font.Bold = True
    Dim o1() As Variant: ReDim o1(0 To Steps, 0 To UBound(h1))
    Dim i As Long
    For i = 0 To Steps
        o1(i, 0) = i: o1(i, 1) = i * dt
        o1(i, 2) = sB_FVw(i) / n: o1(i, 3) = sB_CFHRw(i) / n: o1(i, 4) = sB_Ineffw(i) / n
    Next i
    ws.Range("A4").Resize(Steps + 1, UBound(h1) + 1).Value = o1: ws.Columns.AutoFit

    Dim ws2 As Worksheet: Set ws2 = EnsureSheet("CFH_B_Ledger_FX")
    ws2.Cells.ClearContents
    ws2.Range("A1").Value = "Structure B Layer 2 -- FX GK Option CFH (FIXED Notional Mismatch)"
    ws2.Range("A1").Font.Bold = True
    Dim h2 As Variant: h2 = Array("Step", "t(yr)", "MeanFV", "StdFV", "MeanCFHR(OCI)", "StdCFHR", "CumIneff(P&L)", "CumPostKO_FVTPL(P&L)")
    ws2.Range("A3").Resize(1, UBound(h2) + 1).Value = h2: ws2.Range("A3").Resize(1, UBound(h2) + 1).Font.Bold = True
    Dim o2() As Variant: ReDim o2(0 To Steps, 0 To UBound(h2))
    For i = 0 To Steps
        o2(i, 0) = i: o2(i, 1) = i * dt
        o2(i, 2) = sB_FVx(i) / n: o2(i, 3) = StdFrom(sB_FVx(i), qB_FVx(i), n)
        o2(i, 4) = sB_CFHRx(i) / n: o2(i, 5) = StdFrom(sB_CFHRx(i), qB_CFHRx(i), n)
        o2(i, 6) = sB_Ineffx(i) / n: o2(i, 7) = sB_FVTPL(i) / n
    Next i
    ws2.Range("A4").Resize(Steps + 1, UBound(h2) + 1).Value = o2: ws2.Columns.AutoFit
End Sub

Private Sub WriteSFPProforma(ByVal Steps As Long, ByVal n As Long, ByVal dt As Double, ByVal barrels As Double, ByVal Base_Premium As Double)
    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_SFP_Proforma")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "SFP Pro-forma (cross-path mean, KRW): Structure A vs Structure B"
    ws.Range("A1").Font.Bold = True
    Dim hdr As Variant
    hdr = Array("Step", "t(yr)", "A_Derivative", "A_OCI(CFHR+COH)", "A_RetainedEarnings", "A_#DerivLines", _
                "B_Deriv_WTI", "B_Deriv_FX", "B_OCI(CFHR_W+CFHR_X)", "B_RetainedEarnings", "B_#DerivLines")
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Value = hdr: ws.Range("A3").Resize(1, UBound(hdr) + 1).Font.Bold = True
    Dim out() As Variant: ReDim out(0 To Steps, 0 To UBound(hdr))
    Dim i As Long
    For i = 0 To Steps
        out(i, 0) = i: out(i, 1) = i * dt
        out(i, 2) = sA_FV(i) / n: out(i, 3) = (sA_CFHR(i) + sA_COH(i)) / n
        
        ' [���� ���� 9 ����] �������� ���� ������(Asset = Equity + RE)�� ���������� ������ ���� ����
        out(i, 4) = (sA_FV(i) - sA_CFHR(i) - sA_COH(i)) / n: out(i, 5) = 1
        
        out(i, 6) = sB_FVw(i) / n: out(i, 7) = sB_FVx(i) / n
        out(i, 8) = (sB_CFHRw(i) + sB_CFHRx(i)) / n
        out(i, 9) = (sB_FVw(i) + sB_FVx(i) - sB_CFHRw(i) - sB_CFHRx(i)) / n
        out(i, 10) = 2
    Next i
    ws.Range("A4").Resize(Steps + 1, UBound(hdr) + 1).Value = out: ws.Columns.AutoFit
End Sub

Private Sub WriteEconomicLedger(ByVal Steps As Long, ByVal n As Long, ByVal dt As Double)
    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_Economic")
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Structure A vs B Economic Residual & Post-KO Risk Analytics"
    ws.Range("A1").Font.Bold = True
    Dim hdr As Variant: hdr = Array("Step", "t(yr)", "Mean_EconA", "Std_EconA(sigma_econA)", "Mean_EconB", "Std_EconB(sigma_econB)")
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Value = hdr: ws.Range("A3").Resize(1, UBound(hdr) + 1).Font.Bold = True
    Dim out() As Variant: ReDim out(0 To Steps, 0 To UBound(hdr))
    Dim i As Long
    For i = 0 To Steps
        out(i, 0) = i: out(i, 1) = i * dt
        out(i, 2) = sE_EconA(i) / n: out(i, 3) = StdFrom(sE_EconA(i), qE_EconA(i), n)
        out(i, 4) = sE_EconB(i) / n: out(i, 5) = StdFrom(sE_EconB(i), qE_EconB(i), n)
    Next i
    ws.Range("A4").Resize(Steps + 1, UBound(hdr) + 1).Value = out: ws.Columns.AutoFit
End Sub

Private Sub WriteComparison(ByVal Steps As Long, ByVal n As Long, ByVal Base_Premium As Double, _
                            ByVal barrels As Double, ByVal TV0 As Double, ByVal N_FX As Double, _
                            ByVal S1_0 As Double, ByVal S2_0 As Double, ByVal k As Double)
    Call Run_CFH_FV_Validation(S1_0, S2_0, k)

    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_Comparison")
    ws.Range("A1:G20").ClearContents
    ws.Range("A1").Value = "CFH Matrix Comparison Matrix -- Spec v3.1"
    ws.Range("A1").Font.Bold = True

    Dim i As Long
    Dim r As Long: r = 3
    ws.Cells(r, 1) = "Inception Check": ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2) = "Premium Outflow Check": ws.Cells(r, 3) = TV0
    ws.Cells(r, 4) = "A_FV(0) Mean": ws.Cells(r, 5) = sA_FV(0) / n
    ws.Cells(r, 6) = "B1_FV(0)": ws.Cells(r, 7) = gFvW0
    ws.Cells(r, 8) = "B2_FV(0)": ws.Cells(r, 9) = gFvX0
    r = r + 2

    ws.Cells(r, 1) = "H1: Cumulative Ineffectiveness Comparison (Tautology Check)"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1) = "Structure A Mean |Cum Ineff|": ws.Cells(r, 2) = MeanAbsArr(tA_IneffTot, n)
    r = r + 1
    ws.Cells(r, 1) = "Structure B Mean |Cum Ineff|": ws.Cells(r, 2) = MeanAbsArr(tB_IneffTot, n)
    r = r + 2

    ws.Cells(r, 1) = "H2: Full Horizon Economic Residual Volatility (sigma_econ)"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1) = "Structure A Full Horizon sigma_econ": ws.Cells(r, 2) = Application.WorksheetFunction.StDev(tEconPathA)
    r = r + 1
    ws.Cells(r, 1) = "Structure B Full Horizon sigma_econ": ws.Cells(r, 2) = Application.WorksheetFunction.StDev(tEconPathB)
    r = r + 2

    ws.Cells(r, 1) = "H3: Post-KO Naked Exposure Tail Risk (VaR99) & OCI Reclassification"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    Dim koCount As Long: koCount = 0
    Dim sumReclassifiedA As Double: sumReclassifiedA = 0#
    For i = 1 To n
        If tKOflag(i) = 1 Then
            koCount = koCount + 1
            sumReclassifiedA = sumReclassifiedA + tA_CFHRend(i)
        End If
    Next i
    ws.Cells(r, 1) = "Knock-Out Probability": ws.Cells(r, 2) = koCount / n
    r = r + 1
    ws.Cells(r, 1) = "Structure A Naked Exposure VaR99": ws.Cells(r, 2) = Application.Percentile(tA_NakedMax, 0.99)
    r = r + 1
    ws.Cells(r, 1) = "Structure B Naked Exposure VaR99": ws.Cells(r, 2) = Application.Percentile(tB_NakedMax, 0.99)
    r = r + 1
    ws.Cells(r, 1) = "A Structure OCI->P&L Reclassification Amount": ws.Cells(r, 2) = sumReclassifiedA / n
    r = r + 2

    ws.Cells(r, 1) = "H4: Hedging and Accounting Efficiency Report"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1) = "Number of derivative lines (A vs B)": ws.Cells(r, 2) = "1": ws.Cells(r, 3) = "2 (WTI + FX)"
    ws.Columns("A:I").AutoFit
End Sub

Private Sub Run_CFH_FV_Validation(ByVal S1_0 As Double, ByVal S2_0 As Double, ByVal k As Double)
    Dim wsL As Worksheet: Set wsL = Sheets("LSMC")
    Dim wsE As Worksheet: Set wsE = Sheets("Encoding")
    Dim ws As Worksheet: Set ws = EnsureSheet("CFH_Comparison")

    Dim Lambda As Double: Lambda = wsL.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsL.Range("B2").Value
    Dim JumpVol As Double: JumpVol = wsL.Range("B3").Value
    Dim KO_up As Double: KO_up = wsL.Range("B6").Value
    Dim KO_dn As Double: KO_dn = wsL.Range("B7").Value
    Dim vol1 As Double: vol1 = wsL.Range("B10").Value
    Dim vol2 As Double: vol2 = wsL.Range("B11").Value
    Dim corr As Double: corr = wsL.Range("B12").Value
    Dim r_US As Double: r_US = wsE.Range("B4").Value
    Dim r_KRW As Double: r_KRW = wsE.Range("B5").Value
    Dim barrels As Double: barrels = wsE.Range("B9").Value

    Dim T_WTI As Double: T_WTI = wsE.Range("B15").Value
    Dim T_FX As Double: T_FX = wsE.Range("B16").Value
    Dim T_total As Double: T_total = VMax(T_WTI, T_FX)

    Dim Steps_A As Long: Steps_A = CLng(T_total * STEPS_PER_YEAR)
    Dim Steps_B1 As Long: Steps_B1 = CLng(T_WTI * STEPS_PER_YEAR)
    If Steps_B1 < 1 Then Steps_B1 = 1

    Dim n_sens As Long: n_sens = 10000

    If Beta_Mat_WTI_Steps = 0 Or gFvW0 = 0# Then
        gFvW0 = Calc_WTI_KO_LSMC_Price(Lambda, JumpMean, JumpVol, r_US, KO_up, KO_dn, _
                                       S1_0, S2_0, vol1, Steps_B1, T_WTI, 50000, k, True) * barrels
    End If

    Dim r As Long: r = 25
    ws.Cells(r, 1).Value = "FV Validation Matrix (Query at v1 = 1.05 ITM Subdomain)"
    ws.Cells(r, 1).Font.Bold = True: r = r + 1
    ws.Cells(r, 1) = "tau/T": ws.Cells(r, 2) = "A_Beta_FV": ws.Cells(r, 3) = "A_Reprice_FV": ws.Cells(r, 4) = "B1_Beta_FV": ws.Cells(r, 5) = "B1_Reprice_FV": ws.Cells(r, 6) = "A_Pct_Diff": ws.Cells(r, 7) = "B1_Pct_Diff"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Font.Bold = True: r = r + 1

    Dim S1_test As Double: S1_test = 1.05 * S1_0

    ' [���� ���� 7 ����] fracs ������ 10%, 50%, 90%�� �������� t=0 ������ ���� ���� ���� ����
    Dim fracs As Variant: fracs = Array(0.1, 0.5, 0.9)
    Dim j As Long
    For j = LBound(fracs) To UBound(fracs)
        Dim tfrac As Double: tfrac = fracs(j)

        Dim stp_A As Long: stp_A = CLng(tfrac * Steps_A)
        If stp_A > Steps_A Then stp_A = Steps_A
        Dim tau_A As Double: tau_A = T_total - stp_A * (T_total / Steps_A)
        If tau_A <= 0.001 Then tau_A = 0.001

        Dim stp_B1 As Long: stp_B1 = CLng(tfrac * Steps_B1)
        If stp_B1 > Steps_B1 Then stp_B1 = Steps_B1
        Dim tau_B1 As Double: tau_B1 = T_WTI - stp_B1 * (T_WTI / Steps_B1)
        If tau_B1 <= 0.001 Then tau_B1 = 0.001

        Dim aBetaFV As Double: aBetaFV = QuantoFV_FromBeta(S1_test, S2_0, S1_0, S2_0, stp_A) * barrels
        Dim aRepriceFV As Double
        aRepriceFV = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, r_US, r_KRW - r_US, KO_up, KO_dn, S1_test, S2_0, vol1, vol2, corr, CLng(tau_A * STEPS_PER_YEAR), tau_A, n_sens, S1_0, False) * barrels

        Dim b1BetaFV As Double: b1BetaFV = WTI_KO_Call_BetaFV(S1_test, S1_0, stp_B1, barrels, 0)
        Dim b1RepriceFV As Double
        b1RepriceFV = Calc_WTI_KO_LSMC_Price(Lambda, JumpMean, JumpVol, r_US, KO_up, KO_dn, S1_test, S2_0, vol1, CLng(tau_B1 * STEPS_PER_YEAR), tau_B1, n_sens, S1_0, False) * barrels

        ws.Cells(r, 1) = tfrac
        ws.Cells(r, 2) = aBetaFV
        ws.Cells(r, 3) = aRepriceFV
        ws.Cells(r, 4) = b1BetaFV
        ws.Cells(r, 5) = b1RepriceFV
        ws.Cells(r, 6) = IIf(aRepriceFV <> 0#, Abs(aBetaFV - aRepriceFV) / Abs(aRepriceFV), 0#)
        ws.Cells(r, 7) = IIf(b1RepriceFV <> 0#, Abs(b1BetaFV - b1RepriceFV) / Abs(b1RepriceFV), 0#)
        ws.Cells(r, 6).NumberFormat = "0.0%": ws.Cells(r, 7).NumberFormat = "0.0%"
        r = r + 1
    Next j
End Sub

Private Sub Build_AcctPaths(ByVal n As Long, ByVal Steps As Long, ByVal S0_WTI As Double, ByVal S0_FX As Double, _
        ByVal vol1 As Double, ByVal vol2 As Double, ByVal corr As Double, ByVal Drift1 As Double, ByVal Drift2 As Double, _
        ByVal Lambda As Double, ByVal JumpMean As Double, ByVal JumpVol As Double, ByVal KO_up As Double, ByVal KO_dn As Double, ByVal T_total As Double)
    Dim dt As Double: dt = T_total / Steps
    Dim kappa As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt
    Dim sqdt  As Double: sqdt = Sqr(dt)

    ReDim AP_S1(1 To n, 0 To Steps): ReDim AP_S2(1 To n, 0 To Steps): ReDim AP_KO(1 To n)
    
    Rnd -1
    Randomize 20240101

    Dim sim As Long, stp As Long, S1 As Double, s2 As Double, S1_prev As Double
    Dim z1 As Double, z2 As Double, e2 As Double, nJ As Long, ko_step As Long

    For sim = 1 To n
        S1 = S0_WTI: s2 = S0_FX: S1_prev = S0_WTI: ko_step = 0
        AP_S1(sim, 0) = S0_WTI: AP_S2(sim, 0) = S0_FX
        For stp = 1 To Steps
            z1 = GetNormal(): z2 = GetNormal()
            e2 = corr * z1 + Sqr(1# - corr ^ 2) * z2
            nJ = GetPoisson(Lambda * dt)
            S1 = S1 * Exp((Drift1 - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt + vol1 * sqdt * z1 + JumpSum(nJ, JumpMean, JumpVol))
            s2 = s2 * Exp((Drift2 - 0.5 * vol2 ^ 2) * dt + vol2 * sqdt * e2)
            AP_S1(sim, stp) = S1: AP_S2(sim, stp) = s2
            If ko_step = 0 Then
                If S1 >= KO_up Or S1 <= KO_dn Then
                    ko_step = stp
                ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then
                    If GetUniform() < Exp(-2# * Log(KO_up / S1_prev) * Log(KO_up / S1) / lnVar) Then ko_step = stp
                    If ko_step = 0 And GetUniform() < Exp(-2# * Log(S1_prev / KO_dn) * Log(S1 / KO_dn) / lnVar) Then ko_step = stp
                End If
            End If
            S1_prev = S1
        Next stp
        AP_KO(sim) = ko_step
    Next sim
End Sub

Private Function StdFrom(ByVal sumv As Double, ByVal sumsq As Double, ByVal n As Long) As Double
    If n < 2 Then StdFrom = 0#: Exit Function
    Dim m As Double: m = sumv / n
    Dim v As Double: v = sumsq / n - m * m
    If v < 0# Then v = 0#
    StdFrom = Sqr(v)
End Function

Private Function MeanAbsArr(ByRef a() As Double, ByVal n As Long) As Double
    Dim s As Double, i As Long
    For i = 1 To n: s = s + Abs(a(i)): Next i
    MeanAbsArr = IIf(n > 0, s / n, 0#)
End Function

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet: On Error Resume Next: Set ws = Sheets(nm): On Error GoTo 0
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = nm
        If Err.Number <> 0 Then Set EnsureSheet = Nothing: Exit Function
        On Error GoTo 0
    End If
    Set EnsureSheet = ws
End Function

Private Sub SeedCFHInputs(ByVal ws As Worksheet)
    If Len(CStr(ws.Range("A1").Value)) = 0 Then
        ws.Range("A1").Value = "CFH Accounting Inputs": ws.Range("A1").Font.Bold = True
        ws.Range("A2").Value = "n_acct (realised paths)": ws.Range("B2").Value = 20000
        ws.Range("A3").Value = "n_sens (FV-validation reprice)": ws.Range("B3").Value = 1000
        ws.Columns("A:B").AutoFit
    End If
End Sub

Private Sub AllocAccumulators(ByVal Steps As Long, ByVal n As Long)
    ReDim sA_FV(0 To Steps): ReDim sA_Intr(0 To Steps): ReDim sA_TV(0 To Steps)
    ReDim sA_CFHR(0 To Steps): ReDim sA_COH(0 To Steps): ReDim sA_COHamort(0 To Steps)
    ReDim sA_Ineff(0 To Steps): ReDim sA_Alive(0 To Steps): ReDim qA_FV(0 To Steps): ReDim qA_CFHR(0 To Steps)
    ReDim sB_FVw(0 To Steps): ReDim sB_FVx(0 To Steps): ReDim sB_CFHRw(0 To Steps): ReDim sB_CFHRx(0 To Steps)
    ReDim sB_Ineffw(0 To Steps): ReDim sB_Ineffx(0 To Steps)
    ReDim sB_COHamortw(0 To Steps): ReDim sB_COHamortx(0 To Steps)
    ReDim sB_FVTPL(0 To Steps)
    ReDim qB_FVx(0 To Steps): ReDim qB_CFHRx(0 To Steps)
    ReDim detS_A(0 To Steps, 0 To 7): ReDim detS_Bw(0 To Steps, 0 To 2): ReDim detS_Bx(0 To Steps, 0 To 3)
    ReDim detK_A(0 To Steps, 0 To 7): ReDim detK_Bw(0 To Steps, 0 To 2): ReDim detK_Bx(0 To Steps, 0 To 3)
    ReDim tA_FVpeak(1 To n): ReDim tA_CFHRend(1 To n): ReDim tA_IneffTot(1 To n)
    ReDim tB_FVgross(1 To n): ReDim tB_CFHRend(1 To n): ReDim tB_IneffTot(1 To n): ReDim tB_postKO(1 To n): ReDim tKOflag(1 To n)
    ReDim sE_EconA(0 To Steps): ReDim sE_EconB(0 To Steps): ReDim qE_EconA(0 To Steps): ReDim qE_EconB(0 To Steps)
    ReDim tA_NakedMax(1 To n): ReDim tB_NakedMax(1 To n)
    ReDim tEconPathA(1 To n): ReDim tEconPathB(1 To n)
End Sub

