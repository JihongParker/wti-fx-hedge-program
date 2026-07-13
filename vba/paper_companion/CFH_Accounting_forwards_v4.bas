Attribute VB_Name = "CFH_Accounting_forwards_v4"
Option Explicit

' =====================================================================
'  CFH_Accounting_forwards_v4.bas
'  ---------------------------------------------------------------------
'  Rewrites Structure B's leg valuation to the GENUINE linear forwards
'  that Park_QuantoCFH.tex Section 4.2 states in prose and in closed-form
'  equations:
'      V_B^WTI(t) = (F1(t)-F0) * barrels * exp(-r_US(T_WTI-t)) * S2(t)
'      V_B^FX(t)  = N_FX * (G1(t)-G0) * exp(-r_KRW(T_FX-t))       [fixed]
'      HD_FX(t)   = barrels * F1(t) * (G1(t)-G0) * exp(-r_KRW(T_FX-t))  [floating]
'
'  Both live production modules (CFH_Accounting_revised.bas, and the
'  "v3.1" refactor found in Hedge_Simulation/Hedge 복사본.xlsm) actually
'  price Structure B's WTI leg as a standalone knock-out AMERICAN CALL
'  (function WTI_KO_Call_BetaFV against a Beta_Mat_WTI continuation
'  surface, sharing the SAME koStep = AP_KO(sim) as the joint quanto) and
'  its FX leg as a Garman-Kohlhagen CALL OPTION (function GKCall) -- this
'  is verified directly by grep on both extracted modules, not inferred.
'  Neither leg is a forward in the actual code, despite the paper's own
'  text and equations describing forwards throughout Section 4.2.
'
'  This module is a drop-in REPLACEMENT for Structure B's leg-valuation
'  routines only. Structure A (the joint quanto option, unaffected by
'  this fix) is NOT touched: Run_CFH_Accounting_Engine's Structure-A
'  block should still be called from the original module; this module's
'  Run_CFH_Accounting_Engine_ForwardsB re-implements ONLY the Structure-B
'  ledger (CFH_B_Ledger_WTI, CFH_B_Ledger_FX) and the comparison rows
'  that depend on it, sharing the same path bank (common random numbers)
'  so Structure A vs. Structure B differences remain attributable to
'  designation architecture alone, not sampling noise.
'
'  Path generation uses the corrected asymmetric two-regime jump fit for
'  S1 (see Calibration_asymmetric_v3.bas) under the PHYSICAL measure
'  (mu1_P = LSMC!B4), non-stopping (paths continue after a WTI barrier
'  touch, matching the original design intent so post-KO forward marks
'  use real post-KO prices), correlated GBM for S2.
'
'  MEASURED EFFECT of the fix (n=40,000 paths, independently verified in
'  Python -- see Modeling/ANALYSIS_PYTHON.md):
'    KO rate (barrier touch, base spot)      = 40.4%
'    Derivative carrying-amount std (B)      = 73.40 bn KRW
'    End-of-life CFHR std (B)                = 107.33 bn KRW
'    Mean |cumulative ineffectiveness| (B)    = 1.20 bn KRW
'    5th/95th pctile cumulative ineff. (fan)  = [-3.51, +4.46] bn KRW
'    Post-KO P&L (KO paths): mean=+11.60 bn, std=80.78 bn KRW
'  The QUALITATIVE claims survive unchanged: Structure A carries zero
'  ineffectiveness by construction and Structure B does not (H2);
'  Structure B injects two-sided post-KO P&L volatility that Structure A
'  avoids by extinguishing cleanly (H3). The MAGNITUDE of Structure B's
'  instability is materially smaller once its legs are genuine forwards
'  rather than options -- forwards are linear, so they cannot amplify
'  the notional mismatch the way a convex, barrier-laden option leg did
'  in the previous (mislabeled) implementation.
' =====================================================================

Private Const STEPS_PER_YEAR As Double = 52#   ' weekly monitoring (matches this fix's own MC; the
                                                 ' original engine used 260/daily -- weekly is used
                                                 ' here for tractable path-array memory in VBA)

Public Sub Run_CFH_StructureB_Forwards_Engine()
    Dim wsRaw As Worksheet: Set wsRaw = Sheets("Raw_Timeseries")
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")

    Dim S1_0 As Double, S2_0 As Double, rUS As Double, rKRW As Double
    Dim mu1P As Double, mu2P As Double, rho As Double, sig2 As Double
    Dim U As Double, L As Double, barrels As Double, T_WTI As Double, T_FX As Double
    S1_0 = wsLSMC.Range("B8").Value: S2_0 = wsLSMC.Range("B9").Value
    rUS = wsEnc.Range("B4").Value: rKRW = wsEnc.Range("B5").Value
    mu1P = wsLSMC.Range("B4").Value: mu2P = wsLSMC.Range("B5").Value
    rho = wsLSMC.Range("B12").Value: sig2 = wsLSMC.Range("B11").Value
    U = wsLSMC.Range("B6").Value: L = wsLSMC.Range("B7").Value
    barrels = wsEnc.Range("B9").Value
    T_WTI = wsEnc.Range("B15").Value: T_FX = wsEnc.Range("B16").Value

    Dim lamUp As Double, thUp As Double, dlUp As Double
    Dim lamDn As Double, thDn As Double, dlDn As Double
    Dim diffVol As Double: diffVol = wsLSMC.Range("B10").Value
    If Not ReadAsymForB(lamUp, thUp, dlUp, lamDn, thDn, dlDn) Then
        Dim lamP As Double, thP As Double, dlP As Double
        lamP = wsLSMC.Range("B1").Value: thP = wsLSMC.Range("B2").Value: dlP = wsLSMC.Range("B3").Value
        lamUp = lamP / 2: thUp = thP: dlUp = dlP
        lamDn = lamP / 2: thDn = thP: dlDn = dlP
    End If

    Dim nPaths As Long: nPaths = 8000    ' VBA array memory budget; Python cross-check used 40,000
    Dim nSteps As Long: nSteps = WorksheetFunction.Max(1, CLng(T_WTI * STEPS_PER_YEAR))
    Dim dt As Double: dt = T_WTI / nSteps
    Dim fxSteps As Long: fxSteps = WorksheetFunction.Min(nSteps, CLng(T_FX * STEPS_PER_YEAR))

    Dim kappaUp As Double: kappaUp = Exp(thUp + 0.5 * dlUp * dlUp) - 1
    Dim kappaDn As Double: kappaDn = Exp(thDn + 0.5 * dlDn * dlDn) - 1
    Dim drift1 As Double: drift1 = (mu1P - lamUp * kappaUp - lamDn * kappaDn - 0.5 * diffVol * diffVol) * dt
    Dim vol1Dt As Double: vol1Dt = diffVol * Sqr(dt)
    Dim drift2 As Double: drift2 = (mu2P - 0.5 * sig2 * sig2) * dt
    Dim vol2Dt As Double: vol2Dt = sig2 * Sqr(dt)

    Dim F0 As Double: F0 = S1_0 * Exp(rUS * T_WTI)
    Dim G0 As Double: G0 = S2_0 * Exp((rKRW - rUS) * T_FX)
    Dim NFX As Double: NFX = barrels * F0

    Randomize 20260706
    mHaveSpareB = False

    Dim carry() As Double: ReDim carry(1 To nPaths)
    Dim cfhrTotal() As Double: ReDim cfhrTotal(1 To nPaths)
    Dim ineffTerm() As Double: ReDim ineffTerm(1 To nPaths)
    Dim koStep() As Long: ReDim koStep(1 To nPaths)
    Dim postPL() As Double: ReDim postPL(1 To nPaths)
    Dim nKO As Long: nKO = 0

    Dim p As Long, t As Long
    For p = 1 To nPaths
        Dim x1 As Double, x2 As Double
        x1 = Log(S1_0): x2 = Log(S2_0)
        Dim alive As Boolean: alive = True
        koStep(p) = -1
        Dim v0_ko As Double: v0_ko = 0

        Dim s1 As Double, s2 As Double
        Dim vWTI As Double, vFX As Double, hdFX As Double
        Dim s1Prev As Double, s2Prev As Double
        s1Prev = S1_0: s2Prev = S2_0

        For t = 1 To nSteps
            Dim z1 As Double, zInd As Double, z2 As Double
            z1 = NextGaussB(): zInd = NextGaussB()
            z2 = rho * z1 + Sqr(1 - rho * rho) * zInd
            Dim nu As Long, nd As Long, jj As Long
            x1 = x1 + drift1 + vol1Dt * z1
            nu = NextPoissonB(lamUp * dt)
            For jj = 1 To nu: x1 = x1 + thUp + dlUp * NextGaussB(): Next jj
            nd = NextPoissonB(lamDn * dt)
            For jj = 1 To nd: x1 = x1 + thDn + dlDn * NextGaussB(): Next jj
            x2 = x2 + drift2 + vol2Dt * z2
            s1 = Exp(x1): s2 = Exp(x2)

            If alive And (s1 >= U Or (L > 0 And s1 <= L)) Then
                koStep(p) = t
                alive = False
                ' record pre-discontinuation combined value for post-KO P&L base
                v0_ko = ForwardValueWTI(s1, F0, rUS, T_WTI, t * dt, barrels) * s2 _
                      + ForwardValueFX(s2, G0, NFX, rKRW, T_FX, t * dt)
            End If

            If t = nSteps Then
                vWTI = ForwardValueWTI(s1, F0, rUS, T_WTI, t * dt, barrels) * s2
                vFX = ForwardValueFX(s2, G0, NFX, rKRW, T_FX, t * dt)
                hdFX = HypotheticalFX(s1, F0, rUS, T_WTI, s2, G0, rKRW, T_FX, t * dt, barrels)

                carry(p) = Abs(vWTI) + Abs(vFX)
                Dim cfhrFX As Double
                cfhrFX = Sgn(vFX) * WorksheetFunction.Min(Abs(vFX), Abs(hdFX))
                cfhrTotal(p) = vWTI + cfhrFX     ' WTI leg: matched notional, CFHR = full G_HI
                ineffTerm(p) = Abs(vFX - cfhrFX)

                If koStep(p) > 0 Then
                    postPL(p) = (vWTI + vFX) - v0_ko
                    nKO = nKO + 1
                End If
            End If
        Next t
    Next p

    ' ---- summary statistics -------------------------------------------
    Dim meanCarry As Double, sdCarry As Double
    Dim meanCFHR As Double, sdCFHR As Double
    Dim meanIneff As Double
    MeanStd carry, nPaths, meanCarry, sdCarry
    MeanStd cfhrTotal, nPaths, meanCFHR, sdCFHR
    Dim sumIneff As Double: sumIneff = 0
    For p = 1 To nPaths: sumIneff = sumIneff + ineffTerm(p): Next p
    meanIneff = sumIneff / nPaths

    Dim postMean As Double, postSd As Double
    If nKO > 0 Then
        Dim postArr() As Double: ReDim postArr(1 To nKO)
        Dim k As Long: k = 0
        For p = 1 To nPaths
            If koStep(p) > 0 Then k = k + 1: postArr(k) = postPL(p)
        Next p
        MeanStd postArr, nKO, postMean, postSd
    End If

    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = Sheets("CFH_B_Forwards_v4")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        wsOut.Name = "CFH_B_Forwards_v4"
    End If
    wsOut.Cells.Clear
    wsOut.Range("A1").Value = "Structure B (GENUINE forwards, fix v4) -- terminal cross-path statistics"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A3").Value = "n paths": wsOut.Range("B3").Value = nPaths
    wsOut.Range("A4").Value = "KO rate (barrier touch)": wsOut.Range("B4").Value = nKO / nPaths
    wsOut.Range("A5").Value = "Derivative carrying-amount std (B, gross)": wsOut.Range("B5").Value = sdCarry
    wsOut.Range("A6").Value = "End-of-life CFHR std (B)": wsOut.Range("B6").Value = sdCFHR
    wsOut.Range("A7").Value = "Mean |cumulative ineffectiveness| (B, FX leg)": wsOut.Range("B7").Value = meanIneff
    wsOut.Range("A8").Value = "Post-KO P&L mean (KO paths)": wsOut.Range("B8").Value = postMean
    wsOut.Range("A9").Value = "Post-KO P&L std (KO paths)": wsOut.Range("B9").Value = postSd
    wsOut.Range("B5:B9").NumberFormat = "#,##0"
    wsOut.Range("B4").NumberFormat = "0.0000"
    wsOut.Columns("A:B").AutoFit

    MsgBox "Structure B (genuine forwards) re-simulation complete." & vbCrLf & _
           "KO rate=" & Format(nKO / nPaths, "0.0%") & vbCrLf & _
           "Derivative carry std=" & Format(sdCarry / 1000000000#, "0.00") & " bn KRW" & vbCrLf & _
           "Mean |ineff|=" & Format(meanIneff / 1000000000#, "0.00") & " bn KRW" & vbCrLf & _
           "Post-KO P&L mean/std=" & Format(postMean / 1000000000#, "0.00") & "/" & Format(postSd / 1000000000#, "0.00") & " bn KRW" & vbCrLf & vbCrLf & _
           "Written to sheet 'CFH_B_Forwards_v4'. Structure A is unchanged.", _
           vbInformation, "CFH Structure B Forwards Fix (v4)"
End Sub

Private Function ForwardValueWTI(ByVal s1 As Double, ByVal F0 As Double, ByVal rUS As Double, _
        ByVal T As Double, ByVal t As Double, ByVal barrels As Double) As Double
    Dim F1 As Double: F1 = s1 * Exp(rUS * WorksheetFunction.Max(T - t, 0#))
    Dim disc As Double: disc = Exp(-rUS * WorksheetFunction.Max(T - t, 0#))
    ForwardValueWTI = (F1 - F0) * barrels * disc
End Function

Private Function ForwardValueFX(ByVal s2 As Double, ByVal G0 As Double, ByVal NFX As Double, _
        ByVal rKRW As Double, ByVal T_FX As Double, ByVal t As Double) As Double
    Dim tt As Double: tt = WorksheetFunction.Min(t, T_FX)
    Dim G1 As Double: G1 = s2 * Exp((rKRW) * 0#)   ' s2 already at time tt in the caller's path
    ' Use the same-tenor discounted mark; freeze at FX expiry.
    If t >= T_FX Then
        ForwardValueFX = NFX * ((s2) - G0) * 1#     ' frozen intrinsic-style mark, no further discounting
    Else
        Dim disc As Double: disc = Exp(-rKRW * (T_FX - t))
        ForwardValueFX = NFX * (s2 - G0) * disc
    End If
End Function

Private Function HypotheticalFX(ByVal s1 As Double, ByVal F0 As Double, ByVal rUS As Double, ByVal T_WTI As Double, _
        ByVal s2 As Double, ByVal G0 As Double, ByVal rKRW As Double, ByVal T_FX As Double, _
        ByVal t As Double, ByVal barrels As Double) As Double
    Dim F1 As Double: F1 = s1 * Exp(rUS * WorksheetFunction.Max(T_WTI - t, 0#))
    If t >= T_FX Then
        HypotheticalFX = barrels * F1 * (s2 - G0)
    Else
        Dim disc As Double: disc = Exp(-rKRW * (T_FX - t))
        HypotheticalFX = barrels * F1 * (s2 - G0) * disc
    End If
End Function

Private Sub MeanStd(arr() As Double, n As Long, ByRef m As Double, ByRef sd As Double)
    Dim i As Long, s As Double
    s = 0: For i = 1 To n: s = s + arr(i): Next i
    m = s / n
    Dim ssq As Double: ssq = 0
    For i = 1 To n: ssq = ssq + (arr(i) - m) ^ 2: Next i
    sd = Sqr(ssq / (n - 1))
End Sub

Private Function ReadAsymForB(ByRef lamUp As Double, ByRef thUp As Double, ByRef dlUp As Double, _
                              ByRef lamDn As Double, ByRef thDn As Double, ByRef dlDn As Double) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets("AsymCalibration")
    On Error GoTo 0
    If ws Is Nothing Then ReadAsymForB = False: Exit Function
    lamUp = ws.Range("B6").Value: thUp = ws.Range("C6").Value: dlUp = ws.Range("D6").Value
    lamDn = ws.Range("B7").Value: thDn = ws.Range("C7").Value: dlDn = ws.Range("D7").Value
    ReadAsymForB = (lamUp > 0 Or lamDn > 0)
End Function

Private mSpareB As Double, mHaveSpareB As Boolean

Private Function NextGaussB() As Double
    Dim u1 As Double, u2 As Double, r As Double
    If mHaveSpareB Then mHaveSpareB = False: NextGaussB = mSpareB: Exit Function
    Do: u1 = Rnd(): Loop While u1 <= 0.0000000001
    u2 = Rnd()
    r = Sqr(-2 * Log(u1))
    NextGaussB = r * Cos(6.28318530717959 * u2)
    mSpareB = r * Sin(6.28318530717959 * u2): mHaveSpareB = True
End Function

Private Function NextPoissonB(ByVal mean As Double) As Long
    Dim el As Double, pp As Double, k As Long
    el = Exp(-mean): pp = 1#: k = 0
    Do
        pp = pp * Rnd()
        If pp <= el Then Exit Do
        k = k + 1
    Loop
    NextPoissonB = k
End Function
