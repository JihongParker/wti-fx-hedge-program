Attribute VB_Name = "PaperVerification_revised"
Option Explicit

' =============================================================================
' MERGED 2026-06-22: this file is the "paper verification" bucket of the
' 3-file consolidation (Calibration_revised.bas / DeltaHedging_revised.bas /
' PaperVerification_revised.bas), per user request to reduce the project from
' 9 .bas files down to 3. Content (in order below) is the former:
'   ITM_count_consistency_revised.bas  (ITM/Alive + P/Q diagnostics)
'   PaperIntegritySuite_revised.bas    (4-item sensitivity suite + orchestrator)
' No code changes from the merge itself -- only the module boundary and the
' duplicate Attribute VB_Name / Option Explicit lines were removed.
' RestoreProductionBetaMat (originally ITM_count_consistency_revised.bas) is
' Public and is called directly by PaperIntegritySuite's items below within
' this same module now. See MODEL_SPEC ��25.
' =============================================================================

' #############################################################################
' ===== Originally: ITM_count_consistency_revised.bas =====
' #############################################################################

' =============================================================================
' ITM/Alive diagnostic engine ��� revised
'
' FIXES APPLIED:
'   FIX #2  BB formula corrected to log-space (was: level-space approximation)
'   FIX #6  Chart x-axis label corrected:
'             - column stores t/T (time elapsed), not tau/T (time remaining)
'             - direction: 0 = inception (now), 1 = maturity
'             - old label "0=maturity, 1=present" was backwards on both counts
'
' NEW ENGINES (2026-06-20, rewritten 2026-06-20b):
'   - Run_PQ_Drift_Bias_Diagnostic (16-1)
'   - Run_Sharpe_Sweep_Diagnostic (16-2)
'   - Run_Girsanov_Reweight_Diagnostic (16-3)
' Run_ITM_Diagnostic now triggers all for unified visualization.
'
' 2026-06-21 UPDATE: ITM/Alive core simulation grid changed from weekly (52)
' to daily (260) for full consistency with LSMC, American hedge, and P/Q engines.
' See MODEL_SPEC ��10 and ��14 for updated status.
'
' REVISION NOTE (2026-06-20b): the first pass of these three engines was a
' non-functional stub (caught in review): 16-1 ignored the FX leg and used an
' arbitrary binary delta proxy with an unpopulated chart; 16-2 computed
' "Simulated" and "Theoretical" bias with the literal same formula, so the
' residual was always exactly zero by construction (no Monte Carlo ever ran);
' 16-3 multiplied the Girsanov likelihood ratio by random noise instead of
' reweighting an actual payoff distribution. All three are rewritten below to
' run real path simulations (both WTI+FX legs, closed-form Black-76/GK deltas,
' populated XY/column charts).
' =============================================================================
Public Sub Run_ITM_Diagnostic()

    Dim wsEnc  As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsDiag As Worksheet

    ' FIX (2026-06-21f, perf): this Sub's own core loop is up to ~6 maturities
    ' x 10,000 paths x up to 780 steps (3yr maturity) -- tens of millions of
    ' iterations -- plus it chains 3 more diagnostic Subs below, and NONE of
    ' the 4 had ScreenUpdating/Calculation/EnableEvents toggles before this
    ' fix (unlike American_Delta_revised.bas). See MODEL_SPEC ��19.6.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail

    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Diagnostics").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail
    Set wsDiag = Sheets.Add(After:=Sheets(Sheets.Count))
    wsDiag.Name = "Diagnostics"

    ' FIX #6: column header corrected from "tau_over_T" to "t_over_T"
    ' t/T = i/Steps: 0 ��� inception, 1 = maturity
    wsDiag.Range("A1:D1").Value = Array("Maturity_T", "Step_i", "t_over_T", "P_AliveITM")

    ' Parameters (mirror LSMC sheet ��� Q-measure)
    Dim Lambda   As Double: Lambda = wsLSMC.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLSMC.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsLSMC.Range("B3").Value
    Dim KOUpper  As Double: KOUpper = wsLSMC.Range("B6").Value
    Dim KOLower  As Double: KOLower = wsLSMC.Range("B7").Value
    Dim S1_0     As Double: S1_0 = wsLSMC.Range("B8").Value
    Dim vol1     As Double: vol1 = wsLSMC.Range("B10").Value
    ' REVISED 2026-06-21c: a 2026-06-21b pass changed this from ATM to
    ' S1_0*0.95, reasoning the real contract strike should apply here too ���
    ' but Optimization_English.docx ��II confirms the American/LSMC strike IS
    ' ATM (K=spot); the 0.95-ITM convention is European-only. Reverted to the
    ' original ATM design, which was correct all along. See MODEL_SPEC ��17.1.
    Dim K   As Double: K = S1_0   ' ATM (K = current spot for American per docx)

    Dim r_US  As Double: r_US = wsEnc.Range("B4").Value
    Dim Drift1 As Double: Drift1 = r_US   ' Q-measure override
    Dim kappa  As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#

    Dim maturities() As Variant
    maturities = Array(0.5, 1, 1.5, 2, 2.5, 3)

    Dim n_paths As Long: n_paths = 10000

    Dim outRow As Long: outRow = 2
    Dim mIdx   As Long

    For mIdx = LBound(maturities) To UBound(maturities)

        Dim T     As Double: T = maturities(mIdx)
        Dim Steps As Long:   Steps = CLng(T * STEPS_PER_YEAR)   ' Daily, matches all other engines (LSMC, American, PQ diagnostics)
        Dim dt    As Double: dt = T / Steps
        Dim sqdt  As Double: sqdt = Sqr(dt)

        ' FIX #2: log-space bridge variance
        Dim lnVar As Double: lnVar = vol1 ^ 2 * dt

        Dim itmCount() As Long
        ReDim itmCount(1 To Steps)

        Randomize 12345

        Dim p As Long, i As Long
        For p = 1 To n_paths

            Dim s1      As Double: s1 = S1_0
            Dim S1_prev As Double: S1_prev = S1_0
            Dim isAlive As Boolean: isAlive = True

            For i = 1 To Steps
                If isAlive Then
                    S1_prev = s1

                    Dim z1   As Double: z1 = GetNormal()
                    Dim nJ   As Long:   nJ = GetPoisson(Lambda * dt)
                    Dim jSum As Double: jSum = JumpSum(nJ, JumpMean, JumpVol)

                    s1 = s1 * Exp((Drift1 - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt _
                        + vol1 * sqdt * z1 + jSum)

                    ' (a) endpoint check
                    If s1 >= KOUpper Or s1 <= KOLower Then
                        isAlive = False
                    ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then
                        ' (b) FIX #2: log-space Brownian Bridge correction
                        ' Correct formula: p = exp(-2 * ln(B/S_prev) * ln(B/S) / (vol^2 * dt))
                        ' Replaces original:  (B-S_prev)*(B-S) / (vol^2*S_prev^2*dt)  [level-space error]
                        Dim p_touch_up As Double, p_touch_dn As Double
                        p_touch_up = Exp(-2# * Log(KOUpper / S1_prev) * Log(KOUpper / s1) / lnVar)
                        p_touch_dn = Exp(-2# * Log(S1_prev / KOLower) * Log(s1 / KOLower) / lnVar)
                        If Rnd() < p_touch_up Then isAlive = False
                        If isAlive And Rnd() < p_touch_dn Then isAlive = False
                    End If
                    ' nJ > 0: endpoint check only (gap risk)
                End If

                If isAlive Then
                    If (s1 - K) > 0 Then
                        itmCount(i) = itmCount(i) + 1
                    End If
                End If
            Next i
        Next p

        ' Write results
        ' FIX #6: column C stores t/T = i/Steps (time elapsed fraction)
        '         t/T ��� 0 at inception, 1 at maturity
        For i = 1 To Steps
            wsDiag.Cells(outRow, 1).Value = T
            wsDiag.Cells(outRow, 2).Value = i
            wsDiag.Cells(outRow, 3).Value = i / Steps          ' t/T: 0=now, 1=maturity
            wsDiag.Cells(outRow, 4).Value = itmCount(i) / n_paths
            outRow = outRow + 1
        Next i

    Next mIdx

    ' Threshold reference lines
    wsDiag.Range("F1").Value = "Threshold_10"
    wsDiag.Range("F2").Value = 10 / n_paths
    wsDiag.Range("G1").Value = "Threshold_30"
    wsDiag.Range("G2").Value = 30 / n_paths
    wsDiag.Range("H1").Value = "Threshold_50"
    wsDiag.Range("H2").Value = 50 / n_paths

    Call BuildITMChart(wsDiag, maturities, outRow - 1)

    ' Run all consistency engines when Run_ITM_Diagnostic is executed as macro
    ' (per user request for unified verification)
    Call Run_PQ_Drift_Bias_Diagnostic
    Call Run_Sharpe_Sweep_Diagnostic
    Call Run_Girsanov_Reweight_Diagnostic

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "ITM/Alive diagnostic complete (Q-measure, log-space BB correction)." & vbCrLf & _
           "All P/Q consistency engines executed: DriftBias, SharpeSweep, Girsanov." & vbCrLf & _
           "r_US = " & Format(r_US, "0.0000") & vbCrLf & _
           CStr(UBound(maturities) - LBound(maturities) + 1) & " maturities, " & _
           (outRow - 2) & " total rows" & vbCrLf & _
           "Check 'Diagnostics', 'DriftBias', 'SharpeSweep', 'Girsanov' sheets.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_ITM_Diagnostic failed: " & Err.Description, vbCritical
End Sub

' =============================================================================
' BuildITMChart
' FIX #6: axis label corrected to "t/T  (0=inception, 1=maturity)"
'         Data column stores i/Steps = t/T (time elapsed), NOT tau/T (time remaining)
' =============================================================================
Private Sub BuildITMChart(wsDiag As Worksheet, maturities As Variant, lastRow As Long)

    Dim chtObj As ChartObject
    Set chtObj = wsDiag.ChartObjects.Add(Left:=300, Top:=10, Width:=600, Height:=380)
    chtObj.Chart.ChartType = xlXYScatterLinesNoMarkers

    Dim mIdx As Long

    For mIdx = LBound(maturities) To UBound(maturities)

        Dim T As Double: T = maturities(mIdx)

        Dim firstRow As Long, lastRowForT As Long, r As Long
        firstRow = 0: lastRowForT = 0
        For r = 2 To lastRow
            If wsDiag.Cells(r, 1).Value = T Then
                If firstRow = 0 Then firstRow = r
                lastRowForT = r
            End If
        Next r

        If firstRow > 0 Then
            Dim s As Series
            Set s = chtObj.Chart.SeriesCollection.NewSeries
            s.Name = "T=" & T
            s.XValues = wsDiag.Range(wsDiag.Cells(firstRow, 3), wsDiag.Cells(lastRowForT, 3))
            s.Values = wsDiag.Range(wsDiag.Cells(firstRow, 4), wsDiag.Cells(lastRowForT, 4))
        End If
    Next mIdx

    ' Threshold lines
    Dim thLabels As Variant: thLabels = Array("Threshold_10", "Threshold_30", "Threshold_50")
    Dim thCols   As Variant: thCols = Array(6, 7, 8)
    Dim K As Long
    For K = LBound(thLabels) To UBound(thLabels)
        Dim sTh As Series
        Set sTh = chtObj.Chart.SeriesCollection.NewSeries
        sTh.Name = thLabels(K)
        sTh.XValues = Array(0, 1)
        Dim thVal As Double: thVal = wsDiag.Cells(2, thCols(K)).Value
        sTh.Values = Array(thVal, thVal)
        sTh.Format.Line.DashStyle = msoLineDash
    Next K

    With chtObj.Chart
        .Axes(xlValue).ScaleType = xlScaleLogarithmic
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "P(Alive & ITM)  [log scale]"

        .Axes(xlCategory).HasTitle = True
        ' FIX #6: corrected label ��� data is t/T (time elapsed), 0=inception, 1=maturity
        ' Original was "tau/T (0=maturity, 1=present)" which was wrong on both counts:
        '   (a) the column stores t/T not tau/T
        '   (b) the direction was stated backwards
        .Axes(xlCategory).AxisTitle.Text = "t/T  (0=inception, 1=maturity)"
        .Axes(xlCategory).MinimumScale = 0
        .Axes(xlCategory).MaximumScale = 1

        .HasTitle = True
        .ChartTitle.Text = "ITM & Alive Probability by Maturity (Q-measure, log-space BB)"
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
    End With

End Sub

' =============================================================================
' NEW CONSISTENCY ENGINES (attached to ITM module per request)
' When Run_ITM_Diagnostic is run as macro, these are also executed.
' Visualizations: new sheets with tables + charts (similar to ITM).
' Based on MODEL_SPEC ��16 P/Q proposals.
' =============================================================================

' =============================================================================
' Closed-form deltas (Black-76 for WTI futures option, Garman-Kohlhagen for FX)
' LEGACY (kept for reference).
' Main P/Q diagnostics now use actual LSMC Beta_Mat deltas (see SimulateDriftBias)
' for consistency with American_Delta_revised.bas and paper requirements.
' =============================================================================
Private Function BS_Black76_Delta(ByVal f As Double, ByVal K As Double, ByVal vol As Double, _
                                    ByVal tau As Double, ByVal r As Double) As Double
    If tau <= 1E-06 Or vol <= 0# Then
        BS_Black76_Delta = IIf(f > K, 1#, 0#)
        Exit Function
    End If
    Dim d1 As Double: d1 = (Log(f / K) + 0.5 * vol ^ 2 * tau) / (vol * Sqr(tau))
    BS_Black76_Delta = Exp(-r * tau) * WorksheetFunction.Norm_S_Dist(d1, True)
End Function

Private Function GK_FX_Delta(ByVal s As Double, ByVal K As Double, ByVal vol As Double, _
                              ByVal tau As Double, ByVal r_for As Double, ByVal r_dom As Double) As Double
    If tau <= 1E-06 Or vol <= 0# Then
        GK_FX_Delta = IIf(s > K, 1#, 0#)
        Exit Function
    End If
    Dim d1 As Double: d1 = (Log(s / K) + (r_dom - r_for + 0.5 * vol ^ 2) * tau) / (vol * Sqr(tau))
    GK_FX_Delta = Exp(-r_for * tau) * WorksheetFunction.Norm_S_Dist(d1, True)
End Function

' =============================================================================
' SimulateDriftBias ��� shared core for 16-1 and 16-2.
' Runs a REAL Monte Carlo under the given P-drifts (Q-measure fixed at r_US /
' r_KRW-r_US per ��3), with the same BB-KO logic as the ITM engine above.
' Both legs (WTI theta1, FX theta2) are included:
'   Simulated_Bias  = realized discretized integral  -�� ��_t��S_t��(��^P-��^Q)��dt
'                     using path-by-path deltas from LSMC Beta_Mat (same as
'                     the actual American hedging engine in American_Delta_revised.bas)
'   Theoretical_Bias = the ��3 sanity-check approximation
'                     -(avg delta exposure)��(��^P-��^Q)��T, using the SAME
'                     simulated paths' average ����S exposure.
' A nonzero residual is expected and meaningful: it measures how much the
' single-average approximation misses relative to the full path integral
' (e.g. when �� and S are correlated through KO/vol clustering).
' Delta source changed to Beta_Mat (2026-06-20) for paper-level consistency
' with the main hedging engine (previously used closed-form for speed).
'
' REVISION (2026-06-20c):
'   - Steps now CLng(T*260), matching the daily grid used by the actual
'     American hedge engine (ITM/Alive core also switched to daily on 2026-06-21).
'   - primeBeta lets callers that loop at a fixed T (Sharpe sweep) prime
'     Beta_Mat once and reuse it, instead of re-pricing on every iteration.
'   - Theoretical_Bias exposure is now averaged over the FULL calendar grid
'     (post-KO steps contribute zero exposure) instead of only alive_steps,
'     so it is no longer overweighted for KO-shortened paths. Avg_Delta_WTI/FX
'     (the reported diagnostic columns) still average over alive_steps only,
'     since those report "delta while the hedge was actually live."
' =============================================================================
Private Sub SimulateDriftBias(ByVal T As Double, ByVal Drift1_sim As Double, ByVal Drift2_sim As Double, _
                               ByVal n_paths As Long, _
                               ByRef sim_bias As Double, ByRef theo_bias As Double, _
                               ByRef avg_delta_WTI As Double, ByRef avg_delta_FX As Double, _
                               Optional ByVal primeBeta As Boolean = True)

    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsEnc  As Worksheet: Set wsEnc = Sheets("Encoding")

    Dim Lambda   As Double: Lambda = wsLSMC.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLSMC.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsLSMC.Range("B3").Value
    Dim KOUpper  As Double: KOUpper = wsLSMC.Range("B6").Value
    Dim KOLower  As Double: KOLower = wsLSMC.Range("B7").Value
    Dim S1_0     As Double: S1_0 = wsLSMC.Range("B8").Value
    Dim S2_0     As Double: S2_0 = wsLSMC.Range("B9").Value
    Dim vol1     As Double: vol1 = wsLSMC.Range("B10").Value
    Dim vol2     As Double: vol2 = wsLSMC.Range("B11").Value
    Dim corr     As Double: corr = wsLSMC.Range("B12").Value

    Dim r_US  As Double: r_US = wsEnc.Range("B4").Value
    Dim r_KRW As Double: r_KRW = wsEnc.Range("B5").Value
    ' REVISED 2026-06-21c: was K_WTI=Encoding!B18 (113, a stress-shock input,
    ' not the strike) before this review; a 2026-06-21b pass then over-corrected
    ' to S1_0*0.95. Optimization_English.docx ��II confirms the American/LSMC
    ' engine is priced ATM (K=spot) ��� only the European leg uses 0.95*Spot.
    ' See MODEL_SPEC ��17.1.
    Dim K_WTI As Double: K_WTI = S1_0   ' ATM, matches Optimization_English.docx
    ' FIX (2026-06-21d): removed a dead `K_FX = Encoding!B19` read that was
    ' never referenced anywhere in this Sub. B19 is the FX stress-shock level
    ' (Run_LSMC_Engine's Stress_Premium grid input), not a strike ��� there is
    ' no FX strike concept anywhere else in this engine (the FX leg is
    ' normalized by S2_0, not a strike). Keeping an unused variable sourced
    ' from B19 under a "K_*" name was exactly the kind of stress-cell/strike
    ' mix-up already fixed elsewhere in ��17.1; harmless only because nothing
    ' read it, but removed so it can't be wired in by accident later. See
    ' MODEL_SPEC ��19.

    Dim kappa    As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim Q_wedge1 As Double: Q_wedge1 = Drift1_sim - r_US
    Dim Q_wedge2 As Double: Q_wedge2 = Drift2_sim - (r_KRW - r_US)

    Dim Steps As Long: Steps = CLng(T * STEPS_PER_YEAR)   ' daily, matches American_Delta_revised
    Dim dt    As Double: dt = T / Steps
    Dim sqdt  As Double: sqdt = Sqr(dt)
    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt

    ' Populate Beta_Mat under Q-measure for this T (using actual LSMC deltas
    ' for consistency with American_Delta_revised.bas hedging engine).
    ' This replaces previous closed-form Black-76/GK for paper accuracy.
    ' Skipped when primeBeta:=False (caller already primed Beta_Mat for this T ���
    ' avoids redundant repricing, e.g. across Sharpe-sweep factors at fixed T).
    If primeBeta Then
        Dim qDrift1 As Double: qDrift1 = r_US
        Dim qDrift2 As Double: qDrift2 = r_KRW - r_US
        Dim n_for_beta As Long: n_for_beta = 10000
        Dim dummy As Double
        dummy = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, qDrift1, qDrift2, _
                                KOUpper, KOLower, S1_0, S2_0, vol1, vol2, corr, _
                                Steps, T, n_for_beta, K_WTI, True)
    End If

    Randomize 12345

    Dim total_bias    As Double
    Dim total_exp_WTI As Double, total_exp_FX As Double
    Dim total_d_WTI   As Double, total_d_FX   As Double

    Dim p As Long, i As Long
    For p = 1 To n_paths
        Dim s1      As Double: s1 = S1_0
        Dim S2      As Double: S2 = S2_0
        Dim S1_prev As Double: S1_prev = S1_0
        Dim isAlive As Boolean: isAlive = True

        Dim path_bias   As Double: path_bias = 0#
        Dim sum_exp_WTI As Double: sum_exp_WTI = 0#
        Dim sum_exp_FX  As Double: sum_exp_FX = 0#
        Dim sum_d_WTI   As Double: sum_d_WTI = 0#
        Dim sum_d_FX    As Double: sum_d_FX = 0#
        Dim alive_steps As Long:   alive_steps = 0

        For i = 1 To Steps
            If isAlive Then
                S1_prev = s1
                Dim z1   As Double: z1 = GetNormal()
                Dim z2   As Double: z2 = GetNormal()
                Dim e2   As Double: e2 = corr * z1 + Sqr(1 - corr ^ 2) * z2
                Dim nJ   As Long:   nJ = GetPoisson(Lambda * dt)
                Dim jSum As Double: jSum = JumpSum(nJ, JumpMean, JumpVol)

                s1 = s1 * Exp((Drift1_sim - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt + vol1 * sqdt * z1 + jSum)
                S2 = S2 * Exp((Drift2_sim - 0.5 * vol2 ^ 2) * dt + vol2 * sqdt * e2)

                If s1 >= KOUpper Or s1 <= KOLower Then
                    isAlive = False
                ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then
                    Dim p_up As Double, p_dn As Double
                    p_up = Exp(-2# * Log(KOUpper / S1_prev) * Log(KOUpper / s1) / lnVar)
                    p_dn = Exp(-2# * Log(S1_prev / KOLower) * Log(s1 / KOLower) / lnVar)
                    If Rnd() < p_up Then isAlive = False
                    If isAlive And Rnd() < p_dn Then isAlive = False
                End If

                If isAlive Then
                    Dim v1_norm As Double: v1_norm = s1 / K_WTI
                    Dim v2_norm As Double: v2_norm = S2 / S2_0
                    Dim b0 As Double: b0 = Beta_Mat(i, 0)
                    Dim b1 As Double: b1 = Beta_Mat(i, 1)
                    Dim b2 As Double: b2 = Beta_Mat(i, 2)
                    Dim b3 As Double: b3 = Beta_Mat(i, 3)
                    Dim b4 As Double: b4 = Beta_Mat(i, 4)
                    Dim b5 As Double: b5 = Beta_Mat(i, 5)

                    Dim d_wti As Double
                    If K_WTI > 0 Then
                        Dim raw_dWTI As Double
                        raw_dWTI = ((b1 + 2# * b3 * v1_norm + b5 * v2_norm) / K_WTI) / S2
                        ' FIX (2026-06-21f, perf): WorksheetFunction.Max/Min ->
                        ' inline clamp. Runs per step per path (up to n_paths x
                        ' Steps times across the 6-maturity DriftBias sweep and
                        ' the Sharpe sweep). See MODEL_SPEC ��19.6.
                        If raw_dWTI < 0# Then
                            d_wti = 0#
                        ElseIf raw_dWTI > 1# Then
                            d_wti = 1#
                        Else
                            d_wti = raw_dWTI
                        End If
                    Else
                        d_wti = 0#
                    End If
                    Dim d_fx As Double: d_fx = d_wti

                    path_bias = path_bias - d_wti * s1 * Q_wedge1 * dt - d_fx * S2 * Q_wedge2 * dt
                    sum_exp_WTI = sum_exp_WTI + d_wti * s1
                    sum_exp_FX = sum_exp_FX + d_fx * S2
                    sum_d_WTI = sum_d_WTI + d_wti
                    sum_d_FX = sum_d_FX + d_fx
                    alive_steps = alive_steps + 1
                End If
            End If
        Next i

        total_bias = total_bias + path_bias

        ' FIX (2026-06-20c): exposure for Theoretical_Bias is averaged over the
        ' FULL calendar grid (Steps), not just alive_steps ��� a KO'd path
        ' contributes zero exposure for its dead time instead of having its
        ' alive-time average exposure extrapolated across the whole T. This
        ' avoids overweighting KO-shortened paths in the -(avg exp)*(wedge)*T
        ' sanity-check formula. Avg_Delta_WTI/FX (reported separately) still
        ' use alive_steps, since those report "delta while actually hedging."
        total_exp_WTI = total_exp_WTI + sum_exp_WTI / Steps
        total_exp_FX = total_exp_FX + sum_exp_FX / Steps
        If alive_steps > 0 Then
            total_d_WTI = total_d_WTI + sum_d_WTI / alive_steps
            total_d_FX = total_d_FX + sum_d_FX / alive_steps
        End If
    Next p

    sim_bias = total_bias / n_paths
    avg_delta_WTI = total_d_WTI / n_paths
    avg_delta_FX = total_d_FX / n_paths

    Dim avg_exp_WTI As Double: avg_exp_WTI = total_exp_WTI / n_paths
    Dim avg_exp_FX  As Double: avg_exp_FX = total_exp_FX / n_paths

    theo_bias = -avg_exp_WTI * Q_wedge1 * T - avg_exp_FX * Q_wedge2 * T
End Sub

' =============================================================================
' RestoreProductionBetaMat ��� FIX (2026-06-20c, Beta_Mat pollution).
' Beta_Mat is a workbook-wide Public array. SimulateDriftBias's priming calls
' overwrite it with diagnostic-grid coefficients (a different T/Steps than the
' production hedge run). If a user runs the P/Q diagnostics and then runs
' Run_American_DeltaHedge WITHOUT re-running Run_LSMC_Engine first, the hedge
' loop would read stale/mismatched coefficients ��� either an immediate
' "Subscript out of range" (if Beta_Mat is now smaller than hedge_steps needs)
' or, worse, silently wrong deltas (if it happens to be large enough).
' This re-primes Beta_Mat to exactly the state Run_LSMC_Engine's Base_Premium
' call would leave it in (daily grid, Strike=S1_0, ATM, n=50000), so
' production hedging is safe immediately after any diagnostic run.
' FIX (2026-06-21, REVISED 2026-06-21b, REVISED AGAIN 2026-06-21c): the first
' pass primed with Strike=Encoding!B18 (113) to match American_Delta_revised
' ��� but B18 is a stress-shock input, not the strike. A second pass then
' over-corrected to Strike=0.95*Spot (the EUROPEAN Black76/GK convention).
' Optimization_English.docx ��II confirms the American/LSMC engine is priced
' at-the-money (K=spot) ��� this was the original value before any of these
' fixes. Corrected here AND in American_Delta_revised.bas / SimulateDriftBias.
' See MODEL_SPEC ��17.1.
' =============================================================================
' FIX (2026-06-21f): changed Private -> Public so PaperIntegritySuite_revised.bas
' (now merged into this same file, see header) can reuse this instead of
' duplicating the re-priming logic a third time. Behavior unchanged.
' See MODEL_SPEC ��19/��20.
Public Sub RestoreProductionBetaMat()
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsEnc  As Worksheet: Set wsEnc = Sheets("Encoding")

    Dim Lambda   As Double: Lambda = wsLSMC.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLSMC.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsLSMC.Range("B3").Value
    Dim Drift1   As Double: Drift1 = wsLSMC.Range("B4").Value
    Dim Drift2   As Double: Drift2 = wsLSMC.Range("B5").Value
    Dim KOUpper  As Double: KOUpper = wsLSMC.Range("B6").Value
    Dim KOLower  As Double: KOLower = wsLSMC.Range("B7").Value
    Dim S1_0     As Double: S1_0 = wsLSMC.Range("B8").Value
    Dim S2_0     As Double: S2_0 = wsLSMC.Range("B9").Value
    Dim vol1     As Double: vol1 = wsLSMC.Range("B10").Value
    Dim vol2     As Double: vol2 = wsLSMC.Range("B11").Value
    Dim corr     As Double: corr = wsLSMC.Range("B12").Value
    Dim K_WTI    As Double: K_WTI = S1_0   ' ATM, matches Optimization_English.docx (American/LSMC is ATM, not the European 0.95*Spot)

    Dim T_WTI   As Double: T_WTI = wsEnc.Range("B15").Value
    Dim T_FX    As Double: T_FX = wsEnc.Range("B16").Value
    Dim T_total As Double: T_total = WorksheetFunction.Max(T_WTI, T_FX)
    If T_total <= 0 Then T_total = wsLSMC.Range("B14").Value
    If T_total <= 0 Then Exit Sub   ' no valid production T to restore against

    Dim Steps As Long: Steps = CLng(T_total * STEPS_PER_YEAR)
    Dim dummy As Double
    dummy = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, _
                            KOUpper, KOLower, S1_0, S2_0, vol1, vol2, corr, _
                            Steps, T_total, 50000, K_WTI, True)
End Sub

Public Sub Run_PQ_Drift_Bias_Diagnostic()
    ' 16-1: For each maturity T, run a real MC under the model's actual P-drift
    ' (Q fixed) and compare Simulated_Bias vs Theoretical_Bias. Both legs included.
    ' FIX (2026-06-21f, perf): toggles added ��� independently callable, and
    ' SimulateDriftBias's own loop is n_paths x Steps per maturity x 6
    ' maturities. See MODEL_SPEC ��19.6.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail

    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsBias As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("DriftBias").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail
    Set wsBias = Sheets.Add(After:=Sheets(Sheets.Count))
    wsBias.Name = "DriftBias"

    wsBias.Range("A1:F1").Value = Array("Maturity_T", "Simulated_Bias", "Theoretical_Bias", _
                                         "Residual", "Avg_Delta_WTI", "Avg_Delta_FX")

    Dim Drift1P As Double: Drift1P = wsLSMC.Range("B4").Value
    Dim Drift2P As Double: Drift2P = wsLSMC.Range("B5").Value

    Dim maturities() As Variant: maturities = Array(0.5, 1, 1.5, 2, 2.5, 3)
    Dim n_paths As Long: n_paths = 5000

    Dim outRow As Long: outRow = 2
    Dim mIdx As Long
    For mIdx = LBound(maturities) To UBound(maturities)
        Dim T As Double: T = maturities(mIdx)
        Dim sim_b As Double, theo_b As Double, ad_w As Double, ad_f As Double
        Call SimulateDriftBias(T, Drift1P, Drift2P, n_paths, sim_b, theo_b, ad_w, ad_f)

        wsBias.Cells(outRow, 1).Value = T
        wsBias.Cells(outRow, 2).Value = sim_b
        wsBias.Cells(outRow, 3).Value = theo_b
        wsBias.Cells(outRow, 4).Value = sim_b - theo_b
        wsBias.Cells(outRow, 5).Value = ad_w
        wsBias.Cells(outRow, 6).Value = ad_f
        outRow = outRow + 1
    Next mIdx

    Dim lastRow As Long: lastRow = outRow - 1

    Dim cht As ChartObject
    Set cht = wsBias.ChartObjects.Add(Left:=420, Top:=10, Width:=550, Height:=350)
    cht.Chart.ChartType = xlXYScatterLinesNoMarkers
    Dim sSim As Series, sTheo As Series
    Set sSim = cht.Chart.SeriesCollection.NewSeries
    sSim.Name = "Simulated_Bias"
    sSim.XValues = wsBias.Range("A2:A" & lastRow)
    sSim.Values = wsBias.Range("B2:B" & lastRow)
    Set sTheo = cht.Chart.SeriesCollection.NewSeries
    sTheo.Name = "Theoretical_Bias"
    sTheo.XValues = wsBias.Range("A2:A" & lastRow)
    sTheo.Values = wsBias.Range("C2:C" & lastRow)
    With cht.Chart
        .HasTitle = True
        .ChartTitle.Text = "P/Q Drift Bias: Simulated vs Theoretical by Maturity"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Maturity T (years)"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Mean Opt P&L Bias"
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
    End With

    wsBias.Columns("A:F").AutoFit

    ' FIX (2026-06-20c): undo the Beta_Mat pollution from the priming calls above
    ' before handing control back, so a subsequent Run_American_DeltaHedge is safe.
    Call RestoreProductionBetaMat

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "PQ_Drift_Bias_Diagnostic complete (real MC, WTI+FX legs)." & vbCrLf & _
           (UBound(maturities) - LBound(maturities) + 1) & " maturities x " & n_paths & " paths." & vbCrLf & _
           "Beta_Mat restored to production state." & vbCrLf & _
           "See 'DriftBias' sheet.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_PQ_Drift_Bias_Diagnostic failed: " & Err.Description, vbCritical
End Sub

Public Sub Run_Sharpe_Sweep_Diagnostic()
    ' 16-2: Scale the P-drift wedge (theta) by a factor, Q-measure held fixed,
    ' and check whether Simulated_Bias actually scales linearly as theory predicts.
    ' FIX (2026-06-21f, perf): toggles added ��� see MODEL_SPEC ��19.6.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail

    Dim wsLSMC  As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsEnc   As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsSweep As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("SharpeSweep").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail
    Set wsSweep = Sheets.Add(After:=Sheets(Sheets.Count))
    wsSweep.Name = "SharpeSweep"

    wsSweep.Range("A1:D1").Value = Array("Mu_P_Factor", "Simulated_Bias", "Theoretical_Bias", "Residual")

    Dim base_d1 As Double: base_d1 = wsLSMC.Range("B4").Value
    Dim base_d2 As Double: base_d2 = wsLSMC.Range("B5").Value
    Dim r_US    As Double: r_US = wsEnc.Range("B4").Value
    Dim r_KRW   As Double: r_KRW = wsEnc.Range("B5").Value

    Dim T As Double: T = 1#
    Dim n_paths As Long: n_paths = 5000
    Dim factors() As Variant: factors = Array(0.5, 0.75, 1, 1.25, 1.5)

    Dim fIdx As Long
    For fIdx = LBound(factors) To UBound(factors)
        Dim factor As Double: factor = factors(fIdx)
        ' Scale only the wedge (P-drift minus Q-drift); Q-measure (r_US, r_KRW-r_US) fixed
        Dim d1_sim As Double: d1_sim = r_US + factor * (base_d1 - r_US)
        Dim d2_sim As Double: d2_sim = (r_KRW - r_US) + factor * (base_d2 - (r_KRW - r_US))

        ' FIX (2026-06-20c): T is fixed at 1 across all factors, and Beta_Mat
        ' priming only depends on Q-measure (factor-independent) ��� so prime
        ' once on the first iteration and reuse for the rest instead of
        ' re-pricing 5 times for an identical result.
        Dim sim_b As Double, theo_b As Double, ad_w As Double, ad_f As Double
        Call SimulateDriftBias(T, d1_sim, d2_sim, n_paths, sim_b, theo_b, ad_w, ad_f, _
                                primeBeta:=(fIdx = LBound(factors)))

        wsSweep.Cells(fIdx + 2, 1).Value = factor
        wsSweep.Cells(fIdx + 2, 2).Value = sim_b
        wsSweep.Cells(fIdx + 2, 3).Value = theo_b
        wsSweep.Cells(fIdx + 2, 4).Value = sim_b - theo_b
    Next fIdx

    Dim lastRow As Long: lastRow = UBound(factors) - LBound(factors) + 2

    Dim cht As ChartObject
    Set cht = wsSweep.ChartObjects.Add(Left:=350, Top:=10, Width:=500, Height:=320)
    cht.Chart.ChartType = xlXYScatterLinesNoMarkers
    Dim sSim As Series, sTheo As Series
    Set sSim = cht.Chart.SeriesCollection.NewSeries
    sSim.Name = "Simulated_Bias"
    sSim.XValues = wsSweep.Range("A2:A" & lastRow)
    sSim.Values = wsSweep.Range("B2:B" & lastRow)
    Set sTheo = cht.Chart.SeriesCollection.NewSeries
    sTheo.Name = "Theoretical_Bias"
    sTheo.XValues = wsSweep.Range("A2:A" & lastRow)
    sTheo.Values = wsSweep.Range("C2:C" & lastRow)
    With cht.Chart
        .HasTitle = True
        .ChartTitle.Text = "Bias vs P-Drift Scale Factor (linearity check)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Mu^P Scale Factor"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Mean Opt P&L Bias"
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
    End With

    wsSweep.Columns("A:D").AutoFit

    ' FIX (2026-06-20c): undo Beta_Mat pollution before returning control.
    Call RestoreProductionBetaMat

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Sharpe_Sweep_Diagnostic complete (real MC, " & _
           (UBound(factors) - LBound(factors) + 1) & " factors x " & n_paths & " paths)." & vbCrLf & _
           "Beta_Mat restored to production state." & vbCrLf & _
           "See 'SharpeSweep' sheet.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_Sharpe_Sweep_Diagnostic failed: " & Err.Description, vbCritical
End Sub

Public Sub Run_Girsanov_Reweight_Diagnostic()
    ' 16-3: Simulate under Q-measure, reweight terminal payoff by the diffusion
    ' Girsanov likelihood ratio, and compare the reweighted mean against a
    ' DIRECT P-measure simulation (ground truth, no reweighting).
    ' Jumps are untouched by the reweighting (P=Q for jumps, ��3 diversifiable-
    ' jump assumption) ��� only the diffusion shocks carry the LR.
    '
    ' FIX (2026-06-20c): the LR previously applied theta2 directly to e2
    ' (= corr*z1 + sqrt(1-corr^2)*z2), but z1 and e2 are NOT independent
    ' (corr(z1,e2) = corr), so exp(-theta1*z1-theta2*e2-...) double-counts the
    ' z1 component and is not a valid Girsanov density. Girsanov requires
    ' shifting the TRULY independent driving shocks (z1, z2). The independent
    ' shift on z2 that reproduces an exact theta2 shift in e2 is:
    '   theta2_indep = (theta2 - corr*theta1) / sqrt(1-corr^2)
    ' LR is now built from (z1, theta1) and (z2, theta2_indep) instead of
    ' (z1, theta1) and (e2, theta2). Also Steps moved to daily (260) to match
    ' the other engines.
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsEnc  As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsG    As Worksheet
    ' FIX (2026-06-21f, perf): toggles added ��� see MODEL_SPEC ��19.6.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail

    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Girsanov").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail
    Set wsG = Sheets.Add(After:=Sheets(Sheets.Count))
    wsG.Name = "Girsanov"

    Dim Lambda   As Double: Lambda = wsLSMC.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLSMC.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsLSMC.Range("B3").Value
    Dim Drift1P  As Double: Drift1P = wsLSMC.Range("B4").Value
    Dim Drift2P  As Double: Drift2P = wsLSMC.Range("B5").Value
    Dim KOUpper  As Double: KOUpper = wsLSMC.Range("B6").Value
    Dim KOLower  As Double: KOLower = wsLSMC.Range("B7").Value
    Dim S1_0     As Double: S1_0 = wsLSMC.Range("B8").Value
    Dim S2_0     As Double: S2_0 = wsLSMC.Range("B9").Value
    Dim vol1     As Double: vol1 = wsLSMC.Range("B10").Value
    Dim vol2     As Double: vol2 = wsLSMC.Range("B11").Value
    Dim corr     As Double: corr = wsLSMC.Range("B12").Value

    Dim r_US  As Double: r_US = wsEnc.Range("B4").Value
    Dim r_KRW As Double: r_KRW = wsEnc.Range("B5").Value
    ' REVISED 2026-06-21c: was K_WTI=Encoding!B18 (113, a stress-shock input,
    ' not the strike); a 2026-06-21b pass over-corrected to the European
    ' 0.95*Spot convention. Optimization_English.docx ��II confirms American/
    ' LSMC is priced ATM. See MODEL_SPEC ��17.1.
    Dim K_WTI As Double: K_WTI = S1_0   ' ATM, matches Optimization_English.docx

    Dim theta1 As Double: theta1 = (Drift1P - r_US) / vol1
    Dim theta2 As Double: theta2 = (Drift2P - (r_KRW - r_US)) / vol2

    ' FIX (2026-06-20c): independent-shock shift for z2 (see header note above).
    Dim theta2_indep As Double: theta2_indep = (theta2 - corr * theta1) / Sqr(1 - corr ^ 2)

    Dim kappa As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim T     As Double: T = 1#
    Dim Steps As Long:   Steps = CLng(T * STEPS_PER_YEAR)   ' daily grid
    Dim dt    As Double: dt = T / Steps
    Dim sqdt  As Double: sqdt = Sqr(dt)
    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt
    Dim n_paths As Long: n_paths = 5000

    Dim p As Long, i As Long

    ' --- Pass 1: Q-measure paths, accumulate Girsanov LR, reweight payoff ---
    Randomize 12345
    Dim sum_rw As Double, sum_rw2 As Double
    For p = 1 To n_paths
        Dim s1      As Double: s1 = S1_0
        Dim S2      As Double: S2 = S2_0
        Dim S1_prev As Double: S1_prev = S1_0
        Dim isAlive As Boolean: isAlive = True
        Dim LR      As Double: LR = 1#

       For i = 1 To Steps
            ' 1. ������(LR) ������ ������ independent diffusion ���� ���� ������ �������� �������� �� ���� ���������� ����������.
            Dim z1   As Double: z1 = GetNormal()
            Dim z2   As Double: z2 = GetNormal()
            
            ' 2. [CRITICAL FIX]: Girsanov Radon-Nikodym derivative sign correction
            ' ���� ��(z1, z2 > 0)�� ���� ������ ����(P)�� ���� ���������� �������� ����
            ' �������� ������ ���� ��������(-)���� ������(+)�� ����������.
            LR = LR * Exp(theta1 * z1 * sqdt - 0.5 * theta1 ^ 2 * dt _
                          + theta2_indep * z2 * sqdt - 0.5 * theta2_indep ^ 2 * dt)

            ' 3. ������ ���� ���� �������� �� KO ������ ������ ������ �������� ���� ����������.
            If isAlive Then
                S1_prev = s1
                Dim e2   As Double: e2 = corr * z1 + Sqr(1# - corr ^ 2) * z2
                Dim nJ   As Long:   nJ = GetPoisson(Lambda * dt)
                Dim jSum As Double: jSum = JumpSum(nJ, JumpMean, JumpVol)

                s1 = s1 * Exp((r_US - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt + vol1 * sqdt * z1 + jSum)
                S2 = S2 * Exp(((r_KRW - r_US) - 0.5 * vol2 ^ 2) * dt + vol2 * sqdt * e2)

                If s1 >= KOUpper Or s1 <= KOLower Then
                    isAlive = False
                ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then
                    Dim p_up As Double, p_dn_jd As Double
                    p_up = Exp(-2# * Log(KOUpper / S1_prev) * Log(KOUpper / s1) / lnVar)
                    p_dn_jd = Exp(-2# * Log(S1_prev / KOLower) * Log(s1 / KOLower) / lnVar)
                    If Rnd() < p_up Then isAlive = False
                    If isAlive And Rnd() < p_dn_jd Then isAlive = False
                End If
            End If
        Next i


        Dim payoff_Q As Double
        payoff_Q = IIf(isAlive, WorksheetFunction.Max(s1 - K_WTI, 0#) * S2, 0#)
        Dim rw As Double: rw = payoff_Q * LR
        sum_rw = sum_rw + rw
        sum_rw2 = sum_rw2 + rw ^ 2
    Next p

    Dim reweighted_mean As Double: reweighted_mean = sum_rw / n_paths
    Dim reweighted_var  As Double: reweighted_var = sum_rw2 / n_paths - reweighted_mean ^ 2

    ' --- Pass 2: DIRECT P-measure simulation (ground truth, no reweighting) ---
    Randomize 12345
    Dim sum_p As Double, sum_p2 As Double
    For p = 1 To n_paths
        Dim S1b      As Double: S1b = S1_0
        Dim S2b      As Double: S2b = S2_0
        Dim S1b_prev As Double: S1b_prev = S1_0
        Dim isAliveB As Boolean: isAliveB = True

        For i = 1 To Steps
            If isAliveB Then
                S1b_prev = S1b
                Dim z1b   As Double: z1b = GetNormal()
                Dim z2b   As Double: z2b = GetNormal()
                Dim e2b   As Double: e2b = corr * z1b + Sqr(1 - corr ^ 2) * z2b
                Dim nJb   As Long:   nJb = GetPoisson(Lambda * dt)
                Dim jSumb As Double: jSumb = JumpSum(nJb, JumpMean, JumpVol)

                S1b = S1b * Exp((Drift1P - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt + vol1 * sqdt * z1b + jSumb)
                S2b = S2b * Exp((Drift2P - 0.5 * vol2 ^ 2) * dt + vol2 * sqdt * e2b)

                If S1b >= KOUpper Or S1b <= KOLower Then
                    isAliveB = False
                ElseIf nJb = 0 And lnVar > 0 And S1b_prev > 0 Then
                    Dim p_upb As Double, p_dnb As Double
                    p_upb = Exp(-2# * Log(KOUpper / S1b_prev) * Log(KOUpper / S1b) / lnVar)
                    p_dnb = Exp(-2# * Log(S1b_prev / KOLower) * Log(S1b / KOLower) / lnVar)
                    If Rnd() < p_upb Then isAliveB = False
                    If isAliveB And Rnd() < p_dnb Then isAliveB = False
                End If
            End If
        Next i

        Dim payoff_P As Double
        payoff_P = IIf(isAliveB, WorksheetFunction.Max(S1b - K_WTI, 0#) * S2b, 0#)
        sum_p = sum_p + payoff_P
        sum_p2 = sum_p2 + payoff_P ^ 2
    Next p

    Dim direct_mean As Double: direct_mean = sum_p / n_paths
    Dim direct_var  As Double: direct_var = sum_p2 / n_paths - direct_mean ^ 2

    ' --- Output ---
    wsG.Range("A1:C1").Value = Array("Method", "Mean_Payoff", "StdDev_Payoff")
    wsG.Range("A2").Value = "Reweighted (Q-paths x Girsanov LR)"
    wsG.Range("B2").Value = reweighted_mean
    wsG.Range("C2").Value = Sqr(WorksheetFunction.Max(reweighted_var, 0#))
    wsG.Range("A3").Value = "Direct (P-measure simulation)"
    wsG.Range("B3").Value = direct_mean
    wsG.Range("C3").Value = Sqr(WorksheetFunction.Max(direct_var, 0#))
    wsG.Range("A4").Value = "Residual (Mean diff)"
    wsG.Range("B4").Value = reweighted_mean - direct_mean
    wsG.Range("A5").Value = "Residual (% of Direct Mean)"
    wsG.Range("B5").Value = (reweighted_mean - direct_mean) / direct_mean
    wsG.Range("B5").NumberFormat = "0.00%"

    Dim cht As ChartObject
    Set cht = wsG.ChartObjects.Add(Left:=350, Top:=10, Width:=400, Height:=300)
    cht.Chart.ChartType = xlColumnClustered
    Dim sBar As Series
    Set sBar = cht.Chart.SeriesCollection.NewSeries
    sBar.Name = "Mean Payoff"
    sBar.XValues = wsG.Range("A2:A3")
    sBar.Values = wsG.Range("B2:B3")
    With cht.Chart
        .HasTitle = True
        .ChartTitle.Text = "Girsanov Reweighting vs Direct P-Simulation"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Mean Terminal Payoff"
        .HasLegend = False
    End With

    wsG.Columns("A:C").AutoFit

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Girsanov_Reweight_Diagnostic complete (real reweighting vs direct P-sim, " & _
           n_paths & " paths each)." & vbCrLf & _
           "Reweighted mean: " & Format(reweighted_mean, "#,##0.0000") & vbCrLf & _
           "Direct P mean:   " & Format(direct_mean, "#,##0.0000") & vbCrLf & _
           "Residual:        " & Format(reweighted_mean - direct_mean, "#,##0.0000") & vbCrLf & vbCrLf & _
           "See 'Girsanov' sheet.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_Girsanov_Reweight_Diagnostic failed: " & Err.Description, vbCritical
End Sub

' #############################################################################
' ===== Originally: PaperIntegritySuite_revised.bas =====
' #############################################################################

' =============================================================================
' Paper-Integrity Sensitivity Suite (NEW, 2026-06-21f)
'
' Four independent items raised against ��16/��17.8's acknowledged limitations,
' each producing a defensible sensitivity table instead of leaving the
' assumption undefended:
'   1) Run_Lambda_Robustness_Sweep      -- k=2.5/3.0/3.5 jump-calibration
'                                          robustness (��16) + jump-gap-risk
'                                          quantification (��17.8-F)
'   2) Run_FD_vs_Regression_Delta_Check -- Beta_Mat regression delta vs true
'                                          finite-difference delta (��17.8)
'   3) Run_DeltaFX_Ratio_Sweep          -- delta_FX = c*delta_WTI sensitivity
'                                          (��17.8)
'   4) Run_JumpRiskPremium_Sensitivity  -- approximate re-pricing sweep under
'                                          jump risk premium != 0 (��17.8)
' Run_Paper_Integrity_Suite chains all 4 (mirrors Run_ITM_Diagnostic's chain
' pattern above, now in this same merged file).
'
' PERFORMANCE DESIGN (see MODEL_SPEC ��19.6/��20 for the full discussion):
'   - All 4 items use reduced n_paths/SIM_RUNS (20,000) vs production
'     (50,000/100,000) -- these are comparison sweeps, not headline numbers.
'   - Item 3 runs ONE price-path pass per simulation and tracks 5 parallel
'     FX-cost accumulators (c only affects the FX leg, not WTI/exercise
'     decisions) instead of 5 independent full reruns.
'   - Item 1 reuses a shared headless hedge-sim helper (RunHeadlessHedgeSim)
'     with NO per-iteration Excel I/O -- only a handful of summary writes
'     after each sweep point.
'   - Common random numbers (fixed Randomize seed per sweep) used throughout
'     so cross-scenario comparisons aren't muddied by independent sampling
'     noise.
'   - ScreenUpdating/Calculation/EnableEvents toggled off for the duration of
'     every Sub, matching the fixes applied elsewhere in this review pass.
'
' CORRECTNESS NOTE: every per-step formula below (BB correction, fade
' smoothing, exercise decision, futures MTM) is copied verbatim from the
' already-reviewed American_Delta_revised.bas / LSMC_revised.bas (now merged
' into DeltaHedging_revised.bas), not reinvented, to minimize the chance of
' introducing new formula bugs.
'
' UNTESTED IN EXCEL: this module was written and reviewed without the ability
' to execute VBA/Excel in this environment. Before citing any of its output
' in the paper, validate at minimum:
'   - Run_Lambda_Robustness_Sweep's k=3.0 row should be close to the current
'     production Base_Premium / KO rate (it uses the same parameters).
'   - Run_DeltaFX_Ratio_Sweep's c=1.0 row should be statistically close to a
'     full Run_American_DeltaHedge run (same parameters, same delta_FX=delta_WTI).
' See MODEL_SPEC ��19/��20.
' =============================================================================

' =============================================================================
' RunHeadlessHedgeSim ��� shared core for Run_Lambda_Robustness_Sweep.
' Single delta_FX=delta_WTI (c=1) Monte Carlo hedge-sim, no Excel I/O, returns
' only summary stats. Mirrors Run_American_DeltaHedge's per-sim loop exactly.
' Assumes T_WTI = T_total (matches current production calibration -- see
' MODEL_SPEC ��19.4's acknowledged limitation on partial-horizon optionality).
' =============================================================================
Private Sub RunHeadlessHedgeSim(ByVal Lambda As Double, ByVal JumpMean As Double, ByVal JumpVol As Double, _
                                 ByVal Drift1 As Double, ByVal Drift2 As Double, _
                                 ByVal KO_up As Double, ByVal KO_dn As Double, _
                                 ByVal S0_WTI As Double, ByVal S0_FX As Double, _
                                 ByVal vol1 As Double, ByVal vol2 As Double, ByVal corr As Double, _
                                 ByVal T_total As Double, ByVal hedge_steps As Long, _
                                 ByVal wti_cont As Double, ByVal WACC As Double, _
                                 ByVal total_premium As Double, _
                                 ByVal n_sim As Long, ByVal seed As Long, _
                                 ByRef mean_tot As Double, ByRef std_tot As Double, _
                                 ByRef var95 As Double, ByRef cvar95 As Double, _
                                 ByRef ko_rate As Double, ByRef ex_rate As Double)

    Const WTI_CONTRACT As Long = 1000
    Const FX_CONTRACT As Double = 100000
    Dim BUFFER As Double: BUFFER = 0.05
    Dim eps As Double: eps = 0.01

    Dim dt As Double: dt = T_total / hedge_steps
    Dim kappa As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt

    ' t=0 delta via finite difference (n_sim paths -- reduced vs production,
    ' this is a sweep point not the headline number)
    Dim V0_base As Double, V0_up1 As Double
    V0_base = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KO_up, KO_dn, _
                               S0_WTI, S0_FX, vol1, vol2, corr, hedge_steps, T_total, n_sim, S0_WTI)
    V0_up1 = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KO_up, KO_dn, _
                              S0_WTI * (1 + eps), S0_FX, vol1, vol2, corr, hedge_steps, T_total, n_sim, S0_WTI)
    Dim d0_WTI As Double: d0_WTI = ((V0_up1 - V0_base) / (S0_WTI * eps)) / S0_FX
    If d0_WTI < 0# Then d0_WTI = 0#
    If d0_WTI > 1# Then d0_WTI = 1#

    Dim tot_profits() As Double: ReDim tot_profits(1 To n_sim)
    Dim ko_flags() As Boolean: ReDim ko_flags(1 To n_sim)
    Dim ex_flags() As Boolean: ReDim ex_flags(1 To n_sim)

    Dim sim As Long, stepIdx As Long
    Dim S_WTI As Double, S_FX As Double, prev_S_WTI As Double
    Dim delta_WTI As Double, delta_FX As Double
    Dim pos_WTI As Double, pos_FX As Double, prev_WTI As Double, prev_FX As Double
    Dim margin_WTI As Double, cumul_FX As Double, cumul_cost As Double
    Dim mtm_WTI As Double, int_margin As Double, cost_FX As Double, int_FX As Double
    Dim z1 As Double, z2 As Double, e2 As Double, nJ As Long
    Dim tau As Double, remain_steps As Long
    Dim V_base As Double, fade As Double, intrinsic As Double, wtiGain As Double
    Dim knocked_out As Boolean, exercised As Boolean
    Dim S_WTI_ex As Double, S_FX_ex As Double, effective_WTI As Double
    Dim raw_dWTI As Double, p_bb_up As Double, p_bb_dn As Double
    Dim v1_norm As Double, v2_norm As Double
    Dim b0 As Double, b1 As Double, b2 As Double, b3 As Double, b4 As Double, b5 As Double
    Dim py As Double, opt_profit As Double, tot_profit As Double

    Randomize seed   ' common random numbers: caller passes the same seed for every sweep point it wants compared cleanly

    For sim = 1 To n_sim
        S_WTI = S0_WTI: S_FX = S0_FX: prev_S_WTI = S0_WTI
        knocked_out = False: exercised = False
        delta_WTI = d0_WTI: delta_FX = d0_WTI

        pos_WTI = delta_WTI * wti_cont
        pos_FX = delta_FX * S0_WTI * wti_cont * WTI_CONTRACT / FX_CONTRACT
        prev_WTI = pos_WTI: prev_FX = pos_FX
        margin_WTI = 0#
        cumul_FX = pos_FX * FX_CONTRACT * S0_FX
        cumul_cost = cumul_FX - margin_WTI

        For stepIdx = 1 To hedge_steps
            If exercised Or knocked_out Then GoTo NextStepH

            z1 = GetNormal(): z2 = GetNormal()
            e2 = corr * z1 + Sqr(1# - corr ^ 2) * z2
            nJ = GetPoisson(Lambda * dt)

            S_WTI = S_WTI * Exp((Drift1 - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt + vol1 * Sqr(dt) * z1 + JumpSum(nJ, JumpMean, JumpVol))
            S_FX = S_FX * Exp((Drift2 - 0.5 * vol2 ^ 2) * dt + vol2 * Sqr(dt) * e2)
            effective_WTI = S_WTI   ' sweep assumption: T_WTI = T_total (see MODEL_SPEC ��19.4)

            If effective_WTI >= KO_up Or effective_WTI <= KO_dn Then
                knocked_out = True: delta_WTI = 0#: delta_FX = 0#: GoTo WriteH
            ElseIf nJ = 0 Then
                If lnVar > 0 And prev_S_WTI > 0 Then
                    p_bb_up = Exp(-2# * Log(KO_up / prev_S_WTI) * Log(KO_up / S_WTI) / lnVar)
                    p_bb_dn = Exp(-2# * Log(prev_S_WTI / KO_dn) * Log(S_WTI / KO_dn) / lnVar)
                    If Rnd() < p_bb_up Then
                        knocked_out = True: delta_WTI = 0#: delta_FX = 0#: GoTo WriteH
                    End If
                    If Rnd() < p_bb_dn Then
                        knocked_out = True: delta_WTI = 0#: delta_FX = 0#: GoTo WriteH
                    End If
                End If
            End If

            tau = T_total - stepIdx * dt
            remain_steps = hedge_steps - stepIdx

            If tau > 0.001 And remain_steps > 0 Then
                v1_norm = effective_WTI / S0_WTI
                v2_norm = S_FX / S0_FX
                b0 = Beta_Mat(stepIdx, 0): b1 = Beta_Mat(stepIdx, 1): b2 = Beta_Mat(stepIdx, 2)
                b3 = Beta_Mat(stepIdx, 3): b4 = Beta_Mat(stepIdx, 4): b5 = Beta_Mat(stepIdx, 5)
                V_base = b0 + b1 * v1_norm + b2 * v2_norm + b3 * (v1_norm ^ 2) + b4 * (v2_norm ^ 2) + b5 * (v1_norm * v2_norm)
                If V_base < 0# Then V_base = 0#

                raw_dWTI = ((b1 + 2# * b3 * v1_norm + b5 * v2_norm) / S0_WTI) / S_FX
                If raw_dWTI < 0# Then
                    delta_WTI = 0#
                ElseIf raw_dWTI > 1# Then
                    delta_WTI = 1#
                Else
                    delta_WTI = raw_dWTI
                End If
                delta_FX = delta_WTI

                If effective_WTI > KO_up * (1# - BUFFER) Then
                    fade = (KO_up - effective_WTI) / (KO_up * BUFFER)
                    If fade < 0# Then fade = 0#
                    delta_WTI = delta_WTI * fade: delta_FX = delta_FX * fade
                End If
                If effective_WTI < KO_dn * (1# + BUFFER) Then
                    fade = (effective_WTI - KO_dn) / (KO_dn * BUFFER)
                    If fade < 0# Then fade = 0#
                    delta_WTI = delta_WTI * fade: delta_FX = delta_FX * fade
                End If

                wtiGain = effective_WTI - S0_WTI
                If wtiGain < 0# Then wtiGain = 0#
                intrinsic = wtiGain * S_FX * wti_cont
                If intrinsic > V_base * wti_cont And intrinsic > 0# Then
                    exercised = True: S_WTI_ex = effective_WTI: S_FX_ex = S_FX
                    delta_WTI = 0#: delta_FX = 0#
                End If
            Else
                delta_WTI = IIf(effective_WTI > S0_WTI, 1#, 0#)
                delta_FX = delta_WTI
            End If

WriteH:
            pos_WTI = delta_WTI * wti_cont
            pos_FX = delta_FX * IIf(S_WTI < S0_WTI, S_WTI, S0_WTI) * wti_cont * WTI_CONTRACT / FX_CONTRACT

            mtm_WTI = prev_WTI * (S_WTI - prev_S_WTI) * WTI_CONTRACT * S_FX
            margin_WTI = margin_WTI + mtm_WTI
            int_margin = margin_WTI * WACC * dt
            margin_WTI = margin_WTI + int_margin

            cost_FX = (pos_FX - prev_FX) * FX_CONTRACT * S_FX
            int_FX = cumul_FX * WACC * dt
            cumul_FX = cumul_FX + cost_FX + int_FX
            cumul_cost = cumul_FX - margin_WTI

            prev_S_WTI = S_WTI
            prev_WTI = pos_WTI
            prev_FX = pos_FX

NextStepH:
        Next stepIdx

        If exercised Then
            wtiGain = S_WTI_ex - S0_WTI: If wtiGain < 0# Then wtiGain = 0#
            py = wtiGain * S_FX_ex * WTI_CONTRACT * wti_cont
        ElseIf Not knocked_out Then
            wtiGain = effective_WTI - S0_WTI: If wtiGain < 0# Then wtiGain = 0#
            py = wtiGain * S_FX * WTI_CONTRACT * wti_cont
        Else
            py = 0#
        End If

        opt_profit = total_premium - py
        tot_profit = opt_profit - cumul_cost

        tot_profits(sim) = tot_profit
        ko_flags(sim) = knocked_out
        ex_flags(sim) = exercised
    Next sim

    Dim ko_count As Long, ex_count As Long, s As Long
    For s = 1 To n_sim
        If ko_flags(s) Then ko_count = ko_count + 1
        If ex_flags(s) Then ex_count = ex_count + 1
    Next s
    ko_rate = ko_count / n_sim
    ex_rate = ex_count / n_sim

    mean_tot = Application.WorksheetFunction.Average(tot_profits)
    std_tot = Application.WorksheetFunction.StDev(tot_profits)
    Dim p5 As Double: p5 = Application.WorksheetFunction.Percentile(tot_profits, 0.05)
    var95 = -p5   ' loss-positive convention

    Dim sumTail As Double, cntTail As Long
    For s = 1 To n_sim
        If tot_profits(s) <= p5 Then
            sumTail = sumTail + tot_profits(s)
            cntTail = cntTail + 1
        End If
    Next s
    If cntTail > 0 Then
        cvar95 = -(sumTail / cntTail)
    Else
        cvar95 = var95
    End If
End Sub

' =============================================================================
' Item 1 (��16, ��17.8-F): k=2.5/3.0/3.5 jump-calibration robustness, plus the
' closed-form jump-gap-risk number P(>=1 jump in horizon)=1-exp(-Lambda*T) for
' each k (itself Lambda-dependent, hence reported per-k rather than once).
' =============================================================================
Public Sub Run_Lambda_Robustness_Sweep()
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsLR As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("LambdaRobustness").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail
    Set wsLR = Sheets.Add(After:=Sheets(Sheets.Count))
    wsLR.Name = "LambdaRobustness"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim Drift1 As Double: Drift1 = wsLSMC.Range("B4").Value
    Dim Drift2 As Double: Drift2 = wsLSMC.Range("B5").Value
    Dim KOUpper As Double: KOUpper = wsLSMC.Range("B6").Value
    Dim KOLower As Double: KOLower = wsLSMC.Range("B7").Value
    Dim S1_0 As Double: S1_0 = wsLSMC.Range("B8").Value
    Dim S2_0 As Double: S2_0 = wsLSMC.Range("B9").Value
    Dim vol2 As Double: vol2 = wsLSMC.Range("B11").Value
    Dim corr As Double: corr = wsLSMC.Range("B12").Value
    Dim T_total As Double: T_total = wsLSMC.Range("B14").Value
    Dim WACC As Double: WACC = wsEnc.Range("B12").Value
    Dim barrels As Double: barrels = wsEnc.Range("B9").Value
    Const WTI_CONTRACT As Long = 1000
    Dim wti_cont As Double: wti_cont = barrels / WTI_CONTRACT

    Dim Steps As Long: Steps = CLng(T_total * STEPS_PER_YEAR)

    ' ��16 robustness table (k, vol1, Lambda, JumpMean, JumpVol) -- hardcoded
    ' from the already-validated decomposition, not re-derived here.
    Dim kVals() As Variant: kVals = Array(2.5, 3#, 3.5)
    Dim vol1Vals() As Variant: vol1Vals = Array(0.2866, 0.3218, 0.3416)
    Dim LambdaVals() As Variant: LambdaVals = Array(15.5, 6.7794, 3.68)
    Dim JumpMeanVals() As Variant: JumpMeanVals = Array(-0.0174, -0.029895, -0.0492)
    Dim JumpVolVals() As Variant: JumpVolVals = Array(0.0685, 0.084429, 0.0926)

    Dim n_sweep As Long: n_sweep = 20000
    Dim sweepSeed As Long: sweepSeed = 20260621

    wsLR.Range("A1:I1").Value = Array("k", "vol1", "Lambda", "JumpMean", "JumpVol", _
                                       "Base_Premium", "P_atLeast1Jump", "KO_Rate", "Mean_Tot_PL")
    wsLR.Range("J1:L1").Value = Array("Std_Tot_PL", "VaR95_Tot_PL", "CVaR95_Tot_PL")

    Dim kIdx As Long
    For kIdx = LBound(kVals) To UBound(kVals)
        Dim Lambda_k As Double: Lambda_k = LambdaVals(kIdx)
        Dim JumpMean_k As Double: JumpMean_k = JumpMeanVals(kIdx)
        Dim JumpVol_k As Double: JumpVol_k = JumpVolVals(kIdx)
        Dim vol1_k As Double: vol1_k = vol1Vals(kIdx)

        Dim Base_Premium_k As Double
        Base_Premium_k = Calc_LSMC_Price(Lambda_k, JumpMean_k, JumpVol_k, Drift1, Drift2, KOUpper, KOLower, _
                                          S1_0, S2_0, vol1_k, vol2, corr, Steps, T_total, n_sweep, S1_0, True)

        Dim pAtLeast1Jump As Double: pAtLeast1Jump = 1# - Exp(-Lambda_k * T_total)
        Dim total_premium_k As Double: total_premium_k = Base_Premium_k * wti_cont

        Dim meanTot As Double, stdTot As Double, var95 As Double, cvar95 As Double, koRate As Double, exRate As Double
        Call RunHeadlessHedgeSim(Lambda_k, JumpMean_k, JumpVol_k, Drift1, Drift2, KOUpper, KOLower, _
                                  S1_0, S2_0, vol1_k, vol2, corr, T_total, Steps, wti_cont, WACC, _
                                  total_premium_k, n_sweep, sweepSeed, meanTot, stdTot, var95, cvar95, koRate, exRate)

        wsLR.Cells(kIdx + 2, 1).Value = kVals(kIdx)
        wsLR.Cells(kIdx + 2, 2).Value = vol1_k
        wsLR.Cells(kIdx + 2, 3).Value = Lambda_k
        wsLR.Cells(kIdx + 2, 4).Value = JumpMean_k
        wsLR.Cells(kIdx + 2, 5).Value = JumpVol_k
        wsLR.Cells(kIdx + 2, 6).Value = Base_Premium_k
        wsLR.Cells(kIdx + 2, 7).Value = pAtLeast1Jump
        wsLR.Cells(kIdx + 2, 7).NumberFormat = "0.0000%"
        wsLR.Cells(kIdx + 2, 8).Value = koRate
        wsLR.Cells(kIdx + 2, 8).NumberFormat = "0.0%"
        wsLR.Cells(kIdx + 2, 9).Value = meanTot
        wsLR.Cells(kIdx + 2, 10).Value = stdTot
        wsLR.Cells(kIdx + 2, 11).Value = var95
        wsLR.Cells(kIdx + 2, 12).Value = cvar95
    Next kIdx

    wsLR.Columns("A:L").AutoFit

    ' Restore Beta_Mat to the production state (the bStoreBeta:=True calls
    ' above overwrote it 3x with sweep-grid coefficients).
    Call RestoreProductionBetaMat

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Lambda Robustness Sweep complete (k=2.5/3.0/3.5, " & n_sweep & " paths each)." & vbCrLf & _
           "See 'LambdaRobustness' sheet. Beta_Mat restored to production state.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_Lambda_Robustness_Sweep failed: " & Err.Description, vbCritical
End Sub

' =============================================================================
' Item 2 (��17.8): Beta_Mat regression delta vs true finite-difference delta,
' at 3 residual-maturity points (t=0, mid-life, near-end).
' =============================================================================
Public Sub Run_FD_vs_Regression_Delta_Check()
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsChk As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("DeltaCheck").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail
    Set wsChk = Sheets.Add(After:=Sheets(Sheets.Count))
    wsChk.Name = "DeltaCheck"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim Lambda As Double: Lambda = wsLSMC.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLSMC.Range("B2").Value
    Dim JumpVol As Double: JumpVol = wsLSMC.Range("B3").Value
    Dim Drift1 As Double: Drift1 = wsLSMC.Range("B4").Value
    Dim Drift2 As Double: Drift2 = wsLSMC.Range("B5").Value
    Dim KOUpper As Double: KOUpper = wsLSMC.Range("B6").Value
    Dim KOLower As Double: KOLower = wsLSMC.Range("B7").Value
    Dim S1_0 As Double: S1_0 = wsLSMC.Range("B8").Value
    Dim S2_0 As Double: S2_0 = wsLSMC.Range("B9").Value
    Dim vol1 As Double: vol1 = wsLSMC.Range("B10").Value
    Dim vol2 As Double: vol2 = wsLSMC.Range("B11").Value
    Dim corr As Double: corr = wsLSMC.Range("B12").Value
    Dim T_total As Double: T_total = wsLSMC.Range("B14").Value

    Dim n_chk As Long: n_chk = 20000
    Dim eps As Double: eps = 0.01

    wsChk.Range("A1:F1").Value = Array("t_elapsed_frac", "Residual_T", "FD_Delta", "Regression_Delta", "Abs_Diff", "Pct_Diff_of_FD")

    Dim fracs() As Variant: fracs = Array(0#, 0.5, 0.9)
    Dim rIdx As Long
    For rIdx = LBound(fracs) To UBound(fracs)
        Dim tau As Double: tau = T_total * (1# - fracs(rIdx))
        If tau < 1# / STEPS_PER_YEAR Then tau = 1# / STEPS_PER_YEAR
        Dim subSteps As Long: subSteps = CLng(tau * STEPS_PER_YEAR)
        If subSteps < 2 Then subSteps = 2

        Dim V_base As Double
        V_base = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KOUpper, KOLower, _
                                  S1_0, S2_0, vol1, vol2, corr, subSteps, tau, n_chk, S1_0, True)

        Dim b1 As Double: b1 = Beta_Mat(1, 1)
        Dim b3 As Double: b3 = Beta_Mat(1, 3)
        Dim b5 As Double: b5 = Beta_Mat(1, 5)
        Dim regDelta As Double: regDelta = ((b1 + 2# * b3 + b5) / S1_0) / S2_0

        Dim V_up As Double, V_dn As Double
        V_up = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KOUpper, KOLower, _
                                S1_0 * (1# + eps), S2_0, vol1, vol2, corr, subSteps, tau, n_chk, S1_0)
        V_dn = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KOUpper, KOLower, _
                                S1_0 * (1# - eps), S2_0, vol1, vol2, corr, subSteps, tau, n_chk, S1_0)
        Dim fdDelta As Double: fdDelta = ((V_up - V_dn) / (2# * eps * S1_0)) / S2_0

        Dim absDiff As Double: absDiff = Abs(fdDelta - regDelta)
        Dim pctDiff As Double
        If Abs(fdDelta) > 1E-07 Then
            pctDiff = absDiff / Abs(fdDelta)
        Else
            pctDiff = 0#
        End If

        wsChk.Cells(rIdx + 2, 1).Value = fracs(rIdx)
        wsChk.Cells(rIdx + 2, 2).Value = tau
        wsChk.Cells(rIdx + 2, 3).Value = fdDelta
        wsChk.Cells(rIdx + 2, 4).Value = regDelta
        wsChk.Cells(rIdx + 2, 5).Value = absDiff
        wsChk.Cells(rIdx + 2, 6).Value = pctDiff
        wsChk.Cells(rIdx + 2, 6).NumberFormat = "0.0%"
    Next rIdx

    wsChk.Range("A6").Value = "Flag rows where Pct_Diff_of_FD > 5% -- regression delta unreliable there (see ��17.8)."
    wsChk.Columns("A:F").AutoFit

    Call RestoreProductionBetaMat

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "FD vs Regression Delta check complete (3 residual-maturity points)." & vbCrLf & _
           "See 'DeltaCheck' sheet. Beta_Mat restored to production state.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_FD_vs_Regression_Delta_Check failed: " & Err.Description, vbCritical
End Sub

' =============================================================================
' Item 3 (��17.8): is delta_FX = delta_WTI (c=1, hard-coded in
' Run_American_DeltaHedge) actually near risk-optimal, or arbitrary? Sweeps
' c in {0.5,0.75,1,1.25,1.5}: delta_FX = c * delta_WTI.
'
' PERFORMANCE: WTI/FX price evolution, BB-KO, Beta_Mat delta, and exercise
' decisions do NOT depend on c (only the FX hedge's cost accounting does) --
' so this runs the expensive part ONCE per simulated path and tracks 5
' parallel FX-cost accumulators, instead of 5 independent full reruns.
' opt_profit is therefore IDENTICAL across all 5 c's by construction (only
' tot_profit, which nets out hedge cost, differs) -- true common random
' numbers with zero added noise on the option P&L side.
'
' Requires Beta_Mat already primed for this hedge grid (same guard as
' Run_American_DeltaHedge ��� run Run_LSMC_Engine first).
' =============================================================================
Public Sub Run_DeltaFX_Ratio_Sweep()
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsCS As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("DeltaFXSweep").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail
    Set wsCS = Sheets.Add(After:=Sheets(Sheets.Count))
    wsCS.Name = "DeltaFXSweep"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim Lambda As Double: Lambda = wsLSMC.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLSMC.Range("B2").Value
    Dim JumpVol As Double: JumpVol = wsLSMC.Range("B3").Value
    Dim Drift1 As Double: Drift1 = wsLSMC.Range("B4").Value
    Dim Drift2 As Double: Drift2 = wsLSMC.Range("B5").Value
    Dim KOUpper As Double: KOUpper = wsLSMC.Range("B6").Value
    Dim KOLower As Double: KOLower = wsLSMC.Range("B7").Value
    Dim S0_WTI As Double: S0_WTI = wsLSMC.Range("B8").Value
    Dim S0_FX As Double: S0_FX = wsLSMC.Range("B9").Value
    Dim vol1 As Double: vol1 = wsLSMC.Range("B10").Value
    Dim vol2 As Double: vol2 = wsLSMC.Range("B11").Value
    Dim corr As Double: corr = wsLSMC.Range("B12").Value
    Dim T_total As Double: T_total = wsLSMC.Range("B14").Value
    Dim WACC As Double: WACC = wsEnc.Range("B12").Value
    Dim barrels As Double: barrels = wsEnc.Range("B9").Value
    Dim Base_Premium As Double: Base_Premium = wsLSMC.Range("J2").Value

    Const WTI_CONTRACT As Long = 1000
    Const FX_CONTRACT As Double = 100000
    Dim wti_cont As Double: wti_cont = barrels / WTI_CONTRACT
    Dim total_premium As Double: total_premium = Base_Premium * wti_cont

    Dim hedge_steps As Long: hedge_steps = CLng(T_total * STEPS_PER_YEAR)
    Dim dt As Double: dt = T_total / hedge_steps
    Dim kappa As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt
    Dim BUFFER As Double: BUFFER = 0.05
    Dim eps As Double: eps = 0.01

    ' Beta_Mat guard. FIX (2026-06-22e): reads the Beta_Mat_Steps Long flag
    ' (set in Calc_LSMC_Price) instead of UBound(Beta_Mat,1), which raises a
    ' (normally-handled) error that Tools>Options "Break on All Errors" turns
    ' into a debugger break -- see MODEL_SPEC ��30. Still fails loud here (this
    ' is a paper-verification tool meant to run after Run_LSMC_Engine, not a
    ' standalone production engine like American which self-primes per ��29).
    If Beta_Mat_Steps <> hedge_steps Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        MsgBox "Beta_Mat is not primed for this hedge grid (expected " & hedge_steps & _
               ", got " & IIf(Beta_Mat_Steps = 0, "unallocated", CStr(Beta_Mat_Steps)) & ")." & vbCrLf & _
               "Run Run_LSMC_Engine first.", vbCritical, "Delta_FX Ratio Sweep ��� Beta_Mat mismatch"
        Exit Sub
    End If

    Dim cFactors() As Variant: cFactors = Array(0.5, 0.75, 1#, 1.25, 1.5)
    Dim nC As Long: nC = UBound(cFactors) - LBound(cFactors) + 1
    Dim n_sweep As Long: n_sweep = 20000

    Dim V0_base As Double, V0_up1 As Double
    V0_base = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KOUpper, KOLower, _
                               S0_WTI, S0_FX, vol1, vol2, corr, hedge_steps, T_total, n_sweep, S0_WTI)
    V0_up1 = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KOUpper, KOLower, _
                              S0_WTI * (1 + eps), S0_FX, vol1, vol2, corr, hedge_steps, T_total, n_sweep, S0_WTI)
    Dim d0_WTI As Double: d0_WTI = ((V0_up1 - V0_base) / (S0_WTI * eps)) / S0_FX
    If d0_WTI < 0# Then d0_WTI = 0#
    If d0_WTI > 1# Then d0_WTI = 1#

    Dim tot_profits() As Double: ReDim tot_profits(1 To nC, 1 To n_sweep)
    Dim ko_flags() As Boolean: ReDim ko_flags(1 To n_sweep)
    Dim ex_flags() As Boolean: ReDim ex_flags(1 To n_sweep)

    Randomize 20260621

    Dim sim As Long, stepIdx As Long, cIdx As Long
    Dim S_WTI As Double, S_FX As Double, prev_S_WTI As Double
    Dim delta_WTI As Double
    Dim pos_WTI As Double, prev_WTI As Double
    Dim margin_WTI As Double, mtm_WTI As Double, int_margin As Double
    Dim z1 As Double, z2 As Double, e2 As Double, nJ As Long
    Dim tau As Double, remain_steps As Long
    Dim V_base As Double, fade As Double, intrinsic As Double, wtiGain As Double
    Dim knocked_out As Boolean, exercised As Boolean
    Dim S_WTI_ex As Double, S_FX_ex As Double, effective_WTI As Double
    Dim raw_dWTI As Double, p_bb_up As Double, p_bb_dn As Double
    Dim v1_norm As Double, v2_norm As Double
    Dim b0 As Double, b1 As Double, b2 As Double, b3 As Double, b4 As Double, b5 As Double
    Dim py As Double, opt_profit As Double
    Dim delta_FX_c As Double, cost_FX As Double, int_FX As Double

    Dim pos_FX_c() As Double: ReDim pos_FX_c(1 To nC)
    Dim prev_FX_c() As Double: ReDim prev_FX_c(1 To nC)
    Dim cumul_FX_c() As Double: ReDim cumul_FX_c(1 To nC)
    Dim cumul_cost_c() As Double: ReDim cumul_cost_c(1 To nC)

    For sim = 1 To n_sweep
        S_WTI = S0_WTI: S_FX = S0_FX: prev_S_WTI = S0_WTI
        knocked_out = False: exercised = False
        delta_WTI = d0_WTI

        pos_WTI = delta_WTI * wti_cont
        prev_WTI = pos_WTI
        margin_WTI = 0#

        For cIdx = 1 To nC
            delta_FX_c = cFactors(LBound(cFactors) + cIdx - 1) * delta_WTI
            pos_FX_c(cIdx) = delta_FX_c * S0_WTI * wti_cont * WTI_CONTRACT / FX_CONTRACT
            prev_FX_c(cIdx) = pos_FX_c(cIdx)
            cumul_FX_c(cIdx) = pos_FX_c(cIdx) * FX_CONTRACT * S0_FX
            cumul_cost_c(cIdx) = cumul_FX_c(cIdx)
        Next cIdx

        For stepIdx = 1 To hedge_steps
            If exercised Or knocked_out Then GoTo NextStepC

            z1 = GetNormal(): z2 = GetNormal()
            e2 = corr * z1 + Sqr(1# - corr ^ 2) * z2
            nJ = GetPoisson(Lambda * dt)

            S_WTI = S_WTI * Exp((Drift1 - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt + vol1 * Sqr(dt) * z1 + JumpSum(nJ, JumpMean, JumpVol))
            S_FX = S_FX * Exp((Drift2 - 0.5 * vol2 ^ 2) * dt + vol2 * Sqr(dt) * e2)
            effective_WTI = S_WTI   ' sweep assumption: T_WTI = T_total (see MODEL_SPEC ��19.4)

            If effective_WTI >= KOUpper Or effective_WTI <= KOLower Then
                knocked_out = True: delta_WTI = 0#: GoTo WriteC
            ElseIf nJ = 0 Then
                If lnVar > 0 And prev_S_WTI > 0 Then
                    p_bb_up = Exp(-2# * Log(KOUpper / prev_S_WTI) * Log(KOUpper / S_WTI) / lnVar)
                    p_bb_dn = Exp(-2# * Log(prev_S_WTI / KOLower) * Log(S_WTI / KOLower) / lnVar)
                    If Rnd() < p_bb_up Then knocked_out = True: delta_WTI = 0#: GoTo WriteC
                    If Rnd() < p_bb_dn Then knocked_out = True: delta_WTI = 0#: GoTo WriteC
                End If
            End If

            tau = T_total - stepIdx * dt
            remain_steps = hedge_steps - stepIdx

            If tau > 0.001 And remain_steps > 0 Then
                v1_norm = effective_WTI / S0_WTI
                v2_norm = S_FX / S0_FX
                b0 = Beta_Mat(stepIdx, 0): b1 = Beta_Mat(stepIdx, 1): b2 = Beta_Mat(stepIdx, 2)
                b3 = Beta_Mat(stepIdx, 3): b4 = Beta_Mat(stepIdx, 4): b5 = Beta_Mat(stepIdx, 5)
                V_base = b0 + b1 * v1_norm + b2 * v2_norm + b3 * (v1_norm ^ 2) + b4 * (v2_norm ^ 2) + b5 * (v1_norm * v2_norm)
                If V_base < 0# Then V_base = 0#

                raw_dWTI = ((b1 + 2# * b3 * v1_norm + b5 * v2_norm) / S0_WTI) / S_FX
                If raw_dWTI < 0# Then
                    delta_WTI = 0#
                ElseIf raw_dWTI > 1# Then
                    delta_WTI = 1#
                Else
                    delta_WTI = raw_dWTI
                End If

                If effective_WTI > KOUpper * (1# - BUFFER) Then
                    fade = (KOUpper - effective_WTI) / (KOUpper * BUFFER)
                    If fade < 0# Then fade = 0#
                    delta_WTI = delta_WTI * fade
                End If
                If effective_WTI < KOLower * (1# + BUFFER) Then
                    fade = (effective_WTI - KOLower) / (KOLower * BUFFER)
                    If fade < 0# Then fade = 0#
                    delta_WTI = delta_WTI * fade
                End If

                wtiGain = effective_WTI - S0_WTI
                If wtiGain < 0# Then wtiGain = 0#
                intrinsic = wtiGain * S_FX * wti_cont
                If intrinsic > V_base * wti_cont And intrinsic > 0# Then
                    exercised = True: S_WTI_ex = effective_WTI: S_FX_ex = S_FX
                    delta_WTI = 0#
                End If
            Else
                delta_WTI = IIf(effective_WTI > S0_WTI, 1#, 0#)
            End If

WriteC:
            pos_WTI = delta_WTI * wti_cont
            Dim fx_notional As Double
            fx_notional = IIf(S_WTI < S0_WTI, S_WTI, S0_WTI) * wti_cont * WTI_CONTRACT

            mtm_WTI = prev_WTI * (S_WTI - prev_S_WTI) * WTI_CONTRACT * S_FX
            margin_WTI = margin_WTI + mtm_WTI
            int_margin = margin_WTI * WACC * dt
            margin_WTI = margin_WTI + int_margin

            For cIdx = 1 To nC
                delta_FX_c = cFactors(LBound(cFactors) + cIdx - 1) * delta_WTI
                pos_FX_c(cIdx) = delta_FX_c * fx_notional / FX_CONTRACT
                cost_FX = (pos_FX_c(cIdx) - prev_FX_c(cIdx)) * FX_CONTRACT * S_FX
                int_FX = cumul_FX_c(cIdx) * WACC * dt
                cumul_FX_c(cIdx) = cumul_FX_c(cIdx) + cost_FX + int_FX
                cumul_cost_c(cIdx) = cumul_FX_c(cIdx) - margin_WTI
                prev_FX_c(cIdx) = pos_FX_c(cIdx)
            Next cIdx

            prev_S_WTI = S_WTI
            prev_WTI = pos_WTI

NextStepC:
        Next stepIdx

        If exercised Then
            wtiGain = S_WTI_ex - S0_WTI: If wtiGain < 0# Then wtiGain = 0#
            py = wtiGain * S_FX_ex * WTI_CONTRACT * wti_cont
        ElseIf Not knocked_out Then
            wtiGain = effective_WTI - S0_WTI: If wtiGain < 0# Then wtiGain = 0#
            py = wtiGain * S_FX * WTI_CONTRACT * wti_cont
        Else
            py = 0#
        End If

        opt_profit = total_premium - py
        ko_flags(sim) = knocked_out
        ex_flags(sim) = exercised

        For cIdx = 1 To nC
            tot_profits(cIdx, sim) = opt_profit - cumul_cost_c(cIdx)
        Next cIdx
    Next sim

    wsCS.Range("A1:G1").Value = Array("c_factor", "Mean_Tot_PL", "Std_Tot_PL", "VaR95_Tot_PL", "CVaR95_Tot_PL", "KO_Rate", "Exercise_Rate")

    Dim ko_count As Long, ex_count As Long, s As Long
    For s = 1 To n_sweep
        If ko_flags(s) Then ko_count = ko_count + 1
        If ex_flags(s) Then ex_count = ex_count + 1
    Next s

    For cIdx = 1 To nC
        Dim oneC() As Double: ReDim oneC(1 To n_sweep)
        For s = 1 To n_sweep
            oneC(s) = tot_profits(cIdx, s)
        Next s

        Dim meanC As Double: meanC = Application.WorksheetFunction.Average(oneC)
        Dim stdC As Double: stdC = Application.WorksheetFunction.StDev(oneC)
        Dim p5C As Double: p5C = Application.WorksheetFunction.Percentile(oneC, 0.05)
        Dim var95C As Double: var95C = -p5C

        Dim sumTail As Double, cntTail As Long
        sumTail = 0#: cntTail = 0
        For s = 1 To n_sweep
            If oneC(s) <= p5C Then
                sumTail = sumTail + oneC(s)
                cntTail = cntTail + 1
            End If
        Next s
        Dim cvar95C As Double
        If cntTail > 0 Then cvar95C = -(sumTail / cntTail) Else cvar95C = var95C

        wsCS.Cells(cIdx + 1, 1).Value = cFactors(LBound(cFactors) + cIdx - 1)
        wsCS.Cells(cIdx + 1, 2).Value = meanC
        wsCS.Cells(cIdx + 1, 3).Value = stdC
        wsCS.Cells(cIdx + 1, 4).Value = var95C
        wsCS.Cells(cIdx + 1, 5).Value = cvar95C
        wsCS.Cells(cIdx + 1, 6).Value = ko_count / n_sweep
        wsCS.Cells(cIdx + 1, 6).NumberFormat = "0.0%"
        wsCS.Cells(cIdx + 1, 7).Value = ex_count / n_sweep
        wsCS.Cells(cIdx + 1, 7).NumberFormat = "0.0%"
    Next cIdx

    wsCS.Range("A9").Value = "NOTE: KO_Rate/Exercise_Rate identical across all c (c only affects the FX leg, not WTI/exercise decisions)."
    wsCS.Columns("A:G").AutoFit

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Delta_FX Ratio Sweep complete (c=0.5..1.5, " & n_sweep & " paths, single price-path pass)." & vbCrLf & _
           "See 'DeltaFXSweep' sheet.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_DeltaFX_Ratio_Sweep failed: " & Err.Description, vbCritical
End Sub

' =============================================================================
' Item 4 (��17.8): approximate sensitivity of Base_Premium/Stress_Premium to
' the jump-risk-premium=0 assumption. NOT a rigorous Girsanov jump-measure
' change (that needs an Esscher transform on both intensity and jump-size
' distribution to stay risk-neutral) -- this is a simple re-pricing under a
' scaled Lambda^Q, to show sensitivity if the true risk-neutral jump
' intensity differs from the historical (diversifiable-jump) estimate.
' Disclose in the paper as a sensitivity check, not a measure-change result.
' Never touches Beta_Mat (no bStoreBeta:=True calls), so no restore needed.
' =============================================================================
Public Sub Run_JumpRiskPremium_Sensitivity()
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsJR As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("JumpPremiumSens").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail
    Set wsJR = Sheets.Add(After:=Sheets(Sheets.Count))
    wsJR.Name = "JumpPremiumSens"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim Lambda0 As Double: Lambda0 = wsLSMC.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLSMC.Range("B2").Value
    Dim JumpVol As Double: JumpVol = wsLSMC.Range("B3").Value
    Dim KOUpper As Double: KOUpper = wsLSMC.Range("B6").Value
    Dim KOLower As Double: KOLower = wsLSMC.Range("B7").Value
    Dim S1_0 As Double: S1_0 = wsLSMC.Range("B8").Value
    Dim S2_0 As Double: S2_0 = wsLSMC.Range("B9").Value
    Dim vol1 As Double: vol1 = wsLSMC.Range("B10").Value
    Dim vol2 As Double: vol2 = wsLSMC.Range("B11").Value
    Dim corr As Double: corr = wsLSMC.Range("B12").Value
    Dim T_total As Double: T_total = wsLSMC.Range("B14").Value
    Dim r_US As Double: r_US = wsEnc.Range("B4").Value
    Dim r_KRW As Double: r_KRW = wsEnc.Range("B5").Value
    Dim StressWTI As Double: StressWTI = wsEnc.Range("B18").Value
    Dim StressKRW As Double: StressKRW = wsEnc.Range("B19").Value

    Dim Drift1Q As Double: Drift1Q = r_US
    Dim Drift2Q As Double: Drift2Q = r_KRW - r_US

    Dim Steps As Long: Steps = CLng(T_total * STEPS_PER_YEAR)
    Dim n_sens As Long: n_sens = 20000

    wsJR.Range("A1:D1").Value = Array("Lambda_Q_Factor", "Lambda_Q", "Base_Premium", "Stress_Premium")

    Dim factors() As Variant: factors = Array(1#, 1.25, 1.5, 2#)
    Dim fIdx As Long
    For fIdx = LBound(factors) To UBound(factors)
        Dim factor As Double: factor = factors(fIdx)
        Dim LambdaQ As Double: LambdaQ = Lambda0 * factor

        Dim basePrem As Double
        basePrem = Calc_LSMC_Price(LambdaQ, JumpMean, JumpVol, Drift1Q, Drift2Q, KOUpper, KOLower, _
                                    S1_0, S2_0, vol1, vol2, corr, Steps, T_total, n_sens, S1_0)
        Dim stressPrem As Double
        stressPrem = Calc_LSMC_Price(LambdaQ, JumpMean, JumpVol, Drift1Q, Drift2Q, KOUpper, KOLower, _
                                      StressWTI, StressKRW, vol1, vol2, corr, Steps, T_total, n_sens, S1_0)

        wsJR.Cells(fIdx + 2, 1).Value = factor
        wsJR.Cells(fIdx + 2, 2).Value = LambdaQ
        wsJR.Cells(fIdx + 2, 3).Value = basePrem
        wsJR.Cells(fIdx + 2, 4).Value = stressPrem
    Next fIdx

    wsJR.Range("A7").Value = "NOTE: re-pricing sensitivity sweep, NOT a Girsanov jump-measure change"
    wsJR.Range("A8").Value = "(that needs an Esscher transform on intensity AND jump-size dist to stay"
    wsJR.Range("A9").Value = "risk-neutral -- out of scope here). See MODEL_SPEC ��17.8 / ��19/��20."

    wsJR.Columns("A:D").AutoFit

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Jump Risk Premium sensitivity complete (4 Lambda^Q factors)." & vbCrLf & _
           "See 'JumpPremiumSens' sheet. Reminder: approximate re-pricing sweep," & vbCrLf & _
           "not a rigorous measure change -- disclose accordingly.", vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_JumpRiskPremium_Sensitivity failed: " & Err.Description, vbCritical
End Sub

' =============================================================================
' Orchestrator ��� runs all 4 items in sequence, mirroring Run_ITM_Diagnostic's
' chain pattern above (now in this same merged file). Each item is also
' independently callable.
' =============================================================================
Public Sub Run_Paper_Integrity_Suite()
    Call Run_ITM_Diagnostic
    Call Run_PQ_Drift_Bias_Diagnostic
    Call Run_Sharpe_Sweep_Diagnostic
    Call Run_Girsanov_Reweight_Diagnostic
    Call Run_Lambda_Robustness_Sweep
    Call Run_FD_vs_Regression_Delta_Check
    Call Run_DeltaFX_Ratio_Sweep
    Call Run_JumpRiskPremium_Sensitivity
    Call Run_Payoff_Parity_Check

    MsgBox "Paper Integrity Suite complete." & vbCrLf & _
           "Check sheets: ITM_Diagnostic, PQ_DriftBias, SharpeSweep, Girsanov, " & vbCrLf & _
           "LambdaRobustness, DeltaCheck, DeltaFXSweep, JumpPremiumSens, PayoffParityCheck.", vbInformation
End Sub

' =============================================================================
' Payoff-Parity Consistency Check
' -----------------------------------------------------------------------------
' Purpose: isolate the PURE delta-methodology effect by decomposing total P&L
' into (a) payoff-structure differences and (b) hedge cost differences.
'
' Background: the three engines collect the same LSMC American premium but
' compute payoff differently:
'   American  -> max(S_WTI_ex - K, 0) * S_FX  [early exercise possible]
'   European  -> max(S_WTI_T  - K, 0) * S_FX  [terminal spot]
'   Asian     -> max(avg_WTI  - K, 0) * S_FX  [arithmetic average]
'
' Hedge cost (= opt_profit - tot_profit) depends only on delta strategy, not
' on which payoff formula was used.  This sub reports:
'   1. Hedge cost statistics for all three engines (full population).
'   2. Same stats for the "clean subset": non-KO paths where American did NOT
'      early-exercise -> payoff is numerically identical across American and
'      European (both = max(S_WTI_T - K, 0) * S_FX).  Asian remains different
'      in this subset because avg_WTI != S_WTI_T.
'   3. Quantifies the payoff structure bias (American opt_profit - European
'      opt_profit) so the paper can report it as a confound, not a delta effect.
'
' Column layout (TABLE_START = 30, data rows TABLE_START+1 .. TABLE_START+n):
'   American_Delta  col 12=opt_profit  col 13=tot_profit  col 15=KO  col 16=exercise
'   European_Delta  col 12=opt_profit  col 13=tot_profit  col 14=KO
'   Asian_Delta     col 13=opt_profit  col 14=tot_profit  col 15=KO
'
' Prerequisites: Run all three engines (Run_All_DeltaHedge_Engines) first.
' =============================================================================
Public Sub Run_Payoff_Parity_Check()
    Const TABLE_START As Long = 30

    Dim wsAm As Worksheet, wsEu As Worksheet, wsAs As Worksheet, wsOut As Worksheet
    On Error Resume Next
    Set wsAm = Sheets("American_Delta")
    Set wsEu = Sheets("European_Delta")
    Set wsAs = Sheets("Asian_Delta")
    On Error GoTo 0

    If wsAm Is Nothing Or wsEu Is Nothing Or wsAs Is Nothing Then
        MsgBox "One or more engine sheets missing.  Run all three engines first.", vbCritical
        Exit Sub
    End If

    Dim nA As Long, nE As Long, nX As Long
    nA = CLng(wsAm.Range("B1").Value)
    nE = CLng(wsEu.Range("B1").Value)
    nX = CLng(wsAs.Range("B1").Value)

    If nA < 1 Or nA <> nE Or nE <> nX Then
        MsgBox "SIM_RUNS mismatch or zero across sheets.  Re-run all three engines first.", vbCritical
        Exit Sub
    End If
    Dim n As Long: n = nA

    ' --- read arrays from sheets (bulk read for speed) ---
    Dim rAm As Variant, rEu As Variant, rAs As Variant
    rAm = wsAm.Range(wsAm.Cells(TABLE_START + 1, 12), wsAm.Cells(TABLE_START + n, 16)).Value
    rEu = wsEu.Range(wsEu.Cells(TABLE_START + 1, 12), wsEu.Cells(TABLE_START + n, 14)).Value
    rAs = wsAs.Range(wsAs.Cells(TABLE_START + 1, 13), wsAs.Cells(TABLE_START + n, 15)).Value
    ' rAm(i, 1)=opt_profit  (2)=tot_profit  (3)=jumploss  (4)=KO  (5)=exercise
    ' rEu(i, 1)=opt_profit  (2)=tot_profit  (3)=KO
    ' rAs(i, 1)=opt_profit  (2)=tot_profit  (3)=KO

    Dim i As Long
    ' accumulators -- full population
    Dim sumHcA As Double, sumHcE As Double, sumHcX As Double
    Dim sumSqA As Double, sumSqE As Double, sumSqX As Double
    ' clean subset (non-KO, American non-exercised)
    Dim sumHcA_c As Double, sumHcE_c As Double
    Dim sumSqA_c As Double, sumSqE_c As Double
    Dim nClean As Long
    ' payoff-structure bias
    Dim sumBiasAE As Double   ' American opt_profit - European opt_profit
    Dim sumBiasAX As Double   ' American opt_profit - Asian opt_profit

    Dim hcA As Double, hcE As Double, hcX As Double
    Dim koA As Boolean, koE As Boolean, koX As Boolean, exA As Boolean

    For i = 1 To n
        hcA = CDbl(rAm(i, 1)) - CDbl(rAm(i, 2))   ' opt - tot
        hcE = CDbl(rEu(i, 1)) - CDbl(rEu(i, 2))
        hcX = CDbl(rAs(i, 1)) - CDbl(rAs(i, 2))

        koA = (CStr(rAm(i, 4)) = "Y")
        koE = (CStr(rEu(i, 3)) = "Y")
        koX = (CStr(rAs(i, 3)) = "Y")
        exA = (CStr(rAm(i, 5)) <> "N")

        sumHcA = sumHcA + hcA
        sumHcE = sumHcE + hcE
        sumHcX = sumHcX + hcX
        sumSqA = sumSqA + hcA ^ 2
        sumSqE = sumSqE + hcE ^ 2
        sumSqX = sumSqX + hcX ^ 2

        ' payoff-structure bias (for all paths, payoff difference = opt_profit difference)
        sumBiasAE = sumBiasAE + (CDbl(rAm(i, 1)) - CDbl(rEu(i, 1)))
        sumBiasAX = sumBiasAX + (CDbl(rAm(i, 1)) - CDbl(rAs(i, 1)))

        ' clean subset: non-KO on all three engines AND no American early exercise
        If Not koA And Not koE And Not koX And Not exA Then
            nClean = nClean + 1
            sumHcA_c = sumHcA_c + hcA
            sumHcE_c = sumHcE_c + hcE
            sumSqA_c = sumSqA_c + hcA ^ 2
            sumSqE_c = sumSqE_c + hcE ^ 2
        End If
    Next i

    Dim meanA As Double, meanE As Double, meanX As Double
    Dim sdA As Double, sdE As Double, sdX As Double
    meanA = sumHcA / n:  sdA = Sqr(sumSqA / n - meanA ^ 2)
    meanE = sumHcE / n:  sdE = Sqr(sumSqE / n - meanE ^ 2)
    meanX = sumHcX / n:  sdX = Sqr(sumSqX / n - meanX ^ 2)

    Dim meanA_c As Double, meanE_c As Double
    Dim sdA_c As Double, sdE_c As Double
    If nClean > 1 Then
        meanA_c = sumHcA_c / nClean: sdA_c = Sqr(sumSqA_c / nClean - meanA_c ^ 2)
        meanE_c = sumHcE_c / nClean: sdE_c = Sqr(sumSqE_c / nClean - meanE_c ^ 2)
    End If

    ' --- output sheet ---
    On Error Resume Next
    Set wsOut = Sheets("PayoffParityCheck")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = "PayoffParityCheck"
        Set wsOut = Sheets("PayoffParityCheck")
    End If
    wsOut.Cells.Clear

    Dim r As Long: r = 1
    wsOut.Cells(r, 1) = "Payoff-Parity Consistency Check"
    wsOut.Cells(r, 1).Font.Bold = True
    r = r + 1
    wsOut.Cells(r, 1) = "N (simulations)": wsOut.Cells(r, 2) = n
    r = r + 2

    ' Section 1: Hedge Cost (full population)
    wsOut.Cells(r, 1) = "1. Hedge Cost Distribution ��� Full Population (n = " & n & ")"
    wsOut.Cells(r, 1).Font.Bold = True
    r = r + 1
    wsOut.Cells(r, 2) = "American (LSMC)"
    wsOut.Cells(r, 3) = "European (Black-76)"
    wsOut.Cells(r, 4) = "Asian (TW)"
    wsOut.Cells(r, 5) = "EU - AM"
    wsOut.Cells(r, 6) = "AS - AM"
    r = r + 1
    wsOut.Cells(r, 1) = "Mean (KRW)":       wsOut.Cells(r, 2) = meanA: wsOut.Cells(r, 3) = meanE: wsOut.Cells(r, 4) = meanX
    wsOut.Cells(r, 5) = meanE - meanA:      wsOut.Cells(r, 6) = meanX - meanA
    r = r + 1
    wsOut.Cells(r, 1) = "StdDev (KRW)":     wsOut.Cells(r, 2) = sdA:   wsOut.Cells(r, 3) = sdE:   wsOut.Cells(r, 4) = sdX
    r = r + 1
    wsOut.Cells(r, 1) = "Note"
    wsOut.Cells(r, 2) = "Hedge cost = opt_profit - tot_profit.  Depends only on delta strategy."
    r = r + 2

    ' Section 2: Hedge Cost (clean subset)
    wsOut.Cells(r, 1) = "2. Hedge Cost ��� Clean Subset: non-KO, non-early-exercise (n = " & nClean & ")"
    wsOut.Cells(r, 1).Font.Bold = True
    wsOut.Cells(r, 7) = Format(nClean / n, "0.0%") & " of total"
    r = r + 1
    wsOut.Cells(r, 2) = "American (LSMC)"
    wsOut.Cells(r, 3) = "European (Black-76)"
    wsOut.Cells(r, 5) = "EU - AM"
    r = r + 1
    wsOut.Cells(r, 1) = "Mean (KRW)":   wsOut.Cells(r, 2) = meanA_c: wsOut.Cells(r, 3) = meanE_c
    wsOut.Cells(r, 5) = meanE_c - meanA_c
    r = r + 1
    wsOut.Cells(r, 1) = "StdDev (KRW)": wsOut.Cells(r, 2) = sdA_c:   wsOut.Cells(r, 3) = sdE_c
    r = r + 1
    wsOut.Cells(r, 1) = "Note"
    wsOut.Cells(r, 2) = "In this subset American and European payoff are identical (max(S_WTI_T-K,0)*S_FX)."
    wsOut.Cells(r, 3) = "Hedge cost difference here = pure LSMC vs Black-76 delta effect."
    wsOut.Cells(r, 4) = "Asian excluded: avg_WTI != S_WTI_T even on identical paths."
    r = r + 2

    ' Section 3: Payoff-Structure Bias
    wsOut.Cells(r, 1) = "3. Payoff-Structure Bias (opt_profit difference on same paths)"
    wsOut.Cells(r, 1).Font.Bold = True
    r = r + 1
    wsOut.Cells(r, 2) = "Mean(AM opt) - Mean(EU opt) (KRW)": wsOut.Cells(r, 3) = sumBiasAE / n
    r = r + 1
    wsOut.Cells(r, 2) = "Mean(AM opt) - Mean(AS opt) (KRW)": wsOut.Cells(r, 3) = sumBiasAX / n
    r = r + 1
    wsOut.Cells(r, 2) = "If AM-EU bias >> 0: early exercise drove payoff difference, not delta."
    r = r + 1
    wsOut.Cells(r, 2) = "If AM-AS bias >> 0: avg_WTI < S_WTI_T (contango path, Asian cheaper)."
    r = r + 2

    ' Section 4: Interpretation guide for paper
    wsOut.Cells(r, 1) = "4. Paper Interpretation Guide"
    wsOut.Cells(r, 1).Font.Bold = True
    r = r + 1
    wsOut.Cells(r, 2) = "Primary delta comparison metric: hedge cost (Section 1 above)."
    r = r + 1
    wsOut.Cells(r, 2) = "Cleanest delta comparison: Section 2 (non-KO, non-exercise subset, AM vs EU)."
    r = r + 1
    wsOut.Cells(r, 2) = "Total profit comparison is confounded by Section 3 payoff-structure bias."
    r = r + 1
    wsOut.Cells(r, 2) = "Asian total profit comparison should be presented separately with explicit payoff-structure caveat."

    wsOut.Columns("A:G").AutoFit

    MsgBox "Payoff-Parity Check complete." & vbCrLf & _
           "Clean subset (non-KO, non-exercise): " & nClean & " / " & n & _
           " (" & Format(nClean / n, "0.0%") & ")" & vbCrLf & _
           "AM-EU hedge cost gap (full): " & Format((meanE - meanA) / 1000000, "0.00") & " M KRW" & vbCrLf & _
           "AM-EU hedge cost gap (clean): " & Format((meanE_c - meanA_c) / 1000000, "0.00") & " M KRW" & vbCrLf & _
           "See sheet: PayoffParityCheck", vbInformation, "Payoff-Parity Check"
End Sub
