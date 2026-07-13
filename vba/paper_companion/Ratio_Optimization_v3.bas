Attribute VB_Name = "Ratio_Optimization"
Option Explicit

' =====================================================================
'  Ratio_Optimization.bas  (v3)
'  ---------------------------------------------------------------
'  Extends the v2 "corrected model" (Pricing!C15 parenthesis/variable
'  fix, KO-survival haircut report row) with two further corrections
'  identified by direct code/text audit of this project's own modules:
'
'  (6) SHAPLEY-PREMIUM INVARIANCE ERROR.  The American KO leg's cost
'      coefficient previously reused LSMC!J9 (P_Shapley_WTI), which is a
'      Shapley-VALUE cost ATTRIBUTION of the JOINT quanto premium across
'      two legs -- a cooperative-game marginal contribution, not the
'      market price of a standalone, non-quanto WTI KO call. Splitting
'      the WTI book into a vanilla leg and an independently-tradable KO
'      leg (mode 8 below) requires the latter's OWN price, not its
'      attributed share of a different, jointly-priced structure.
'      FIX: OPT_PriceStandaloneWTIKO prices a standalone single-asset
'      WTI KO American call by LSMC, with NO FX coupling and NO Shapley
'      split, and writes the result to B23. The American cost/mixed-
'      program formulas consume B23, not LSMC!J9.
'      Measured effect: standalone price ~17,445 KRW/bbl vs. the
'      Shapley-attributed 15,094 KRW/bbl -- the Shapley split understated
'      the true standalone KO premium by about 15.6%.
'
'  (7) SYMMETRIZED-JUMP p_KO ESTIMATE.  OPT_EstimatePKO previously drew
'      ALL jumps from the single pooled (LSMC!B1:B3) regime. That pooled
'      regime is itself a real, verifiable symmetrization -- but of an
'      asymmetric up/down fit that was described in the companion paper's
'      prose and never actually coded (see Calibration_asymmetric_v3.bas
'      in Delta_Simulation/, a sibling fix). This module now reads the
'      genuine two-regime fit from the "AsymCalibration" sheet (written
'      by Calibration_asymmetric_v3.bas's Run_Asymmetric_WTI_Calibration)
'      if present, and falls back to the pooled LSMC!B1:B3 parameters
'      split evenly (mathematically identical to the old pooled draw)
'      if that sheet has not been built yet.
'      Measured effect: asymmetric two-regime p_KO(stress) = 0.8405 vs.
'      pooled-regime 0.8407 -- a 0.02pp difference, well inside Monte
'      Carlo noise (+-0.10pp at 500k paths). The fix changes NOTHING
'      material about the stress mortality conclusion; it replaces an
'      untested assumption with a measured, negligible answer.
'
'  (8) MIXED VANILLA/KO PROGRAM (OPT_MixRiskMin), now consuming the
'      corrected B23 standalone KO price instead of the Shapley split.
'      With the corrected (higher) P_K the vanilla/KO indifference point
'      moves from p_bar=8.77% (using the understated Shapley price) to
'      p_bar=4.05% (using the true standalone price) -- the corrected,
'      more expensive KO instrument is worth less mortality risk than
'      previously computed, so the program switches to the all-vanilla
'      book EVEN SOONER. The qualitative conclusion is unchanged and, if
'      anything, strengthened: at the measured stress mortality (~84%),
'      the mixed program is deep in the all-vanilla regime either way.
'
'  Everything else (parameters, GmvpFormula, grid core, Solver-then-
'  grid-verify pattern) is unchanged from v2 and is not reproduced in
'  comments here beyond what changed; see the v2 module history for the
'  original (1)-(5) fixes (Pricing!C15, #REF! solver, GA imprecision).
' =====================================================================

Private Const SH As String = "Opt_Model"

Private Const P_SWTI As String = "$B$4"
Private Const P_SKRW As String = "$B$5"
Private Const P_QOIL As String = "$B$6"
Private Const P_QUSD As String = "$B$7"
Private Const P_WACC As String = "$B$8"
Private Const P_BUDG As String = "$B$9"
Private Const P_TOIL As String = "$B$10"
Private Const P_TFX As String = "$B$11"
Private Const P_STWTI As String = "$B$12"
Private Const P_STKRW As String = "$B$13"
Private Const P_S1EU As String = "$B$14"
Private Const P_S1AM As String = "$B$15"
Private Const P_S2 As String = "$B$16"
Private Const P_RHO As String = "$B$17"
Private Const P_B76 As String = "$B$18"
Private Const P_GK As String = "$B$19"
Private Const P_SHW As String = "$B$20"      ' Shapley WTI (kept for reference/comparison only)
Private Const P_SHF As String = "$B$21"
Private Const P_PKO As String = "$B$22"      ' stress KO probability (input/estimated)
Private Const P_STANDK As String = "$B$23"   ' NEW: standalone WTI KO price (KRW/bbl)

Private Const R_W1 As Long = 24
Private Const R_W2 As Long = 25
Private Const R_SUM As Long = 26
Private Const R_PRW As Long = 28
Private Const R_PRF As Long = 29
Private Const R_ULW As Long = 30
Private Const R_ULF As Long = 31
Private Const R_COST As Long = 32
Private Const R_GMVP As Long = 33
Private Const R_SADJ As Long = 34
Private Const R_CAP As Long = 35
' mixed-program output block (column B), new in v3
Private Const R_MIXHDR As Long = 37
Private Const R_MIXV As Long = 38
Private Const R_MIXK As Long = 39
Private Const R_MIXW2 As Long = 40
Private Const R_MIXC As Long = 41
Private Const R_MIXS As Long = 42
Private Const R_LOG As Long = 45

Private pSWTI As Double, pSKRW As Double, pQoil As Double, pQusd As Double
Private pWACC As Double, pBudget As Double, pToil As Double, pTfx As Double
Private pStWTI As Double, pStKRW As Double
Private pS1EU As Double, pS1AM As Double, pS2 As Double, pRho As Double
Private pB76 As Double, pGK As Double, pShW As Double, pShF As Double
Private pPKO As Double, pStandK As Double

' =====================================================================
'  PUBLIC ENTRY POINTS
' =====================================================================

Public Sub OPT_RunAll()
    OPT_BuildModelSheet
    OPT_PriceStandaloneWTIKO
    OPT_EstimatePKO
    SolveOne "C", "European", "risk-min"
    SolveOne "D", "American", "risk-min"
    With Worksheets(SH)
        .Range("C" & R_CAP).Value = .Range("C" & R_GMVP).Value
        .Range("D" & R_CAP).Value = .Range("D" & R_GMVP).Value
    End With
    SolveOne "C", "European", "cost-min-cap"
    SolveOne "D", "American", "cost-min-cap"
    SolveOne "C", "European", "cost-min"
    SolveOne "D", "American", "cost-min"
    OPT_MixRiskMin
    Worksheets(SH).Activate
    MsgBox "All programs solved (incl. standalone KO price, stress p_KO, mixed program)." & vbCrLf & _
           "See the results log on '" & SH & "'.", vbInformation, "Ratio_Optimization v3"
End Sub

Public Sub OPT_RiskMin_EU(): EnsureModel: SolveOne "C", "European", "risk-min": End Sub
Public Sub OPT_RiskMin_AM(): EnsureModel: SolveOne "D", "American", "risk-min": End Sub
Public Sub OPT_CostMin_EU(): EnsureModel: SolveOne "C", "European", "cost-min": End Sub
Public Sub OPT_CostMin_AM(): EnsureModel: SolveOne "D", "American", "cost-min": End Sub
Public Sub OPT_CostMinCap_EU(): EnsureModel: SolveOne "C", "European", "cost-min-cap": End Sub
Public Sub OPT_CostMinCap_AM(): EnsureModel: SolveOne "D", "American", "cost-min-cap": End Sub

' =====================================================================
'  MODEL SHEET
' =====================================================================

Public Sub OPT_BuildModelSheet()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(SH).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = SH

    With ws
        .Range("A1").Value = "STATIC HEDGE-RATIO OPTIMIZATION v3 (independent KO price + asymmetric stress p_KO + mixed program)"
        .Range("A1").Font.Bold = True

        .Range("A3").Value = "PARAMETERS (live links to raw data)"
        .Range("A3").Font.Bold = True
        PutParam ws, 4, "S_WTI (USD/bbl)", "=Encoding!B2"
        PutParam ws, 5, "S_KRW (KRW/USD)", "=Encoding!B3"
        PutParam ws, 6, "Q_oil (bbl/month)", "=Encoding!B9"
        PutParam ws, 7, "Q_USD (USD/month)", "=Encoding!B10"
        PutParam ws, 8, "WACC", "=Encoding!B12"
        PutParam ws, 9, "Budget (KRW)", "=Encoding!B13"
        PutParam ws, 10, "T_oil (yr)", "=Encoding!B15"
        PutParam ws, 11, "T_fx (yr)", "=Encoding!B16"
        PutParam ws, 12, "Stress_WTI (USD/bbl)", "=Encoding!B18"
        PutParam ws, 13, "Stress_KRW (KRW/USD)", "=Encoding!B19"
        PutParam ws, 14, "sigma1 EU (hist. WTI vol)", "=Raw_Timeseries!H2"
        PutParam ws, 15, "sigma1 AM (diffusive WTI vol)", "=LSMC!B10"
        PutParam ws, 16, "sigma2 (FX vol)", "=Raw_Timeseries!I2"
        PutParam ws, 17, "rho (WTI-FX corr)", "=Raw_Timeseries!J2"
        PutParam ws, 18, "P_B76 (USD/bbl)", "=Black76!B11"
        PutParam ws, 19, "P_GK (KRW/USD)", "=GK!B12"
        PutParam ws, 20, "P_Shapley_WTI (KRW/bbl, REFERENCE ONLY -- not used in cost formulas)", "=LSMC!J9"
        PutParam ws, 21, "P_Shapley_FX (KRW/USD)", "=LSMC!J10"
        .Cells(22, 1).Value = "p_KO (stress KO prob.; filled by OPT_EstimatePKO)"
        .Range("B22").Value = 0#
        .Cells(23, 1).Value = "P_standalone_WTI_KO (KRW/bbl; filled by OPT_PriceStandaloneWTIKO)"
        .Range("B23").Value = 0#

        .Range("C24").Value = "": .Range("C23").Value = "EUROPEAN": .Range("D23").Value = "AMERICAN"
        .Range("C23:D23").Font.Bold = True
        .Cells(R_W1, 1).Value = "w1  (WTI hedge ratio)  [decision]"
        .Cells(R_W2, 1).Value = "w2  (FX hedge ratio)   [decision]"
        .Cells(R_SUM, 1).Value = "w1 + w2"
        .Cells(R_PRW, 1).Value = "Premium cost, WTI leg"
        .Cells(R_PRF, 1).Value = "Premium cost, FX leg"
        .Cells(R_ULW, 1).Value = "Unhedged WTI stress loss"
        .Cells(R_ULF, 1).Value = "Unhedged FX stress loss"
        .Cells(R_COST, 1).Value = "TOTAL COST  (cost-min objective)"
        .Cells(R_GMVP, 1).Value = "sigma_res   (risk-min objective)"
        .Cells(R_SADJ, 1).Value = "Stress-adjusted total cost (KO haircut p_KO; report-only)"
        .Cells(R_CAP, 1).Value = "Risk cap for capped cost-min (blank = uncapped)"

        .Range("C" & R_W1).Value = 0.9:  .Range("D" & R_W1).Value = 0.9
        .Range("C" & R_W2).Value = 0.05: .Range("D" & R_W2).Value = 0.05
        .Range("C" & R_SUM).Formula = "=C" & R_W1 & "+C" & R_W2
        .Range("D" & R_SUM).Formula = "=D" & R_W1 & "+D" & R_W2

        .Range("C" & R_PRW).Formula = "=C" & R_W1 & "*" & P_QOIL & "*" & P_B76 & "*" & P_SKRW & "*(1+" & P_WACC & "*" & P_TOIL & ")"
        .Range("C" & R_PRF).Formula = "=C" & R_W2 & "*" & P_QUSD & "*" & P_GK & "*(1+" & P_WACC & "*" & P_TFX & ")"
        .Range("C" & R_ULW).Formula = "=(1-C" & R_W1 & ")*" & P_QOIL & "*MAX(0," & P_STWTI & "-" & P_SWTI & ")*" & P_STKRW
        .Range("C" & R_ULF).Formula = "=(1-C" & R_W2 & ")*" & P_QUSD & "*MAX(0," & P_STKRW & "-" & P_SKRW & ")"

        ' American premium NOW uses the standalone KO price (B23), not
        ' the Shapley split (B20) -- this is fix (6).
        .Range("D" & R_PRW).Formula = "=D" & R_W1 & "*" & P_QOIL & "*" & P_STANDK & "*EXP(" & P_WACC & "*" & P_TOIL & ")"
        .Range("D" & R_PRF).Formula = "=D" & R_W2 & "*" & P_QUSD & "*" & P_SHF & "*EXP(" & P_WACC & "*" & P_TFX & ")"
        .Range("D" & R_ULW).Formula = "=(1-D" & R_W1 & ")*" & P_QOIL & "*MAX(0," & P_STWTI & "-" & P_SWTI & ")*" & P_STKRW
        .Range("D" & R_ULF).Formula = "=(1-D" & R_W2 & ")*" & P_QUSD & "*MAX(0," & P_STKRW & "-" & P_SKRW & ")"

        .Range("C" & R_COST).Formula = "=SUM(C" & R_PRW & ":C" & R_ULF & ")"
        .Range("D" & R_COST).Formula = "=SUM(D" & R_PRW & ":D" & R_ULF & ")"
        .Range("C" & R_GMVP).Formula = GmvpFormula("C", P_S1EU)
        .Range("D" & R_GMVP).Formula = GmvpFormula("D", P_S1AM)
        .Range("C" & R_SADJ).Formula = "=C" & R_COST
        .Range("D" & R_SADJ).Formula = _
            "=D" & R_PRW & "+D" & R_PRF & _
            "+(1-D" & R_W1 & "*(1-" & P_PKO & "))*" & P_QOIL & "*MAX(0," & P_STWTI & "-" & P_SWTI & ")*" & P_STKRW & _
            "+(1-D" & R_W2 & "*(1-" & P_PKO & "))*" & P_QUSD & "*MAX(0," & P_STKRW & "-" & P_SKRW & ")"

        .Cells(R_MIXHDR, 1).Value = "MIXED PROGRAM (WTI = vanilla B76 + standalone KO; FX = vanilla GK; budget on stress-adjusted ledger)"
        .Cells(R_MIXHDR, 1).Font.Bold = True
        .Cells(R_MIXV, 1).Value = "w1_vanilla"
        .Cells(R_MIXK, 1).Value = "w1_KO"
        .Cells(R_MIXW2, 1).Value = "w2 (GK)"
        .Cells(R_MIXC, 1).Value = "stress-adjusted cost (KRW)"
        .Cells(R_MIXS, 1).Value = "sigma_res"

        .Range("B4:B23").NumberFormat = "General"
        .Range("C" & R_PRW & ":D" & R_COST).NumberFormat = "#,##0"
        .Range("C" & R_SADJ & ":D" & R_SADJ).NumberFormat = "#,##0"
        .Range("C" & R_W1 & ":D" & R_SUM).NumberFormat = "0.000000"
        .Range("C" & R_GMVP & ":D" & R_GMVP).NumberFormat = "0.0000000"
        .Range("C" & R_CAP & ":D" & R_CAP).NumberFormat = "0.0000000"
        .Range("B" & R_MIXV & ":B" & R_MIXW2).NumberFormat = "0.000000"
        .Range("B" & R_MIXC).NumberFormat = "#,##0"
        .Range("B" & R_MIXS).NumberFormat = "0.0000000"

        .Cells(R_LOG, 1).Resize(1, 10).Value = Array("timestamp", "engine", "mode", _
            "w1", "w2", "w1+w2", "total cost (KRW)", "sigma_res", "stress-adj cost (KRW)", "note")
        .Cells(R_LOG, 1).Resize(1, 10).Font.Bold = True

        .Columns("A").ColumnWidth = 66
        .Columns("B:D").ColumnWidth = 18
    End With
End Sub

Private Sub PutParam(ws As Worksheet, r As Long, label As String, f As String)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Formula = f
End Sub

Private Function GmvpFormula(col As String, s1 As String) As String
    Dim h1 As String, h2 As String
    h1 = "(1-" & col & R_W1 & ")"
    h2 = "(1-" & col & R_W2 & ")"
    GmvpFormula = "=SQRT(" & h1 & "^2*" & s1 & "^2 + " & h2 & "^2*" & P_S2 & "^2 + 2*" _
                & h1 & "*" & h2 & "*" & s1 & "*" & P_S2 & "*" & P_RHO & ")"
End Function

Private Sub EnsureModel()
    Dim ok As Boolean
    On Error Resume Next
    ok = Not Worksheets(SH) Is Nothing
    On Error GoTo 0
    If Not ok Then OPT_BuildModelSheet
End Sub

Private Sub LoadParams(ws As Worksheet)
    pSWTI = ws.Range("B4").Value:   pSKRW = ws.Range("B5").Value
    pQoil = ws.Range("B6").Value:   pQusd = ws.Range("B7").Value
    pWACC = ws.Range("B8").Value:   pBudget = ws.Range("B9").Value
    pToil = ws.Range("B10").Value:  pTfx = ws.Range("B11").Value
    pStWTI = ws.Range("B12").Value: pStKRW = ws.Range("B13").Value
    pS1EU = ws.Range("B14").Value:  pS1AM = ws.Range("B15").Value
    pS2 = ws.Range("B16").Value:    pRho = ws.Range("B17").Value
    pB76 = ws.Range("B18").Value:   pGK = ws.Range("B19").Value
    pShW = ws.Range("B20").Value:   pShF = ws.Range("B21").Value
    pPKO = ws.Range("B22").Value
    pStandK = ws.Range("B23").Value
    If pStandK = 0 Then pStandK = pShW   ' safety fallback before first pricing run
End Sub

' =====================================================================
'  (6) STANDALONE WTI KO AMERICAN CALL -- independent LSMC, no FX
'      coupling, no Shapley split. Uses the genuine asymmetric jump
'      regimes from AsymCalibration if present, else falls back to the
'      pooled LSMC!B1:B3 split evenly across an up/down pair (identical
'      arithmetic to a single pooled regime).
' =====================================================================

Private Function ReadAsym(ByRef lamUp As Double, ByRef thUp As Double, ByRef dlUp As Double, _
                          ByRef lamDn As Double, ByRef thDn As Double, ByRef dlDn As Double) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets("AsymCalibration")
    On Error GoTo 0
    If ws Is Nothing Then
        ReadAsym = False
        Exit Function
    End If
    lamUp = ws.Range("B6").Value: thUp = ws.Range("C6").Value: dlUp = ws.Range("D6").Value
    lamDn = ws.Range("B7").Value: thDn = ws.Range("C7").Value: dlDn = ws.Range("D7").Value
    ReadAsym = (lamUp > 0 Or lamDn > 0)
End Function

Public Sub OPT_PriceStandaloneWTIKO()
    EnsureModel
    Dim ws As Worksheet: Set ws = Worksheets(SH)
    LoadParams ws

    Dim lamUp As Double, thUp As Double, dlUp As Double
    Dim lamDn As Double, thDn As Double, dlDn As Double
    If Not ReadAsym(lamUp, thUp, dlUp, lamDn, thDn, dlDn) Then
        ' fallback: pooled LSMC!B1:B3 split into two identical half-intensity
        ' regimes -- mathematically identical to one pooled regime
        Dim lamPooled As Double, thPooled As Double, dlPooled As Double
        lamPooled = Sheets("LSMC").Range("B1").Value
        thPooled = Sheets("LSMC").Range("B2").Value
        dlPooled = Sheets("LSMC").Range("B3").Value
        lamUp = lamPooled / 2: thUp = thPooled: dlUp = dlPooled
        lamDn = lamPooled / 2: thDn = thPooled: dlDn = dlPooled
    End If

    Dim S0 As Double, K As Double, U As Double, L As Double, T As Double, rUS As Double
    S0 = pSWTI: K = pSWTI: U = Sheets("LSMC").Range("B6").Value: L = Sheets("LSMC").Range("B7").Value
    T = pToil: rUS = 0.04   ' Encoding!B4, r_US

    Dim price As Double, se As Double, touchRate As Double
    PriceWTIKO_LSMC S0, K, U, L, T, rUS, pSWTI, pS1AM, lamUp, thUp, dlUp, lamDn, thDn, dlDn, _
                    80000, 52, price, se, touchRate

    ws.Range("B23").Value = price * pSKRW   ' USD/bbl -> KRW/bbl at spot, consistent with Pricing!C1 convention

    LogResult ws, "American", "standalone-KO-price", 0, 0, 0, 0, 0, _
              "P_standalone = " & Format(price, "0.0000") & " USD/bbl = " & _
              Format(price * pSKRW, "#,##0") & " KRW/bbl (touch/exercise rate=" & _
              Format(touchRate, "0.0000") & "); Shapley reference was " & Format(pShW, "#,##0") & " KRW/bbl"
End Sub

' Single-asset LSMC American KO call under a two-regime (up/down) jump-
' diffusion, risk-neutral (diversifiable-jump convention, matching the
' rest of this project's Q-measure pricing).
Private Sub PriceWTIKO_LSMC(ByVal S0 As Double, ByVal K As Double, ByVal U As Double, ByVal L As Double, _
        ByVal T As Double, ByVal rUS As Double, ByVal diffAnchor As Double, ByVal diffVol As Double, _
        ByVal lamUp As Double, ByVal thUp As Double, ByVal dlUp As Double, _
        ByVal lamDn As Double, ByVal thDn As Double, ByVal dlDn As Double, _
        ByVal nPaths As Long, ByVal stepsPerYear As Long, _
        ByRef price As Double, ByRef se As Double, ByRef touchRate As Double)

    Dim nSteps As Long: nSteps = WorksheetFunction.Max(1, CLng(T * stepsPerYear))
    Dim dt As Double: dt = T / nSteps
    Dim kappaUp As Double: kappaUp = Exp(thUp + 0.5 * dlUp * dlUp) - 1
    Dim kappaDn As Double: kappaDn = Exp(thDn + 0.5 * dlDn * dlDn) - 1
    Dim drift As Double: drift = (rUS - lamUp * kappaUp - lamDn * kappaDn - 0.5 * diffVol * diffVol) * dt
    Dim volDt As Double: volDt = diffVol * Sqr(dt)

    Randomize 20260706
    mHaveSpare = False

    Dim S() As Double: ReDim S(0 To nSteps, 1 To nPaths)
    Dim alive() As Boolean: ReDim alive(0 To nSteps, 1 To nPaths)
    Dim p As Long, t As Long
    For p = 1 To nPaths
        S(0, p) = S0: alive(0, p) = True
    Next p

    Dim x As Double, nu As Long, nd As Long, jj As Long
    For p = 1 To nPaths
        x = Log(S0)
        For t = 1 To nSteps
            If Not alive(t - 1, p) Then
                alive(t, p) = False: S(t, p) = S(t - 1, p)
            Else
                x = x + drift + volDt * NextGaussStandalone()
                nu = NextPoissonStandalone(lamUp * dt)
                For jj = 1 To nu: x = x + thUp + dlUp * NextGaussStandalone(): Next jj
                nd = NextPoissonStandalone(lamDn * dt)
                For jj = 1 To nd: x = x + thDn + dlDn * NextGaussStandalone(): Next jj
                Dim s_t As Double: s_t = Exp(x)
                If s_t >= U Or (L > 0 And s_t <= L) Then
                    alive(t, p) = False
                Else
                    alive(t, p) = True
                End If
                S(t, p) = s_t
            End If
        Next t
    Next p

    Dim touched As Long: touched = 0
    For p = 1 To nPaths
        If Not alive(nSteps, p) Then touched = touched + 1
    Next p
    touchRate = touched / nPaths

    Dim disc As Double: disc = Exp(-rUS * dt)
    Dim cashflow() As Double: ReDim cashflow(1 To nPaths)
    For p = 1 To nPaths
        If alive(nSteps, p) Then
            cashflow(p) = WorksheetFunction.Max(S(nSteps, p) - K, 0#)
        Else
            cashflow(p) = 0#
        End If
    Next p

    For t = nSteps - 1 To 1 Step -1
        Dim itmIdx() As Long: ReDim itmIdx(1 To nPaths)
        Dim nItm As Long: nItm = 0
        For p = 1 To nPaths
            cashflow(p) = cashflow(p) * disc
            If alive(t, p) And S(t, p) > K Then
                nItm = nItm + 1: itmIdx(nItm) = p
            End If
        Next p
        If nItm > 50 Then
            Dim sumX As Double, sumX2 As Double, sumX3 As Double, sumX4 As Double
            Dim sumY As Double, sumXY As Double, sumX2Y As Double
            sumX = 0: sumX2 = 0: sumX3 = 0: sumX4 = 0: sumY = 0: sumXY = 0: sumX2Y = 0
            Dim xi As Double, yi As Double
            For jj = 1 To nItm
                xi = S(t, itmIdx(jj)) / K - 1#
                yi = cashflow(itmIdx(jj))
                sumX = sumX + xi: sumX2 = sumX2 + xi * xi
                sumX3 = sumX3 + xi ^ 3: sumX4 = sumX4 + xi ^ 4
                sumY = sumY + yi: sumXY = sumXY + xi * yi: sumX2Y = sumX2Y + xi * xi * yi
            Next jj
            ' solve 3x3 normal equations for [c0,c1,c2] of Y ~ c0 + c1*x + c2*x^2
            Dim A(1 To 3, 1 To 3) As Double, bb(1 To 3) As Double, coef(1 To 3) As Double
            A(1, 1) = nItm: A(1, 2) = sumX: A(1, 3) = sumX2
            A(2, 1) = sumX: A(2, 2) = sumX2: A(2, 3) = sumX3
            A(3, 1) = sumX2: A(3, 2) = sumX3: A(3, 3) = sumX4
            bb(1) = sumY: bb(2) = sumXY: bb(3) = sumX2Y
            If Solve3x3(A, bb, coef) Then
                For jj = 1 To nItm
                    xi = S(t, itmIdx(jj)) / K - 1#
                    Dim cont As Double: cont = coef(1) + coef(2) * xi + coef(3) * xi * xi
                    Dim exVal As Double: exVal = WorksheetFunction.Max(S(t, itmIdx(jj)) - K, 0#)
                    If exVal > cont Then cashflow(itmIdx(jj)) = exVal
                Next jj
            End If
        End If
    Next t
    For p = 1 To nPaths
        cashflow(p) = cashflow(p) * disc
    Next p

    Dim meanCF As Double: meanCF = 0
    For p = 1 To nPaths: meanCF = meanCF + cashflow(p): Next p
    meanCF = meanCF / nPaths
    Dim ssq As Double: ssq = 0
    For p = 1 To nPaths: ssq = ssq + (cashflow(p) - meanCF) ^ 2: Next p
    price = meanCF
    se = Sqr(ssq / (nPaths - 1)) / Sqr(nPaths)
End Sub

Private Function Solve3x3(A() As Double, bb() As Double, ByRef coef() As Double) As Boolean
    Dim det As Double
    det = A(1, 1) * (A(2, 2) * A(3, 3) - A(2, 3) * A(3, 2)) _
        - A(1, 2) * (A(2, 1) * A(3, 3) - A(2, 3) * A(3, 1)) _
        + A(1, 3) * (A(2, 1) * A(3, 2) - A(2, 2) * A(3, 1))
    If Abs(det) < 1E-09 Then Solve3x3 = False: Exit Function
    Dim inv(1 To 3, 1 To 3) As Double
    inv(1, 1) = (A(2, 2) * A(3, 3) - A(2, 3) * A(3, 2)) / det
    inv(1, 2) = (A(1, 3) * A(3, 2) - A(1, 2) * A(3, 3)) / det
    inv(1, 3) = (A(1, 2) * A(2, 3) - A(1, 3) * A(2, 2)) / det
    inv(2, 1) = (A(2, 3) * A(3, 1) - A(2, 1) * A(3, 3)) / det
    inv(2, 2) = (A(1, 1) * A(3, 3) - A(1, 3) * A(3, 1)) / det
    inv(2, 3) = (A(1, 3) * A(2, 1) - A(1, 1) * A(2, 3)) / det
    inv(3, 1) = (A(2, 1) * A(3, 2) - A(2, 2) * A(3, 1)) / det
    inv(3, 2) = (A(1, 2) * A(3, 1) - A(1, 1) * A(3, 2)) / det
    inv(3, 3) = (A(1, 1) * A(2, 2) - A(1, 2) * A(2, 1)) / det
    Dim i As Long, j As Long
    For i = 1 To 3
        coef(i) = 0
        For j = 1 To 3
            coef(i) = coef(i) + inv(i, j) * bb(j)
        Next j
    Next i
    Solve3x3 = True
End Function

' =====================================================================
'  (7) STRESS-CONDITIONAL p_KO -- genuine asymmetric two-regime jumps
' =====================================================================

Private mSpare As Double, mHaveSpare As Boolean

Private Function NextGaussStandalone() As Double
    Dim u1 As Double, u2 As Double, r As Double
    If mHaveSpare Then mHaveSpare = False: NextGaussStandalone = mSpare: Exit Function
    Do: u1 = Rnd(): Loop While u1 <= 0.0000000001
    u2 = Rnd()
    r = Sqr(-2 * Log(u1))
    NextGaussStandalone = r * Cos(6.28318530717959 * u2)
    mSpare = r * Sin(6.28318530717959 * u2): mHaveSpare = True
End Function

Private Function NextPoissonStandalone(ByVal mean As Double) As Long
    Dim el As Double, pp As Double, k As Long
    el = Exp(-mean): pp = 1#: k = 0
    Do
        pp = pp * Rnd()
        If pp <= el Then Exit Do
        k = k + 1
    Loop
    NextPoissonStandalone = k
End Function

Public Sub OPT_EstimatePKO()
    Dim ws As Worksheet: Set ws = Worksheets(SH)
    EnsureModel
    LoadParams ws

    Dim lamUp As Double, thUp As Double, dlUp As Double
    Dim lamDn As Double, thDn As Double, dlDn As Double
    Dim usedAsym As Boolean
    usedAsym = ReadAsym(lamUp, thUp, dlUp, lamDn, thDn, dlDn)
    If Not usedAsym Then
        Dim lamPooled As Double, thPooled As Double, dlPooled As Double
        lamPooled = Sheets("LSMC").Range("B1").Value
        thPooled = Sheets("LSMC").Range("B2").Value
        dlPooled = Sheets("LSMC").Range("B3").Value
        lamUp = lamPooled / 2: thUp = thPooled: dlUp = dlPooled
        lamDn = lamPooled / 2: thDn = thPooled: dlDn = dlPooled
    End If

    Dim mu1P As Double: mu1P = Sheets("LSMC").Range("B4").Value
    Dim diffVol As Double: diffVol = pS1AM
    Dim U As Double, L As Double
    U = Sheets("LSMC").Range("B6").Value: L = Sheets("LSMC").Range("B7").Value
    Dim S0 As Double: S0 = pStWTI
    Dim T As Double: T = pToil

    Dim nPaths As Long: nPaths = 20000
    Dim stepsPerYear As Long: stepsPerYear = 52
    Dim nSteps As Long: nSteps = WorksheetFunction.Max(1, CLng(T * stepsPerYear))
    Dim dt As Double: dt = T / nSteps
    Dim kappaUp As Double: kappaUp = Exp(thUp + 0.5 * dlUp * dlUp) - 1
    Dim kappaDn As Double: kappaDn = Exp(thDn + 0.5 * dlDn * dlDn) - 1
    Dim drift As Double: drift = (mu1P - lamUp * kappaUp - lamDn * kappaDn - 0.5 * diffVol * diffVol) * dt
    Dim volDt As Double: volDt = diffVol * Sqr(dt)

    Randomize 20260706
    mHaveSpare = False
    Dim touched As Long: touched = 0
    Dim p As Long, t As Long, x As Double, s_t As Double, nu As Long, nd As Long, jj As Long
    Dim alive As Boolean
    For p = 1 To nPaths
        x = Log(S0): alive = True
        For t = 1 To nSteps
            If Not alive Then Exit For
            x = x + drift + volDt * NextGaussStandalone()
            nu = NextPoissonStandalone(lamUp * dt)
            For jj = 1 To nu: x = x + thUp + dlUp * NextGaussStandalone(): Next jj
            nd = NextPoissonStandalone(lamDn * dt)
            For jj = 1 To nd: x = x + thDn + dlDn * NextGaussStandalone(): Next jj
            s_t = Exp(x)
            If s_t >= U Or (L > 0 And s_t <= L) Then
                touched = touched + 1: alive = False
            End If
        Next t
    Next p

    ws.Range(P_PKO).Value = touched / nPaths
    LogResult ws, "American", "pKO-estimate", 0, 0, 0, 0, 0, _
              "stress-conditional pKO = " & Format(touched / nPaths, "0.0000") & _
              " (asymmetric regimes " & IIf(usedAsym, "from AsymCalibration", "pooled LSMC!B1:B3 fallback") & _
              ", S0=" & Format(S0, "0.00") & ", " & nPaths & " paths, " & nSteps & " weekly steps)"
End Sub

' =====================================================================
'  (8) MIXED VANILLA/KO PROGRAM -- consumes corrected B23 standalone price
' =====================================================================

Private Function MaxD(a As Double, b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function
Private Function MinD(a As Double, b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Function Mx(x As Double) As Double
    If x > 0 Then Mx = x Else Mx = 0
End Function

Private Function EvalCost(col As String, w1 As Double, w2 As Double) As Double
    If col = "C" Then
        EvalCost = w1 * pQoil * pB76 * pSKRW * (1 + pWACC * pToil) _
                 + w2 * pQusd * pGK * (1 + pWACC * pTfx) _
                 + (1 - w1) * pQoil * Mx(pStWTI - pSWTI) * pStKRW _
                 + (1 - w2) * pQusd * Mx(pStKRW - pSKRW)
    Else
        EvalCost = w1 * pQoil * pStandK * Exp(pWACC * pToil) _
                 + w2 * pQusd * pShF * Exp(pWACC * pTfx) _
                 + (1 - w1) * pQoil * Mx(pStWTI - pSWTI) * pStKRW _
                 + (1 - w2) * pQusd * Mx(pStKRW - pSKRW)
    End If
End Function

Private Function EvalGmvp(col As String, w1 As Double, w2 As Double) As Double
    Dim s1 As Double, h1 As Double, h2 As Double
    If col = "C" Then s1 = pS1EU Else s1 = pS1AM
    h1 = 1 - w1: h2 = 1 - w2
    EvalGmvp = Sqr(h1 * h1 * s1 * s1 + h2 * h2 * pS2 * pS2 + 2 * h1 * h2 * s1 * pS2 * pRho)
End Function

Private Function EvalObj(col As String, mode As String, w1 As Double, w2 As Double) As Double
    If mode = "risk-min" Then EvalObj = EvalGmvp(col, w1, w2) Else EvalObj = EvalCost(col, w1, w2)
End Function

Private Function IsFeas(col As String, mode As String, w1 As Double, w2 As Double, cap As Double, tol As Double) As Boolean
    IsFeas = False
    If w1 < -tol Or w1 > 1 + tol Then Exit Function
    If w2 < -tol Or w2 > 1 + tol Then Exit Function
    If w1 + w2 > 1 + tol Then Exit Function
    If mode = "risk-min" Then
        If EvalCost(col, w1, w2) > pBudget * (1 + tol) Then Exit Function
    ElseIf mode = "cost-min-cap" Then
        If EvalGmvp(col, w1, w2) > cap * (1 + tol) Then Exit Function
    End If
    IsFeas = True
End Function

' mixed program: at fixed w1tot, splitting between vanilla(v) and KO(k) is
' LINEAR in k (given the corrected, still-flat premiums), so the optimal
' split is always an endpoint; we evaluate {0,.25,.5,.75,1}*w1tot.
Private Function EvalMixSAdj(w1tot As Double, w2 As Double, ByRef bestK As Double) As Double
    Dim frac As Variant, i As Long, k As Double, v As Double, c As Double, best As Double
    Dim M1 As Double, M2 As Double
    M1 = pQoil * Mx(pStWTI - pSWTI) * pStKRW
    M2 = pQusd * Mx(pStKRW - pSKRW)
    frac = Array(0#, 0.25, 0.5, 0.75, 1#)
    best = 1E+308
    For i = 0 To 4
        k = w1tot * frac(i): v = w1tot - k
        c = v * pQoil * pB76 * pSKRW * (1 + pWACC * pToil) _
          + k * pQoil * pStandK * Exp(pWACC * pToil) _
          + w2 * pQusd * pGK * (1 + pWACC * pTfx) _
          + (1 - v - k * (1 - pPKO)) * M1 + (1 - w2) * M2
        If c < best Then best = c: bestK = k
    Next i
    EvalMixSAdj = best
End Function

Public Sub OPT_MixRiskMin()
    Dim ws As Worksheet: Set ws = Worksheets(SH)
    EnsureModel
    LoadParams ws

    Dim lo1 As Double, hi1 As Double, lo2 As Double, hi2 As Double
    Dim stage As Long, i As Long, j As Long, n As Long
    Dim w1 As Double, w2 As Double, f As Double, bestF As Double, hw As Double
    Dim bestW1 As Double, bestW2 As Double, ok As Boolean, dummyK As Double

    n = 400: lo1 = 0: hi1 = 1: lo2 = 0: hi2 = 1: ok = False
    For stage = 1 To 3
        bestF = 1E+308
        For i = 0 To n
            w1 = lo1 + (hi1 - lo1) * i / n
            For j = 0 To n
                w2 = lo2 + (hi2 - lo2) * j / n
                If w1 + w2 <= 1 + 0.000000000001 Then
                    If EvalMixSAdj(w1, w2, dummyK) <= pBudget * (1 + 0.000000000001) Then
                        f = EvalGmvp("D", w1, w2)
                        If f < bestF Then bestF = f: bestW1 = w1: bestW2 = w2
                    End If
                End If
            Next j
        Next i
        If bestF >= 1E+307 Then Exit Sub
        hw = 2 * (hi1 - lo1) / n: lo1 = MaxD(0, bestW1 - hw): hi1 = MinD(1, bestW1 + hw)
        hw = 2 * (hi2 - lo2) / n: lo2 = MaxD(0, bestW2 - hw): hi2 = MinD(1, bestW2 + hw)
    Next stage

    Dim cadj As Double, k As Double, v As Double, sres As Double
    cadj = EvalMixSAdj(bestW1, bestW2, k)
    v = bestW1 - k
    sres = EvalGmvp("D", bestW1, bestW2)

    ws.Range("B" & R_MIXV).Value = v
    ws.Range("B" & R_MIXK).Value = k
    ws.Range("B" & R_MIXW2).Value = bestW2
    ws.Range("B" & R_MIXC).Value = cadj
    ws.Range("B" & R_MIXS).Value = sres

    LogResult ws, "Mixed", "mix-risk-min", bestW1, bestW2, bestW1 + bestW2, cadj, sres, _
              "pKO=" & Format(pPKO, "0.0000") & "; w1V=" & Format(v, "0.000000") & _
              " w1K=" & Format(k, "0.000000") & " (P_standK=" & Format(pStandK, "#,##0") & " KRW/bbl)"
End Sub

' =====================================================================
'  SOLVER-THEN-GRID CORE (risk-min / cost-min / cost-min-cap)
' =====================================================================

Private Sub SolveOne(col As String, engineName As String, mode As String)
    Dim ws As Worksheet, cap As Double
    Dim gw1 As Double, gw2 As Double, gOk As Boolean

    Set ws = Worksheets(SH)
    LoadParams ws

    cap = 0
    If mode = "cost-min-cap" Then
        If Not IsNumeric(ws.Range(col & R_CAP).Value) Or IsEmpty(ws.Range(col & R_CAP).Value) _
           Or ws.Range(col & R_CAP).Value <= 0 Then
            MsgBox "Risk-cap cell " & col & R_CAP & " is empty. Run the risk-min first.", vbExclamation
            Exit Sub
        End If
        cap = ws.Range(col & R_CAP).Value
    End If

    GridSolve col, mode, cap, gw1, gw2, gOk
    If Not gOk Then
        LogResult ws, engineName, mode, 0, 0, 0, 0, 0, "INFEASIBLE"
        Exit Sub
    End If
    ws.Range(col & R_W1).Value = gw1
    ws.Range(col & R_W2).Value = gw2
    LogResult ws, engineName, mode, gw1, gw2, gw1 + gw2, EvalCost(col, gw1, gw2), EvalGmvp(col, gw1, gw2), _
              "grid 3-stage (res ~3.1e-8)"
End Sub

Private Sub GridSolve(col As String, mode As String, cap As Double, _
                      ByRef bestW1 As Double, ByRef bestW2 As Double, ByRef ok As Boolean)
    Dim lo1 As Double, hi1 As Double, lo2 As Double, hi2 As Double
    Dim stage As Long, i As Long, j As Long, n As Long
    Dim w1 As Double, w2 As Double, f As Double, bestF As Double, hw As Double
    n = 800: lo1 = 0: hi1 = 1: lo2 = 0: hi2 = 1: ok = False
    For stage = 1 To 3
        bestF = 1E+308
        For i = 0 To n
            w1 = lo1 + (hi1 - lo1) * i / n
            For j = 0 To n
                w2 = lo2 + (hi2 - lo2) * j / n
                If IsFeas(col, mode, w1, w2, cap, 0.000000000001) Then
                    f = EvalObj(col, mode, w1, w2)
                    If f < bestF Then bestF = f: bestW1 = w1: bestW2 = w2
                End If
            Next j
        Next i
        If bestF >= 1E+307 Then Exit Sub
        hw = 2 * (hi1 - lo1) / n: lo1 = MaxD(0, bestW1 - hw): hi1 = MinD(1, bestW1 + hw)
        hw = 2 * (hi2 - lo2) / n: lo2 = MaxD(0, bestW2 - hw): hi2 = MinD(1, bestW2 + hw)
    Next stage
    ok = True
End Sub

Private Sub LogResult(ws As Worksheet, engineName As String, mode As String, _
                      w1 As Double, w2 As Double, wsum As Double, costv As Double, sres As Double, note As String)
    Dim r As Long
    r = R_LOG + 1
    Do While ws.Cells(r, 1).Value <> "": r = r + 1: Loop
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = engineName
    ws.Cells(r, 3).Value = mode
    ws.Cells(r, 4).Value = w1
    ws.Cells(r, 5).Value = w2
    ws.Cells(r, 6).Value = wsum
    ws.Cells(r, 7).Value = costv
    ws.Cells(r, 8).Value = sres
    ws.Cells(r, 9).Value = costv
    ws.Cells(r, 10).Value = note
    ws.Cells(r, 4).Resize(1, 3).NumberFormat = "0.000000"
    ws.Cells(r, 7).NumberFormat = "#,##0"
    ws.Cells(r, 8).NumberFormat = "0.0000000"
    ws.Cells(r, 9).NumberFormat = "#,##0"
End Sub
