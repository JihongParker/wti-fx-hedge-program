Attribute VB_Name = "Ratio_Optimization"
Option Explicit

' =====================================================================
'  Ratio_Optimization.bas
'  ---------------------------------------------------------------
'  Single, self-contained module for the static WTI/FX hedge-ratio
'  optimization.  Import into Hedge_Optimization.xlsm and run:
'
'      OPT_RunAll        - builds the model sheet and runs all solves
'
'  or run pieces individually:
'
'      OPT_BuildModelSheet
'      OPT_RiskMin_EU    / OPT_RiskMin_AM
'      OPT_CostMin_EU    / OPT_CostMin_AM        (no risk cap)
'      OPT_CostMinCap_EU / OPT_CostMinCap_AM     (risk cap = last risk-min)
'
'  Everything is rebuilt from RAW parameters on every run.  No value
'  in this module or on the model sheet is hard-coded: all inputs are
'  live formula links to the source sheets (Encoding, Raw_Timeseries,
'  Black76, GK, LSMC), and the Solver model is registered from scratch
'  by code each time.
'
'  ---------------------------------------------------------------
'  DEFECTS OF THE ORIGINAL WORKBOOK, AND HOW THIS MODULE FIXES THEM
'  ---------------------------------------------------------------
'  (1) Pricing!C15 parenthesis error.  The stored formula was
'          =(1-(LSMC!J13)*Encoding!B9*MAX(0,(Encoding!B19-Encoding!B3)))
'      i.e.  1 - w1*Q_oil*gap   instead of   (1-w1)*Q_oil*gap.
'      FIX: the American unhedged-WTI loss is written correctly as
'          (1-w1) * Q_oil * MAX(0, Stress_WTI - S_WTI) * Stress_KRW.
'
'  (2) Pricing!C15 variable mixing.  The WTI leg's unhedged loss used
'      the FX stress gap (Stress_KRW - S_KRW) times the OIL quantity.
'      FIX: the WTI leg uses the WTI stress gap (Stress_WTI - S_WTI),
'      converted at Stress_KRW, mirroring the European C5 structure.
'
'  (3) Degenerate cost minimization.  Because of (1)+(2) the American
'      cost function carried no unhedged-loss penalty of meaningful
'      size, so a pure cost minimizer returned (w1,w2)=(0,0) - "hedge
'      nothing".  With the corrected penalty the cost minimum is the
'      economically sensible corner (1,0); see reference solutions.
'
'  (4) Broken #REF! Solver registrations on the Pricing sheet
'      (solver_adj / solver_lhs1 / solver_rhs1 all point at deleted
'      cells).  FIX: this module never relies on stored solver names;
'      it calls SolverReset and re-registers the full model on every
'      run, so it cannot rot when cells move.
'
'  (5) Evolutionary-solver imprecision.  The stored GA solutions sat
'      slightly inside the feasible boundary (up to 5e-4 off in the
'      ratios, KRW 75.3M / 663.3M off the efficient frontier at equal
'      risk).  FIX: both programs are CONVEX -- sigma_res is a norm of
'      an affine map of (w1,w2) and every constraint is affine -- so a
'      GRG Nonlinear local optimum is the global optimum.  This module
'      therefore uses GRG (Engine 1) with tight Precision/Convergence
'      and automatic scaling, which lands on the exact vertex; no
'      genetic algorithm is needed or wanted.
'
'  ---------------------------------------------------------------
'  REFERENCE SOLUTIONS (corrected model, budget = KRW 45bn;
'  independently computed with scipy SLSQP multistart, for checking
'  the Solver output -- expect agreement to ~1e-6):
'
'    European  risk-min : w1=0.970486  w2=0.029514  cost=45,000,000,000
'                         sigma_res=0.0916021   (sum & budget both active)
'    American  risk-min : w1=0.971581  w2=0.028419  cost=45,000,000,000
'                         sigma_res=0.0912158   (sum & budget both active)
'    European  cost-min : w1=1  w2=0  cost=42,736,755,546  sigma=0.0925765
'    American  cost-min : w1=1  w2=0  cost=33,477,827,504  sigma=0.0925765
'    Cost-min with risk cap = risk-min sigma returns the risk-min
'    vertex itself in both engines (exact-duality consistency check).
'
'  ---------------------------------------------------------------
'  WHY THE AMERICAN PREMIUMS STAY AT BASE (TODAY-SPOT) PRICING
'  ---------------------------------------------------------------
'  The LSMC Shapley premiums (LSMC!J9/J10) are priced at today's spot
'  (S_WTI = 78.94).  For THIS static allocation that is correct and
'  must not be changed:
'    - The budget cap governs TODAY's cash outflow: the premium is
'      paid at inception, at inception prices.  Re-pricing it at the
'      stress spot would charge the stress scenario twice (once in
'      the premium, once in the unhedged-loss terms).
'    - The optimizer consumes the premium only as a CONSTANT per unit
'      of w.  It never touches the LSMC regression surface (Beta_Mat)
'      or any delta.  The known fragility of the regression delta near
'      the KO barrier (FD-vs-regression divergence) contaminates
'      DELTAS -- local derivatives of a noisy surface -- not the
'      time-0 price, which is an average over the full path bank and
'      is far more robust.
'  The Beta_Mat / stress-delta problem is real, but it lives in the
'  DYNAMIC hedging layer (DeltaHedging_revised.bas): tracking hedge
'  performance and naked exposure AFTER the spot has moved to 113
'  requires a stress-centered LSMC re-fit (S1(0)=113, S2(0)=1550) and
'  the Run_FD_vs_Regression_Delta_Check gate.  None of that machinery
'  is (or should be) called from this static module.
'
'  ---------------------------------------------------------------
'  KO-SURVIVAL HAIRCUT (p_KO) -- REPORT-ONLY STRESS ADJUSTMENT
'  ---------------------------------------------------------------
'  One stress effect DOES leak into the static ledger: the workbook's
'  unhedged-loss terms assume the hedged fraction w is fully protected
'  under stress.  For the American KO structure that is optimistic --
'  at WTI 113 the upper barrier (120) is close, and a knocked-out call
'  protects nothing.  Parameter p_KO (Opt_Model!B22, default 0) is the
'  probability that the KO structure is dead under the stress scenario;
'  the report row "Stress-adjusted total cost" prices the American
'  protected fraction as w*(1-p_KO):
'      adj cost = premiums + (1 - w1*(1-p_KO))*UL_WTI
'                          + (1 - w2*(1-p_KO))*UL_FX
'  (the same WTI barrier kills both legs of the joint structure).
'  Fill p_KO from a stress-centered LSMC run, or as a shortcut use the
'  workbook's measured KO statistics (barrier-touch 0.4369; KO tally
'  net of early exercise 0.2309).
'  DELIBERATELY NOT A CONSTRAINT: the stress-loss floor p_KO*105.59bn
'  KRW is irreducible by any (w1,w2), so "stress-adjusted cost <= 45bn"
'  is infeasible for p_KO > ~10.9% (minimum possible adjusted cost is
'  57.86bn at p=0.2309 and 79.61bn at p=0.4369).  That infeasibility is
'  itself the finding: a KO-barrier hedge cannot deliver barrier-proof
'  stress protection.  The optimizer therefore keeps the workbook's
'  TOTAL COST <= Budget constraint unchanged, and the adjusted figure
'  is computed and logged for risk reporting only.
'
'  Requires: Solver add-in enabled (Tools > Add-ins > Solver).
' =====================================================================

Private Const SH As String = "Opt_Model"

' parameter cells (column B of Opt_Model)
Private Const P_SWTI As String = "$B$4"      ' WTI spot
Private Const P_SKRW As String = "$B$5"      ' USD/KRW spot
Private Const P_QOIL As String = "$B$6"      ' monthly oil need (bbl)
Private Const P_QUSD As String = "$B$7"      ' monthly USD need
Private Const P_WACC As String = "$B$8"
Private Const P_BUDG As String = "$B$9"
Private Const P_TOIL As String = "$B$10"
Private Const P_TFX As String = "$B$11"
Private Const P_STWTI As String = "$B$12"    ' stress WTI
Private Const P_STKRW As String = "$B$13"    ' stress KRW
Private Const P_S1EU As String = "$B$14"     ' sigma1 (historical, EU)
Private Const P_S1AM As String = "$B$15"     ' sigma1 (diffusive, AM)
Private Const P_S2 As String = "$B$16"       ' sigma2 (FX)
Private Const P_RHO As String = "$B$17"
Private Const P_B76 As String = "$B$18"      ' Black-76 WTI call (USD/bbl)
Private Const P_GK As String = "$B$19"       ' GK FX call (KRW/USD)
Private Const P_SHW As String = "$B$20"      ' LSMC Shapley WTI (KRW/bbl)
Private Const P_SHF As String = "$B$21"      ' LSMC Shapley FX (KRW/USD)
Private Const P_PKO As String = "$B$22"      ' p_KO: stress KO probability (input, default 0)

' engine blocks: column C = European, column D = American
Private Const R_W1 As Long = 24              ' decision w1
Private Const R_W2 As Long = 25              ' decision w2
Private Const R_SUM As Long = 26             ' w1+w2
Private Const R_PRW As Long = 28             ' premium WTI
Private Const R_PRF As Long = 29             ' premium FX
Private Const R_ULW As Long = 30             ' unhedged WTI loss (FIXED)
Private Const R_ULF As Long = 31             ' unhedged FX loss
Private Const R_COST As Long = 32            ' TOTAL COST  (cost-min target)
Private Const R_GMVP As Long = 33            ' sigma_res   (risk-min target)
Private Const R_SADJ As Long = 34            ' stress-adjusted cost (report-only)
Private Const R_CAP As Long = 35             ' optional risk cap input
Private Const R_LOG As Long = 38             ' results log header row
Private pSWTI As Double, pSKRW As Double, pQoil As Double, pQusd As Double
Private pWACC As Double, pBudget As Double, pToil As Double, pTfx As Double
Private pStWTI As Double, pStKRW As Double
Private pS1EU As Double, pS1AM As Double, pS2 As Double, pRho As Double
Private pB76 As Double, pGK As Double, pShW As Double, pShF As Double

' =====================================================================
'  PUBLIC ENTRY POINTS
' =====================================================================

Public Sub OPT_RunAll()
    OPT_BuildModelSheet
    SolveOne "C", "European", "risk-min"
    SolveOne "D", "American", "risk-min"
    ' copy the freshly solved risk-min sigmas into the cap cells so the
    ' capped cost-min runs answer: "cheapest allocation, same risk"
    With Worksheets(SH)
        .Range("C" & R_CAP).Value = .Range("C" & R_GMVP).Value
        .Range("D" & R_CAP).Value = .Range("D" & R_GMVP).Value
    End With
    SolveOne "C", "European", "cost-min-cap"
    SolveOne "D", "American", "cost-min-cap"
    SolveOne "C", "European", "cost-min"
    SolveOne "D", "American", "cost-min"
    Worksheets(SH).Activate
    MsgBox "All six solves finished. See the results log on '" & SH & "'.", _
           vbInformation, "Ratio_Optimization"
End Sub

Public Sub OPT_RiskMin_EU()
    EnsureModel
    SolveOne "C", "European", "risk-min"
End Sub

Public Sub OPT_RiskMin_AM()
    EnsureModel
    SolveOne "D", "American", "risk-min"
End Sub

Public Sub OPT_CostMin_EU()
    EnsureModel
    SolveOne "C", "European", "cost-min"
End Sub

Public Sub OPT_CostMin_AM()
    EnsureModel
    SolveOne "D", "American", "cost-min"
End Sub

Public Sub OPT_CostMinCap_EU()
    EnsureModel
    SolveOne "C", "European", "cost-min-cap"
End Sub

Public Sub OPT_CostMinCap_AM()
    EnsureModel
    SolveOne "D", "American", "cost-min-cap"
End Sub

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
        .Range("A1").Value = "STATIC HEDGE-RATIO OPTIMIZATION (self-contained; rebuilt by Ratio_Optimization.bas)"
        .Range("A1").Font.Bold = True

        ' ---- parameters: live links to raw data ------------------
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
        PutParam ws, 20, "P_Shapley_WTI (KRW/bbl)", "=LSMC!J9"
        PutParam ws, 21, "P_Shapley_FX (KRW/USD)", "=LSMC!J10"
        ' p_KO is a manual INPUT (not a link): probability the American KO
        ' structure is dead under the stress scenario.  0 = workbook-verbatim
        ' behaviour.  Fill from a stress-centered LSMC run (or, as a rough
        ' shortcut, the measured KO stats: 0.4369 barrier-touch / 0.2309 net).
        .Cells(22, 1).Value = "p_KO (stress KO prob., INPUT; 0 = off)"
        .Range("B22").Value = 0#

        ' ---- engine blocks ---------------------------------------
        .Range("C23").Value = "EUROPEAN": .Range("D23").Value = "AMERICAN"
        .Range("C23:D23").Font.Bold = True
        .Cells(R_W1, 1).Value = "w1  (WTI hedge ratio)  [decision]"
        .Cells(R_W2, 1).Value = "w2  (FX hedge ratio)   [decision]"
        .Cells(R_SUM, 1).Value = "w1 + w2"
        .Cells(R_PRW, 1).Value = "Premium cost, WTI leg"
        .Cells(R_PRF, 1).Value = "Premium cost, FX leg"
        .Cells(R_ULW, 1).Value = "Unhedged WTI stress loss  [FIXED: (1-w1)*Q_oil*MAX(0,StressWTI-S_WTI)*StressKRW]"
        .Cells(R_ULF, 1).Value = "Unhedged FX stress loss"
        .Cells(R_COST, 1).Value = "TOTAL COST  (cost-min objective)"
        .Cells(R_GMVP, 1).Value = "sigma_res   (risk-min objective)"
        .Cells(R_SADJ, 1).Value = "Stress-adjusted total cost (KO haircut p_KO; report-only, NOT constrained)"
        .Cells(R_CAP, 1).Value = "Risk cap for capped cost-min (blank = uncapped)"

        ' decision starting point (any interior feasible-ish point)
        .Range("C" & R_W1).Value = 0.9:  .Range("D" & R_W1).Value = 0.9
        .Range("C" & R_W2).Value = 0.05: .Range("D" & R_W2).Value = 0.05
        .Range("C" & R_SUM).Formula = "=C" & R_W1 & "+C" & R_W2
        .Range("D" & R_SUM).Formula = "=D" & R_W1 & "+D" & R_W2

        ' European cost block (structure of Pricing!C1:C9, verbatim)
        .Range("C" & R_PRW).Formula = "=C" & R_W1 & "*" & P_QOIL & "*" & P_B76 & "*" & P_SKRW & "*(1+" & P_WACC & "*" & P_TOIL & ")"
        .Range("C" & R_PRF).Formula = "=C" & R_W2 & "*" & P_QUSD & "*" & P_GK & "*(1+" & P_WACC & "*" & P_TFX & ")"
        .Range("C" & R_ULW).Formula = "=(1-C" & R_W1 & ")*" & P_QOIL & "*MAX(0," & P_STWTI & "-" & P_SWTI & ")*" & P_STKRW
        .Range("C" & R_ULF).Formula = "=(1-C" & R_W2 & ")*" & P_QUSD & "*MAX(0," & P_STKRW & "-" & P_SKRW & ")"

        ' American cost block (Shapley premiums; unhedged legs CORRECTED
        ' per header notes (1) and (2))
        .Range("D" & R_PRW).Formula = "=D" & R_W1 & "*" & P_QOIL & "*" & P_SHW & "*EXP(" & P_WACC & "*" & P_TOIL & ")"
        .Range("D" & R_PRF).Formula = "=D" & R_W2 & "*" & P_QUSD & "*" & P_SHF & "*EXP(" & P_WACC & "*" & P_TFX & ")"
        .Range("D" & R_ULW).Formula = "=(1-D" & R_W1 & ")*" & P_QOIL & "*MAX(0," & P_STWTI & "-" & P_SWTI & ")*" & P_STKRW
        .Range("D" & R_ULF).Formula = "=(1-D" & R_W2 & ")*" & P_QUSD & "*MAX(0," & P_STKRW & "-" & P_SKRW & ")"

        .Range("C" & R_COST).Formula = "=SUM(C" & R_PRW & ":C" & R_ULF & ")"
        .Range("D" & R_COST).Formula = "=SUM(D" & R_PRW & ":D" & R_ULF & ")"

        ' residual volatility (identical to Encoding!C24 / LSMC!J15 form)
        .Range("C" & R_GMVP).Formula = GmvpFormula("C", P_S1EU)
        .Range("D" & R_GMVP).Formula = GmvpFormula("D", P_S1AM)

        ' stress-adjusted total cost (report-only; see header notes).
        ' European leg is a vanilla Black-76 call (no barrier): adjusted
        ' cost = TOTAL COST.  American leg: protected fraction w*(1-p_KO).
        .Range("C" & R_SADJ).Formula = "=C" & R_COST
        .Range("D" & R_SADJ).Formula = _
            "=D" & R_PRW & "+D" & R_PRF & _
            "+(1-D" & R_W1 & "*(1-" & P_PKO & "))*" & P_QOIL & "*MAX(0," & P_STWTI & "-" & P_SWTI & ")*" & P_STKRW & _
            "+(1-D" & R_W2 & "*(1-" & P_PKO & "))*" & P_QUSD & "*MAX(0," & P_STKRW & "-" & P_SKRW & ")"

        ' formats
        .Range("B4:B22").NumberFormat = "General"
        .Range("C" & R_PRW & ":D" & R_COST).NumberFormat = "#,##0"
        .Range("C" & R_SADJ & ":D" & R_SADJ).NumberFormat = "#,##0"
        .Range("C" & R_W1 & ":D" & R_SUM).NumberFormat = "0.000000"
        .Range("C" & R_GMVP & ":D" & R_GMVP).NumberFormat = "0.0000000"
        .Range("C" & R_CAP & ":D" & R_CAP).NumberFormat = "0.0000000"

        ' results log header
        .Cells(R_LOG, 1).Resize(1, 10).Value = Array("timestamp", "engine", "mode", _
            "w1", "w2", "w1+w2", "total cost (KRW)", "sigma_res", _
            "stress-adj cost (KRW)", "solver msg code")
        .Cells(R_LOG, 1).Resize(1, 10).Font.Bold = True

        .Columns("A").ColumnWidth = 62
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
    On Error Resume Next
    Dim ok As Boolean
    ok = Not Worksheets(SH) Is Nothing
    On Error GoTo 0
    If Not ok Then OPT_BuildModelSheet
End Sub

' =====================================================================
'  SOLVER CORE  (fresh registration on every call -- nothing stored)
' =====================================================================

Private Sub SolveOne(col As String, engineName As String, mode As String)
    Dim ws As Worksheet, objCell As String, wCells As String
    Dim capCell As String, res As Variant, tag As String
    Dim cap As Double, sw1 As Double, sw2 As Double
    Dim gw1 As Double, gw2 As Double, gOk As Boolean, useGrid As Boolean

    Set ws = Worksheets(SH)
    ws.Activate
    LoadParams ws
    wCells = "$" & col & "$" & R_W1 & ":$" & col & "$" & R_W2
    capCell = "$" & col & "$" & R_CAP

    cap = 0
    If mode = "cost-min-cap" Then
        If Not IsNumeric(ws.Range(col & R_CAP).Value) Or IsEmpty(ws.Range(col & R_CAP).Value) _
           Or ws.Range(col & R_CAP).Value <= 0 Then
            MsgBox "Risk-cap cell " & capCell & " is empty. Run the risk-min first " & _
                   "(OPT_RunAll does this automatically) or type a cap.", vbExclamation
            Exit Sub
        End If
        cap = ws.Range(col & R_CAP).Value
    End If

    ' reset decision cells to a clean interior start
    ws.Range(col & R_W1).Value = 0.9
    ws.Range(col & R_W2).Value = 0.05

    If mode = "risk-min" Then
        objCell = "$" & col & "$" & R_GMVP
    Else
        objCell = "$" & col & "$" & R_COST
    End If

    ' -----------------------------------------------------------------
    ' Attempt 1: Excel Solver (GRG Nonlinear).  On some builds (notably
    ' Mac Excel) the Application.Run Solver calls return 0 WITHOUT
    ' actually solving, so the result below is never trusted blindly:
    ' it is verified against the built-in grid optimizer and replaced
    ' if infeasible or beaten.
    ' -----------------------------------------------------------------
    res = "solver-error"
    On Error GoTo solverFailed
    EnsureSolverAddin
    Application.Run SolverName("SolverReset")
    ' objective: minimize (MaxMinVal 2), GRG Nonlinear (Engine 1)
    Application.Run SolverName("SolverOk"), objCell, 2, 0, wCells, 1, "GRG Nonlinear"
    ' box constraints
    Application.Run SolverName("SolverAdd"), wCells, 1, "1"   ' w <= 1
    Application.Run SolverName("SolverAdd"), wCells, 3, "0"   ' w >= 0
    ' allocation envelope
    Application.Run SolverName("SolverAdd"), "$" & col & "$" & R_SUM, 1, "1"
    If mode = "risk-min" Then
        Application.Run SolverName("SolverAdd"), "$" & col & "$" & R_COST, 1, P_BUDG
    ElseIf mode = "cost-min-cap" Then
        Application.Run SolverName("SolverAdd"), "$" & col & "$" & R_GMVP, 1, capCell
    End If
    ' mode = "cost-min": box + sum only (the re-posed broken Pricing model)

    ' SolverOptions (positional): MaxTime, Iterations, Precision,
    ' AssumeLinear, StepThru, Estimates, Derivatives, SearchOption,
    ' IntTolerance, Scaling, Convergence, AssumeNonNeg.
    On Error Resume Next   ' tolerate signature drift across Excel builds
    Application.Run SolverName("SolverOptions"), 1000, 32767, 1E-08, False, False, _
        1, 1, 1, 1, True, 1E-07, False
    On Error GoTo solverFailed
    res = Application.Run(SolverName("SolverSolve"), True)
    Application.Run SolverName("SolverFinish"), 1   ' keep final values
    GoTo solverDone
solverFailed:
    res = "solver-error"
    Err.Clear
    Resume solverDone
solverDone:
    On Error GoTo 0
    sw1 = ws.Range(col & R_W1).Value
    sw2 = ws.Range(col & R_W2).Value

    ' -----------------------------------------------------------------
    ' Attempt 2 (verification + fallback): pure-VBA 3-stage grid
    ' refinement, no add-in required, deterministic, ~3e-8 resolution.
    ' Both programs are convex, so the global grid winner is the
    ' global optimum to within the grid resolution.
    ' -----------------------------------------------------------------
    GridSolve col, mode, cap, gw1, gw2, gOk

    useGrid = True
    If IsFeas(col, mode, sw1, sw2, cap, 1E-06) Then
        If EvalObj(col, mode, sw1, sw2) <= EvalObj(col, mode, gw1, gw2) * (1 + 1E-09) Then
            useGrid = False        ' Solver point is feasible and as good
        End If
    End If

    If useGrid And gOk Then
        ws.Range(col & R_W1).Value = gw1
        ws.Range(col & R_W2).Value = gw2
        tag = CStr(res) & " -> grid-fallback"
    ElseIf useGrid Then
        tag = CStr(res) & " (grid failed too - inspect)"
    Else
        tag = CStr(res) & " (solver verified)"
    End If

    LogResult ws, engineName, mode, col, tag
End Sub

' =====================================================================
'  BUILT-IN OPTIMIZER (verification + fallback; no Solver add-in)
' =====================================================================

' module-level parameter cache, loaded from Opt_Model!B4:B21


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
End Sub

Private Function Mx(x As Double) As Double
    If x > 0 Then Mx = x Else Mx = 0
End Function

' identical arithmetic to the Opt_Model sheet formulas
Private Function EvalCost(col As String, w1 As Double, w2 As Double) As Double
    If col = "C" Then   ' European: Black-76 / GK premiums
        EvalCost = w1 * pQoil * pB76 * pSKRW * (1 + pWACC * pToil) _
                 + w2 * pQusd * pGK * (1 + pWACC * pTfx) _
                 + (1 - w1) * pQoil * Mx(pStWTI - pSWTI) * pStKRW _
                 + (1 - w2) * pQusd * Mx(pStKRW - pSKRW)
    Else                ' American: LSMC Shapley premiums (corrected legs)
        EvalCost = w1 * pQoil * pShW * Exp(pWACC * pToil) _
                 + w2 * pQusd * pShF * Exp(pWACC * pTfx) _
                 + (1 - w1) * pQoil * Mx(pStWTI - pSWTI) * pStKRW _
                 + (1 - w2) * pQusd * Mx(pStKRW - pSKRW)
    End If
End Function

Private Function EvalGmvp(col As String, w1 As Double, w2 As Double) As Double
    Dim s1 As Double, h1 As Double, h2 As Double
    If col = "C" Then s1 = pS1EU Else s1 = pS1AM
    h1 = 1 - w1: h2 = 1 - w2
    EvalGmvp = Sqr(h1 * h1 * s1 * s1 + h2 * h2 * pS2 * pS2 _
                   + 2 * h1 * h2 * s1 * pS2 * pRho)
End Function

Private Function EvalObj(col As String, mode As String, w1 As Double, w2 As Double) As Double
    If mode = "risk-min" Then
        EvalObj = EvalGmvp(col, w1, w2)
    Else
        EvalObj = EvalCost(col, w1, w2)
    End If
End Function

Private Function IsFeas(col As String, mode As String, w1 As Double, w2 As Double, _
                        cap As Double, tol As Double) As Boolean
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

' three-stage refined grid search over [0,1]^2 (801 points per axis per
' stage; final resolution ~3e-8 in each ratio)
Private Sub GridSolve(col As String, mode As String, cap As Double, _
                      ByRef bestW1 As Double, ByRef bestW2 As Double, _
                      ByRef ok As Boolean)
    Dim lo1 As Double, hi1 As Double, lo2 As Double, hi2 As Double
    Dim stage As Long, i As Long, j As Long, n As Long
    Dim w1 As Double, w2 As Double, f As Double, bestF As Double, hw As Double

    n = 800
    lo1 = 0: hi1 = 1: lo2 = 0: hi2 = 1
    ok = False

    For stage = 1 To 3
        bestF = 1E+308
        For i = 0 To n
            w1 = lo1 + (hi1 - lo1) * i / n
            For j = 0 To n
                w2 = lo2 + (hi2 - lo2) * j / n
                If IsFeas(col, mode, w1, w2, cap, 1E-12) Then
                    f = EvalObj(col, mode, w1, w2)
                    If f < bestF Then
                        bestF = f: bestW1 = w1: bestW2 = w2
                    End If
                End If
            Next j
        Next i
        If bestF >= 1E+307 Then Exit Sub   ' no feasible point found
        hw = 2 * (hi1 - lo1) / n
        lo1 = MaxD(0, bestW1 - hw): hi1 = MinD(1, bestW1 + hw)
        hw = 2 * (hi2 - lo2) / n
        lo2 = MaxD(0, bestW2 - hw): hi2 = MinD(1, bestW2 + hw)
    Next stage
    ok = True
End Sub

Private Function MaxD(a As Double, b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function MinD(a As Double, b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Sub LogResult(ws As Worksheet, engineName As String, mode As String, _
                      col As String, res As String)
    Dim r As Long
    r = R_LOG + 1
    Do While ws.Cells(r, 1).Value <> "": r = r + 1: Loop
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = engineName
    ws.Cells(r, 3).Value = mode
    ws.Cells(r, 4).Value = ws.Range(col & R_W1).Value
    ws.Cells(r, 5).Value = ws.Range(col & R_W2).Value
    ws.Cells(r, 6).Value = ws.Range(col & R_SUM).Value
    ws.Cells(r, 7).Value = ws.Range(col & R_COST).Value
    ws.Cells(r, 8).Value = ws.Range(col & R_GMVP).Value
    ws.Cells(r, 9).Value = ws.Range(col & R_SADJ).Value
    ws.Cells(r, 10).Value = res         ' e.g. "0 (solver verified)" / "-> grid-fallback"
    ws.Cells(r, 4).Resize(1, 3).NumberFormat = "0.000000"
    ws.Cells(r, 7).NumberFormat = "#,##0"
    ws.Cells(r, 8).NumberFormat = "0.0000000"
    ws.Cells(r, 9).NumberFormat = "#,##0"
End Sub

' =====================================================================
'  SOLVER ADD-IN PLUMBING
' =====================================================================

Private Function SolverName(proc As String) As String
    SolverName = "Solver.xlam!" & proc
End Function

Private Sub EnsureSolverAddin()
    On Error Resume Next
    Application.AddIns("Solver Add-In").Installed = True
    On Error GoTo 0
End Sub
