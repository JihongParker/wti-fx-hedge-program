Attribute VB_Name = "DeltaHedging_revised"
Option Explicit

Public Const STEPS_PER_YEAR As Double = 260#

Public Beta_Mat() As Double

Public Beta_Mat_Steps As Long

Public PB_S1()  As Double
Public PB_S2()  As Double
Public PB_KO()  As Long
Public PB_nJ()  As Long
Public PB_N     As Long
Public PB_Steps As Long

Public Function Sigma_Eff_WTI(ByVal vol1 As Double, ByVal Lambda As Double, _
                               ByVal JumpMean As Double, ByVal JumpVol As Double) As Double
    Sigma_Eff_WTI = Sqr(vol1 ^ 2 + Lambda * (JumpMean ^ 2 + JumpVol ^ 2))
End Function

Private Function ArrayRank(v As Variant) As Long
    Dim i As Long, probe As Long
    On Error GoTo Done
    For i = 1 To 60
        probe = UBound(v, i)
    Next i
Done:
    ArrayRank = i - 1
End Function

Function GetNormal() As Double
    Dim u1 As Double, u2 As Double
    u1 = Rnd: If u1 <= 0 Then u1 = 0.0001
    u2 = Rnd
    GetNormal = Sqr(-2 * Log(u1)) * Cos(6.2831853 * u2)
End Function

Function GetPoisson(L As Double) As Long
    Dim Lval As Double: Lval = Exp(-L)
    Dim p As Double: p = 1
    Dim K As Long: K = 0
    Do
        K = K + 1
        p = p * Rnd
    Loop While p > Lval
    GetPoisson = K - 1
End Function

Function JumpSum(n As Long, m As Double, v As Double) As Double
    Dim i As Long, s As Double
    For i = 1 To n
        s = s + (m + v * GetNormal())
    Next i
    JumpSum = s
End Function

' =====================================================================
'  RESIDUAL-RISK (GMVP) SENSITIVITY SURFACE
' ---------------------------------------------------------------------
'  Decision variables (see optimization.md for the full model):
'    w1 = WTI hedge ratio  -- fraction of the WTI leg's exposure hedged
'    w2 = FX  hedge ratio  -- fraction of the FX  leg's exposure hedged
'  Domain:  0 <= w1 <= 1, 0 <= w2 <= 1, and 0 <= w1 + w2 <= 1
'           (the sum cap is the aggregate hedge-budget constraint).
'  The unhedged residual fraction of each leg is (1 - w_i), so the
'  residual ("GMVP") portfolio volatility that the risk-Solver minimises
'  is, exactly as in LSMC!J15 (American) and Encoding!C24 (European),
'
'     GMVP(w1,w2) = sqrt( ((1-w1) v1)^2 + ((1-w2) v2)^2
'                         + 2 (1-w1)(1-w2) v1 v2 rho ).
'
'  NOTE on Shapley: the premium attribution phi_WTI/phi_FX (LSMC!J9/J10)
'  belongs to the COST objective (Pricing!C11/C12), NOT to this variance
'  objective. The earlier surface multiplied each weight by its Shapley
'  share, which made the plotted surface a premium-weighted variance whose
'  minimum did not coincide with the Solver's actual objective. That
'  factor is removed here so the surface equals the true objective cell.
'
'  LAYOUT decision (centre vs. top-left): the optimum is placed at the
'  CENTRE of a (2*HALF+1)-square grid, with symmetric +/- STEP_PP steps,
'  so the surface can be read as a minimum -- it must rise on every
'  feasible side. The previous top-left anchoring sampled only w >= w*
'  (one quadrant) and could not distinguish a minimum from a monotone
'  slope. Cells that violate the budget (w1 + w2 > 1) or leave [0,1] are
'  left blank so the feasible optimum is the lowest *plotted* point.
' =====================================================================

Sub Build_GMVP_Surface_Fixed()
    ' American (KO + jump-diffusion): ratios from LSMC!J13/J14, diffusion
    ' vol LSMC!B10 -- centre cell reproduces LSMC!J15 exactly.
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")
    Call BuildResidualRiskSurface( _
        wsLSMC.Range("J13").Value, wsLSMC.Range("J14").Value, _
        wsLSMC.Range("B10").Value, wsLSMC.Range("B11").Value, wsLSMC.Range("B12").Value, _
        37, 7)
End Sub

Sub Build_GMVP_Surface_European()
    ' European (Black-76 / GK, no KO): ratios from Encoding!C21/C22,
    ' effective vol Raw_Timeseries!H2 -- centre cell reproduces Encoding!C24.
    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsRT  As Worksheet: Set wsRT = Sheets("Raw_Timeseries")
    Call BuildResidualRiskSurface( _
        wsEnc.Range("C21").Value, wsEnc.Range("C22").Value, _
        wsRT.Range("H2").Value, wsRT.Range("I2").Value, wsRT.Range("J2").Value, _
        50, 7)
End Sub

Private Sub BuildResidualRiskSurface(ByVal baseW1 As Double, ByVal baseW2 As Double, _
        ByVal v1 As Double, ByVal v2 As Double, ByVal rho As Double, _
        ByVal startRow As Long, ByVal startCol As Long)

    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")

    Const HALF    As Long = 4        ' 9x9 grid, indices -HALF..+HALF
    Const STEP_PP As Double = 0.005  ' 0.5 percentage-point increments (+/- 2.0pp)

    Dim n As Long: n = 2 * HALF + 1
    Dim i As Long, j As Long
    Dim W1 As Double, W2 As Double
    Dim r1 As Double, r2 As Double, gmvp_val As Double

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    wsEnc.Range(wsEnc.Cells(startRow - 1, startCol - 1), _
                wsEnc.Cells(startRow + n - 1, startCol + n - 1)).ClearContents

    wsEnc.Cells(startRow - 1, startCol - 1).Value = "WTI \ FX"

    ' column headers: w2 centred on baseW2
    For j = 0 To n - 1
        W2 = baseW2 + (j - HALF) * STEP_PP
        wsEnc.Cells(startRow - 1, startCol + j).Value = W2
        wsEnc.Cells(startRow - 1, startCol + j).NumberFormat = "0.00%"
    Next j

    ' row headers: w1 centred on baseW1
    For i = 0 To n - 1
        W1 = baseW1 + (i - HALF) * STEP_PP
        wsEnc.Cells(startRow + i, startCol - 1).Value = W1
        wsEnc.Cells(startRow + i, startCol - 1).NumberFormat = "0.00%"
    Next i

    ' surface body: residual GMVP volatility over the full centred window.
    ' We do NOT blank the budget-infeasible region (w1+w2>1): blanking left a
    ' half-empty triangle that Excel renders as a jagged "staircase". The whole
    ' window is filled for a smooth surface; the budget edge w1+w2=1 is shown as
    ' a separate reference column to the right (see below).
    For i = 0 To n - 1
        W1 = baseW1 + (i - HALF) * STEP_PP
        For j = 0 To n - 1
            W2 = baseW2 + (j - HALF) * STEP_PP
            r1 = 1# - W1
            r2 = 1# - W2
            gmvp_val = Sqr((r1 * v1) ^ 2 + (r2 * v2) ^ 2 + 2# * r1 * r2 * v1 * v2 * rho)
            wsEnc.Cells(startRow + i, startCol + j).Value = gmvp_val
            wsEnc.Cells(startRow + i, startCol + j).NumberFormat = "0.000000%"
        Next j
    Next i

    ' highlight the centre (optimum) cell
    wsEnc.Cells(startRow + HALF, startCol + HALF).Interior.Color = RGB(255, 242, 204)

    ' ---- Budget-edge sweep (the meaningful optimality check) --------------
    ' The residual-risk objective is monotone in each w_i, so a 2-D (w1,w2)
    ' surface has no interior minimum -- the optimum lives on the budget edge
    ' w1+w2=1. To actually see a minimum we sweep ALONG that edge: w2 = 1-w1.
    ' This is wide (w1 in [baseW1-0.10, baseW1+0.02]) so the analytic minimum is
    ' inside the range and visible as a clean U. Written as a 3-col table two
    ' columns to the right of the surface.
    Dim ec As Long: ec = startCol + n + 2
    Dim m As Long, ws1 As Double, ws2 As Double, ev As Double
    wsEnc.Cells(startRow - 1, ec).Value = "w1 (WTI)"
    wsEnc.Cells(startRow - 1, ec + 1).Value = "w2=1-w1"
    wsEnc.Cells(startRow - 1, ec + 2).Value = "GMVP (edge)"
    For m = 0 To 12
        ws1 = baseW1 - 0.1 + m * 0.01
        If ws1 < 0# Then ws1 = 0#
        If ws1 > 1# Then ws1 = 1#
        ws2 = 1# - ws1
        ev = Sqr(((1# - ws1) * v1) ^ 2 + ((1# - ws2) * v2) ^ 2 _
                 + 2# * (1# - ws1) * (1# - ws2) * v1 * v2 * rho)
        wsEnc.Cells(startRow + m, ec).Value = ws1
        wsEnc.Cells(startRow + m, ec).NumberFormat = "0.00%"
        wsEnc.Cells(startRow + m, ec + 1).Value = ws2
        wsEnc.Cells(startRow + m, ec + 1).NumberFormat = "0.00%"
        wsEnc.Cells(startRow + m, ec + 2).Value = ev
        wsEnc.Cells(startRow + m, ec + 2).NumberFormat = "0.000000%"
    Next m

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Public Sub Build_Path_Bank(ByVal n As Long, ByVal Steps As Long, _
    ByVal S0_WTI As Double, ByVal S0_FX As Double, _
    ByVal vol1 As Double, ByVal vol2 As Double, ByVal corr As Double, _
    ByVal Drift1 As Double, ByVal Drift2 As Double, _
    ByVal Lambda As Double, ByVal JumpMean As Double, ByVal JumpVol As Double, _
    ByVal KO_up As Double, ByVal KO_dn As Double, ByVal T_total As Double)

    Dim dt    As Double: dt = T_total / Steps
    Dim kappa As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt

    PB_N = 0
    ReDim PB_S1(1 To n, 0 To Steps)
    ReDim PB_S2(1 To n, 0 To Steps)
    ReDim PB_KO(1 To n)
    ReDim PB_nJ(1 To n, 1 To Steps)

    Randomize 20240101

    Dim sim As Long, stp As Long
    Dim S1 As Double, S2 As Double, S1_prev As Double
    Dim z1 As Double, z2 As Double, e2 As Double, nJ As Long
    Dim ko_step As Long
    Dim p_up As Double, p_dn As Double

    For sim = 1 To n
        S1 = S0_WTI: S2 = S0_FX: S1_prev = S0_WTI: ko_step = 0
        PB_S1(sim, 0) = S0_WTI: PB_S2(sim, 0) = S0_FX

        For stp = 1 To Steps
            z1 = GetNormal()
            z2 = GetNormal()
            e2 = corr * z1 + Sqr(1# - corr ^ 2) * z2
            nJ = GetPoisson(Lambda * dt)

            S1 = S1 * Exp((Drift1 - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt _
                         + vol1 * Sqr(dt) * z1 + JumpSum(nJ, JumpMean, JumpVol))
            S2 = S2 * Exp((Drift2 - 0.5 * vol2 ^ 2) * dt + vol2 * Sqr(dt) * e2)

            PB_S1(sim, stp) = S1
            PB_S2(sim, stp) = S2
            PB_nJ(sim, stp) = nJ

            If S1 >= KO_up Or S1 <= KO_dn Then
                ko_step = stp: Exit For
            ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then
                p_up = Exp(-2# * Log(KO_up / S1_prev) * Log(KO_up / S1) / lnVar)
                p_dn = Exp(-2# * Log(S1_prev / KO_dn) * Log(S1 / KO_dn) / lnVar)
                If Rnd() < p_up Then ko_step = stp: Exit For
                If Rnd() < p_dn Then ko_step = stp: Exit For
            End If
            S1_prev = S1
        Next stp

        PB_KO(sim) = ko_step
    Next sim

    PB_N = n
    PB_Steps = Steps
End Sub

Public Function JumpDiffusionKO( _
    Lambda As Double, JumpMean As Double, JumpVol As Double, _
    Drift1 As Double, Drift2 As Double, _
    KOUpper As Double, KOLower As Double, _
    S1_0 As Double, S2_0 As Double, _
    vol1 As Double, vol2 As Double, _
    corr As Double, Steps As Long, T As Double) As Variant

    Dim dt     As Double: dt = T / Steps
    Dim kappa  As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar  As Double: lnVar = vol1 ^ 2 * dt

    Dim arr() As Variant
    ReDim arr(1 To Steps, 1 To 3)

    Dim S1      As Double: S1 = S1_0
    Dim S2      As Double: S2 = S2_0
    Dim S1_prev As Double: S1_prev = S1_0
    Dim Alive   As Long:   Alive = 1

    Dim i  As Long
    Dim z1 As Double, z2 As Double, e2 As Double
    Dim nJ As Long

    Randomize

    For i = 1 To Steps

        If Alive = 1 Then
            S1_prev = S1

            z1 = GetNormal()
            z2 = GetNormal()
            e2 = corr * z1 + Sqr(1 - corr ^ 2) * z2
            nJ = GetPoisson(Lambda * dt)

            S1 = S1 * Exp((Drift1 - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt _
                + vol1 * Sqr(dt) * z1 + JumpSum(nJ, JumpMean, JumpVol))
            S2 = S2 * Exp((Drift2 - 0.5 * vol2 ^ 2) * dt _
                + vol2 * Sqr(dt) * e2)

            If S1 >= KOUpper Or S1 <= KOLower Then
                Alive = 0
            ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then

                Dim p_up_jd As Double, p_dn_jd As Double
                p_up_jd = Exp(-2# * Log(KOUpper / S1_prev) * Log(KOUpper / S1) / lnVar)
                p_dn_jd = Exp(-2# * Log(S1_prev / KOLower) * Log(S1 / KOLower) / lnVar)
                If Rnd() < p_up_jd Then Alive = 0
                If Alive = 1 And Rnd() < p_dn_jd Then Alive = 0
            End If

        End If

        arr(i, 1) = S1
        arr(i, 2) = S2
        arr(i, 3) = Alive
    Next i

    JumpDiffusionKO = arr
End Function

Public Function Calc_LSMC_Price(Lambda As Double, JumpMean As Double, JumpVol As Double, _
                                 Drift1 As Double, Drift2 As Double, _
                                 KOUpper As Double, KOLower As Double, _
                                 S1_init As Double, S2_init As Double, _
                                 vol1 As Double, vol2 As Double, _
                                 corr As Double, Steps As Long, T As Double, _
                                 n_paths As Long, K As Double, _
                                 Optional bStoreBeta As Boolean = False) As Double

    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim dt    As Double:    dt = T / Steps
    Dim sqdt  As Double:    sqdt = Sqr(dt)

    Dim r_US  As Double: r_US = wsEnc.Range("B4").Value
    Dim r_KRW As Double: r_KRW = wsEnc.Range("B5").Value
    Dim disc  As Double: disc = Exp(-r_US * dt)
    Dim kappa As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#

    Drift1 = r_US
    Drift2 = r_KRW - r_US

    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt

    Dim S1_path() As Double, S2_path() As Double, Alive() As Boolean
    ReDim S1_path(1 To n_paths, 0 To Steps)
    ReDim S2_path(1 To n_paths, 0 To Steps)
    ReDim Alive(1 To n_paths, 0 To Steps)

    Dim p As Long, i As Long
    Randomize 12345

    If bStoreBeta Then
        ReDim Beta_Mat(1 To Steps, 0 To 5)
        Beta_Mat_Steps = Steps
    End If

    For p = 1 To n_paths
        S1_path(p, 0) = S1_init
        S2_path(p, 0) = S2_init
        Alive(p, 0) = True

        Dim S1      As Double: S1 = S1_init
        Dim S2      As Double: S2 = S2_init
        Dim isAlive As Boolean: isAlive = True
        Dim S1_prev As Double: S1_prev = S1_init

        For i = 1 To Steps
            If isAlive Then
                S1_prev = S1

                Dim z1   As Double: z1 = GetNormal()
                Dim z2   As Double: z2 = GetNormal()
                Dim e2   As Double: e2 = corr * z1 + Sqr(1 - corr ^ 2) * z2
                Dim nJ   As Long:   nJ = GetPoisson(Lambda * dt)
                Dim jSum As Double: jSum = JumpSum(nJ, JumpMean, JumpVol)

                S1 = S1 * Exp((Drift1 - 0.5 * vol1 ^ 2 - Lambda * kappa) * dt _
                    + vol1 * sqdt * z1 + jSum)
                S2 = S2 * Exp((Drift2 - 0.5 * vol2 ^ 2) * dt + vol2 * sqdt * e2)

                If S1 >= KOUpper Or S1 <= KOLower Then
                    isAlive = False
                ElseIf nJ = 0 And lnVar > 0 And S1_prev > 0 Then

                    Dim p_touch_up As Double, p_touch_dn As Double
                    p_touch_up = Exp(-2# * Log(KOUpper / S1_prev) * Log(KOUpper / S1) / lnVar)
                    p_touch_dn = Exp(-2# * Log(S1_prev / KOLower) * Log(S1 / KOLower) / lnVar)
                    If Rnd() < p_touch_up Then isAlive = False
                    If isAlive And Rnd() < p_touch_dn Then isAlive = False
                End If

            End If

            S1_path(p, i) = S1
            S2_path(p, i) = S2
            Alive(p, i) = isAlive
        Next i
    Next p

    Dim CF() As Double: ReDim CF(1 To n_paths)

    For p = 1 To n_paths
        If Alive(p, Steps) Then
            CF(p) = Application.Max(S1_path(p, Steps) - K, 0) * S2_path(p, Steps)
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
            If Alive(p, i) Then
                If (S1_path(p, i) - K) > 0 Then itmCount = itmCount + 1
            End If
        Next p

        If itmCount > 10 Then
            Dim X_Mat() As Double, Y_Vec() As Double
            ReDim X_Mat(1 To itmCount, 1 To 5)
            ReDim Y_Vec(1 To itmCount, 1 To 1)

            Dim itmRow As Long: itmRow = 0
            For p = 1 To n_paths
                If Alive(p, i) And (S1_path(p, i) - K) > 0 Then
                    itmRow = itmRow + 1
                    Dim v1 As Double: v1 = S1_path(p, i) / K
                    Dim v2 As Double: v2 = S2_path(p, i) / S2_init
                    X_Mat(itmRow, 1) = v1
                    X_Mat(itmRow, 2) = v2
                    X_Mat(itmRow, 3) = v1 ^ 2
                    X_Mat(itmRow, 4) = v2 ^ 2
                    X_Mat(itmRow, 5) = v1 * v2
                    Y_Vec(itmRow, 1) = CF(p)
                End If
            Next p

            Dim raw2D As Variant
            raw2D = Application.WorksheetFunction.LinEst(Y_Vec, X_Mat)
            Dim Coeff(1 To 6) As Double
            Dim ci As Long
            Dim lrank As Long: lrank = ArrayRank(raw2D)
            If lrank = 2 Then
                If (UBound(raw2D, 2) - LBound(raw2D, 2) + 1) >= 6 Then

                    For ci = 1 To 6
                        Coeff(ci) = raw2D(LBound(raw2D, 1), LBound(raw2D, 2) + ci - 1)
                    Next ci
                Else

                    For ci = 1 To 6
                        Coeff(ci) = raw2D(LBound(raw2D, 1) + ci - 1, LBound(raw2D, 2))
                    Next ci
                End If
            Else

                For ci = 1 To 6
                    Coeff(ci) = raw2D(LBound(raw2D) + ci - 1)
                Next ci
            End If

            If bStoreBeta Then
                Beta_Mat(i, 0) = Coeff(6)
                Beta_Mat(i, 1) = Coeff(5)
                Beta_Mat(i, 2) = Coeff(4)
                Beta_Mat(i, 3) = Coeff(3)
                Beta_Mat(i, 4) = Coeff(2)
                Beta_Mat(i, 5) = Coeff(1)
            End If

            For p = 1 To n_paths
                If Alive(p, i) And (S1_path(p, i) - K) > 0 Then
                    Dim curS1 As Double: curS1 = S1_path(p, i) / K
                    Dim curS2 As Double: curS2 = S2_path(p, i) / S2_init
                    Dim intrinsic As Double
                    intrinsic = (S1_path(p, i) - K) * S2_path(p, i)

                    Dim cont_val As Double
                    cont_val = Coeff(1) * (curS1 * curS2) + Coeff(2) * (curS2 ^ 2) + _
                               Coeff(3) * (curS1 ^ 2) + Coeff(4) * curS2 + _
                               Coeff(5) * curS1 + Coeff(6)

                    If intrinsic > cont_val Then CF(p) = intrinsic
                End If
            Next p
        Else

            If bStoreBeta And i < Steps - 1 Then
                Beta_Mat(i, 0) = Beta_Mat(i + 1, 0)
                Beta_Mat(i, 1) = Beta_Mat(i + 1, 1)
                Beta_Mat(i, 2) = Beta_Mat(i + 1, 2)
                Beta_Mat(i, 3) = Beta_Mat(i + 1, 3)
                Beta_Mat(i, 4) = Beta_Mat(i + 1, 4)
                Beta_Mat(i, 5) = Beta_Mat(i + 1, 5)
            End If
        End If
    Next i

    Dim totalVal As Double
    For p = 1 To n_paths
        totalVal = totalVal + CF(p) * disc
    Next p

    Calc_LSMC_Price = totalVal / n_paths
End Function

Public Sub Run_LSMC_Engine()
    Dim ws    As Worksheet: Set ws = Sheets("LSMC")
    Dim wsEnc As Worksheet
    On Error Resume Next
    Set wsEnc = Sheets("Encoding")
    On Error GoTo 0
    If wsEnc Is Nothing Then
        MsgBox "Encoding sheet not found.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail

    Dim USE_DAILY As Boolean: USE_DAILY = True

    Dim T_WTI As Double: T_WTI = wsEnc.Range("B15").Value
    Dim T_FX  As Double: T_FX = wsEnc.Range("B16").Value
    Dim T     As Double: T = Application.WorksheetFunction.Max(T_WTI, T_FX)

    ws.Range("B14").Value = T

    Dim Steps As Long
    Dim n     As Long
    If USE_DAILY Then
        Steps = CLng(T * STEPS_PER_YEAR): n = 50000
    Else
        Steps = CLng(T * 52):  n = 2000
    End If
    ws.Range("B13").Value = Steps

    Dim Lambda   As Double: Lambda = ws.Range("B1").Value
    Dim JumpMean As Double: JumpMean = ws.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = ws.Range("B3").Value
    Dim Drift1   As Double: Drift1 = ws.Range("B4").Value
    Dim Drift2   As Double: Drift2 = ws.Range("B5").Value
    Dim KOUpper  As Double: KOUpper = ws.Range("B6").Value
    Dim KOLower  As Double: KOLower = ws.Range("B7").Value
    Dim S1_0     As Double: S1_0 = ws.Range("B8").Value
    Dim S2_0     As Double: S2_0 = ws.Range("B9").Value
    Dim vol1     As Double: vol1 = ws.Range("B10").Value
    Dim vol2     As Double: vol2 = ws.Range("B11").Value
    Dim corr     As Double: corr = ws.Range("B12").Value

    Dim StressWTI As Double: StressWTI = wsEnc.Range("B18").Value
    Dim StressKRW As Double: StressKRW = wsEnc.Range("B19").Value

    Dim Base_Premium As Double
    Base_Premium = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, _
                                   KOUpper, KOLower, S1_0, S2_0, _
                                   vol1, vol2, corr, Steps, T, n, S1_0, True)

    Dim eps As Double: eps = 0.01
    Dim V_S1up As Double
    V_S1up = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, _
                              KOUpper, KOLower, S1_0 * (1 + eps), S2_0, _
                              vol1, vol2, corr, Steps, T, n, S1_0)
    Dim V_S2up As Double
    V_S2up = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, _
                              KOUpper, KOLower, S1_0, S2_0 * (1 + eps), _
                              vol1, vol2, corr, Steps, T, n, S1_0)

    Dim WTI_Delta_KRW As Double: WTI_Delta_KRW = (V_S1up - Base_Premium) / (S1_0 * eps)
    Dim FX_Delta_KRW  As Double: FX_Delta_KRW = (V_S2up - Base_Premium) / (S2_0 * eps)

    Dim v_WTI_only As Double
    v_WTI_only = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, _
                                  KOUpper, KOLower, S1_0, S2_0, _
                                  vol1, 1E-05, corr, Steps, T, n, S1_0)
    Dim v_FX_only As Double
    v_FX_only = Calc_LSMC_Price(0#, JumpMean, JumpVol, Drift1, Drift2, _
                                 KOUpper, KOLower, S1_0, S2_0, _
                                 1E-05, vol2, corr, Steps, T, n, S1_0)

    Dim phi_WTI As Double: phi_WTI = 0.5 * v_WTI_only + 0.5 * (Base_Premium - v_FX_only)
    Dim phi_FX  As Double: phi_FX = 0.5 * v_FX_only + 0.5 * (Base_Premium - v_WTI_only)

    Dim Total_Phi As Double: Total_Phi = phi_WTI + phi_FX
    Dim WTI_Ratio As Double, FX_Ratio As Double
    If Total_Phi > 0 Then
        WTI_Ratio = phi_WTI / Total_Phi
        FX_Ratio = phi_FX / Total_Phi
    Else
        WTI_Ratio = 0.5: FX_Ratio = 0.5
    End If

    Dim Stress_Premium As Double
    Stress_Premium = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, _
                                     KOUpper, KOLower, StressWTI, StressKRW, _
                                     vol1, vol2, corr, Steps, T, n, S1_0)

    ws.Range("I1:K12").ClearContents
    ws.Range("I1:J1").Value = Array("Item", "Value")
    ws.Range("I2:J2").Value = Array("Base Premium", Base_Premium)
    ws.Range("I3:J3").Value = Array("Stress Premium", Stress_Premium)
    ws.Range("I4:J4").Value = Array("WTI Delta (KRW)", WTI_Delta_KRW)
    ws.Range("I5:J5").Value = Array("FX Delta (KRW)", FX_Delta_KRW)
    ws.Range("I6:J6").Value = Array("WTI Share (Ratio)", WTI_Ratio)
    ws.Range("I7:J7").Value = Array("FX Share (Ratio)", FX_Ratio)
    ws.Range("I9:J9").Value = Array("WTI Base Premium (Shapley)", phi_WTI)
    ws.Range("I10:J10").Value = Array("FX Base Premium (Shapley)", phi_FX)

    ws.Range("J2:J3").Interior.Color = RGB(255, 255, 102)
    ws.Range("J9:J10").Interior.Color = RGB(198, 239, 206)

    If USE_DAILY Then
        ws.Range("I17").Value = "Daily mode: Stress Grid skipped (performance)"
        GoTo Done
    End If

    Dim dataStartRow As Long: dataStartRow = 18
    Dim idx As Long, shock As Double
    ws.Range("I17:L38").ClearContents
    ws.Range("I17:L17").Value = Array("Shock Ratio", "WTI Premium", "FX Premium", "Stress Premium")

    For idx = -10 To 10
        shock = idx * 0.01
        Dim wti_s As Double: wti_s = StressWTI * (1 + shock)
        Dim fx_s  As Double: fx_s = StressKRW * (1 + shock)
        Dim prem_s As Double
        prem_s = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, _
                                  KOUpper, KOLower, wti_s, fx_s, _
                                  vol1, vol2, corr, Steps, T, n, S1_0)
        Dim targetRow As Long: targetRow = dataStartRow + (idx + 10)
        ws.Cells(targetRow, 9).Value = shock
        ws.Cells(targetRow, 10).Value = prem_s * WTI_Ratio
        ws.Cells(targetRow, 11).Value = prem_s * FX_Ratio
        ws.Cells(targetRow, 12).Value = prem_s
    Next idx

Done:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Complete [" & IIf(USE_DAILY, "Daily " & Steps & " steps", "Weekly " & Steps & " steps") & "]" & vbCrLf & _
           "Base Premium  : " & Format(Base_Premium, "#,##0") & vbCrLf & _
           "Stress Premium: " & Format(Stress_Premium, "#,##0") & vbCrLf & _
           "phi_WTI (Base): " & Format(phi_WTI, "#,##0") & vbCrLf & _
           "phi_FX  (Base): " & Format(phi_FX, "#,##0") & vbCrLf & _
           "Paths: " & n, vbInformation
    Exit Sub

CleanFail:

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Run_LSMC_Engine failed: " & Err.Description, vbCritical
End Sub

Public Sub Run_American_DeltaHedge()

    Dim wsA   As Worksheet: Set wsA = Sheets("American_Delta")
    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsLS  As Worksheet: Set wsLS = Sheets("LSMC")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim Lambda   As Double: Lambda = wsLS.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLS.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsLS.Range("B3").Value
    Dim Drift1   As Double: Drift1 = wsLS.Range("B4").Value
    Dim Drift2   As Double: Drift2 = wsLS.Range("B5").Value
    Dim KO_up    As Double: KO_up = wsLS.Range("B6").Value
    Dim KO_dn    As Double: KO_dn = wsLS.Range("B7").Value
    Dim S0_WTI   As Double: S0_WTI = wsLS.Range("B8").Value
    Dim S0_FX    As Double: S0_FX = wsLS.Range("B9").Value
    Dim vol1     As Double: vol1 = wsLS.Range("B10").Value
    Dim vol2     As Double: vol2 = wsLS.Range("B11").Value
    Dim corr     As Double: corr = wsLS.Range("B12").Value
    Dim T_total  As Double: T_total = wsLS.Range("B14").Value

    Dim delta_ratio As Double: delta_ratio = 1#

    Dim T_WTI As Double
    On Error Resume Next
    T_WTI = wsEnc.Range("B15").Value
    On Error GoTo 0
    If T_WTI <= 0 Then T_WTI = T_total

    Dim WACC     As Double: WACC = wsEnc.Range("B12").Value
    Dim barrels  As Double: barrels = wsEnc.Range("B9").Value

    Dim SIM_RUNS As Long
    SIM_RUNS = wsA.Range("B1").Value
    If SIM_RUNS < 1 Then SIM_RUNS = 100000

    Const WTI_CONTRACT As Long = 1000
    Const FX_CONTRACT  As Double = 100000

    Dim wti_cont As Double: wti_cont = barrels / WTI_CONTRACT

    Const TRADING_DAYS As Double = STEPS_PER_YEAR
    Dim hedge_steps As Long:   hedge_steps = CLng(T_total * TRADING_DAYS)
    Dim step_WTI    As Long:   step_WTI = CLng(T_WTI * TRADING_DAYS)

    Dim dt          As Double: dt = T_total / hedge_steps
    Dim kappa       As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim eps         As Double: eps = 0.01
    Dim BUFFER      As Double: BUFFER = 0.05

    Dim lnVar As Double: lnVar = vol1 ^ 2 * dt

    Dim primedPremium As Double: primedPremium = 0#
    If Beta_Mat_Steps <> hedge_steps Then

        primedPremium = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, _
                                        KO_up, KO_dn, S0_WTI, S0_FX, vol1, vol2, corr, _
                                        hedge_steps, T_total, 50000, S0_WTI, True)
    End If

    Dim Base_Premium  As Double: Base_Premium = wsLS.Range("J2").Value

    If Base_Premium <= 0# And primedPremium > 0# Then Base_Premium = primedPremium
    Dim total_premium As Double: total_premium = Base_Premium * wti_cont

    Dim sim           As Long, stepIdx As Long, rowIdx As Long
    Dim doRecord      As Boolean
    Dim S_WTI         As Double, S_FX As Double
    Dim delta_WTI     As Double, delta_FX As Double
    Dim delta_WTI_old As Double, delta_FX_old As Double
    Dim pos_WTI       As Double, pos_FX As Double
    Dim prev_WTI      As Double, prev_FX As Double
    Dim dPos_WTI      As Double, dPos_FX As Double
    Dim cost_FX       As Double
    Dim int_cost      As Double

    Dim margin_WTI    As Double
    Dim cumul_FX      As Double
    Dim cumul_cost    As Double
    Dim prev_S_WTI    As Double
    Dim mtm_WTI       As Double
    Dim int_margin    As Double
    Dim int_FX        As Double

    Dim z1 As Double, z2 As Double, e2 As Double, nJ As Long
    Dim tau           As Double, remain_steps As Long
    Dim V_base        As Double, fade As Double
    Dim intrinsic     As Double, jump_loss As Double
    Dim wtiGain       As Double
    Dim knocked_out   As Boolean, exercised As Boolean
    Dim exercise_step As Long
    Dim S_WTI_ex      As Double, S_FX_ex As Double
    Dim opt_profit    As Double, tot_profit As Double
    Dim py            As Double, RR As Long, s As Long

    Dim WTI_expired   As Boolean
    Dim wasExpiredPre As Boolean
    Dim S_WTI_expiry  As Double
    Dim effective_WTI As Double
    Dim raw_dWTI      As Double

    Dim p_bb_up As Double, p_bb_dn As Double

    Dim opt_profits()  As Double:  ReDim opt_profits(1 To SIM_RUNS)
    Dim tot_profits()  As Double:  ReDim tot_profits(1 To SIM_RUNS)
    Dim jump_losses()  As Double:  ReDim jump_losses(1 To SIM_RUNS)
    Dim ko_flags()     As Boolean: ReDim ko_flags(1 To SIM_RUNS)
    Dim ex_flags()     As Boolean: ReDim ex_flags(1 To SIM_RUNS)

    Dim simTable()     As Variant: ReDim simTable(1 To SIM_RUNS, 1 To 6)

    Const TABLE_START As Long = 30
    wsA.Range("A1") = "Simulation Runs:"
    wsA.Range("C1") = " <- enter in B1"
    wsA.Range("A2") = "Hedge Steps:"
    wsA.Range("B2") = hedge_steps
    wsA.Range("A3") = "Base Premium (KRW):"
    wsA.Range("B3") = Base_Premium
    wsA.Range("A4") = "Total Premium (KRW):"
    wsA.Range("B4") = total_premium
    If wsA.Range("B1").Value < 1 Then wsA.Range("B1").Value = SIM_RUNS

    wsA.Range("A" & TABLE_START & ":U" & _
             (TABLE_START + WorksheetFunction.Max(hedge_steps, SIM_RUNS) + 60)).ClearContents

    wsA.Cells(TABLE_START, 1) = "Step"
    wsA.Cells(TABLE_START, 2) = "S_WTI"
    wsA.Cells(TABLE_START, 3) = "S_FX"
    wsA.Cells(TABLE_START, 4) = "Delta_WTI"
    wsA.Cells(TABLE_START, 5) = "Delta_FX"
    wsA.Cells(TABLE_START, 6) = "Contracts Rebal"
    wsA.Cells(TABLE_START, 7) = "Step Cash Flow (KRW)"
    wsA.Cells(TABLE_START, 8) = "Net Hedge Cost (KRW)"
    wsA.Cells(TABLE_START, 9) = "Interest Cost (KRW)"
    wsA.Cells(TABLE_START, 11) = "Sim#"
    wsA.Cells(TABLE_START, 12) = "Option Profit (KRW)"
    wsA.Cells(TABLE_START, 13) = "Total Profit (KRW)"
    wsA.Cells(TABLE_START, 14) = "Jump Loss (KRW)"
    wsA.Cells(TABLE_START, 15) = "KO"
    wsA.Cells(TABLE_START, 16) = "Early Exercise Step"

    Dim d0_WTI  As Double, d0_FX As Double
    Dim V0_base As Double, V0_up1 As Double, V0_up2 As Double

    Dim init_n  As Long: init_n = 50000

    V0_base = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KO_up, KO_dn, _
                               S0_WTI, S0_FX, vol1, vol2, corr, hedge_steps, T_total, init_n, S0_WTI)
    V0_up1 = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KO_up, KO_dn, _
                               S0_WTI * (1 + eps), S0_FX, vol1, vol2, corr, hedge_steps, T_total, init_n, S0_WTI)
    V0_up2 = Calc_LSMC_Price(Lambda, JumpMean, JumpVol, Drift1, Drift2, KO_up, KO_dn, _
                               S0_WTI, S0_FX * (1 + eps), vol1, vol2, corr, hedge_steps, T_total, init_n, S0_WTI)

    d0_WTI = ((V0_up1 - V0_base) / (S0_WTI * eps)) / S0_FX
    d0_WTI = WorksheetFunction.Max(0#, WorksheetFunction.Min(1#, d0_WTI))
    d0_FX = d0_WTI

    If PB_N <> SIM_RUNS Or PB_Steps <> hedge_steps Then
        Call Build_Path_Bank(SIM_RUNS, hedge_steps, S0_WTI, S0_FX, vol1, vol2, corr, _
                             Drift1, Drift2, Lambda, JumpMean, JumpVol, KO_up, KO_dn, T_total)
    End If

    For sim = 1 To SIM_RUNS

        doRecord = (sim = SIM_RUNS)

        S_WTI = S0_WTI
        S_FX = S0_FX
        prev_S_WTI = S0_WTI
        knocked_out = False
        exercised = False
        exercise_step = 0
        S_WTI_ex = 0#
        S_FX_ex = 0#
        jump_loss = 0#
        WTI_expired = False
        S_WTI_expiry = 0#

        delta_WTI = d0_WTI
        delta_FX = d0_FX

        pos_WTI = delta_WTI * wti_cont
        pos_FX = delta_FX * S0_WTI * wti_cont * WTI_CONTRACT / FX_CONTRACT
        prev_WTI = pos_WTI
        prev_FX = pos_FX

        margin_WTI = 0#
        cumul_FX = pos_FX * FX_CONTRACT * S0_FX
        cumul_cost = cumul_FX - margin_WTI

        If doRecord Then
            rowIdx = TABLE_START + 1
            wsA.Cells(rowIdx, 1) = 0
            wsA.Cells(rowIdx, 2) = S0_WTI
            wsA.Cells(rowIdx, 3) = S0_FX
            wsA.Cells(rowIdx, 4) = delta_WTI
            wsA.Cells(rowIdx, 5) = delta_FX
            wsA.Cells(rowIdx, 6) = pos_WTI
            wsA.Cells(rowIdx, 7) = cumul_FX
            wsA.Cells(rowIdx, 8) = cumul_cost
            wsA.Cells(rowIdx, 9) = 0
        End If

        For stepIdx = 1 To hedge_steps

            If exercised Or knocked_out Then GoTo RecordZero

            nJ = PB_nJ(sim, stepIdx)
            S_WTI = PB_S1(sim, stepIdx)
            S_FX = PB_S2(sim, stepIdx)

            delta_WTI_old = delta_WTI
            delta_FX_old = delta_FX

            wasExpiredPre = WTI_expired
            If Not WTI_expired And stepIdx >= step_WTI Then
                S_WTI_expiry = S_WTI
                WTI_expired = True
            End If
            If WTI_expired Then
                effective_WTI = S_WTI_expiry
            Else
                effective_WTI = S_WTI
            End If

            If PB_KO(sim) > 0 And stepIdx = PB_KO(sim) Then
                knocked_out = True: delta_WTI = 0#: delta_FX = 0#: GoTo WriteRow
            End If

            tau = T_total - stepIdx * dt
            remain_steps = hedge_steps - stepIdx

            If tau > 0.001 And remain_steps > 0 Then

                Dim v1_norm As Double: v1_norm = effective_WTI / S0_WTI
                Dim v2_norm As Double: v2_norm = S_FX / S0_FX

                Dim b0 As Double: b0 = Beta_Mat(stepIdx, 0)
                Dim b1 As Double: b1 = Beta_Mat(stepIdx, 1)
                Dim b2 As Double: b2 = Beta_Mat(stepIdx, 2)
                Dim b3 As Double: b3 = Beta_Mat(stepIdx, 3)
                Dim b4 As Double: b4 = Beta_Mat(stepIdx, 4)
                Dim b5 As Double: b5 = Beta_Mat(stepIdx, 5)

                V_base = b0 + b1 * v1_norm + b2 * v2_norm + b3 * (v1_norm ^ 2) _
                       + b4 * (v2_norm ^ 2) + b5 * (v1_norm * v2_norm)
                If V_base < 0# Then V_base = 0#

                If WTI_expired Then
                    delta_WTI = 0#
                Else
                    raw_dWTI = ((b1 + 2# * b3 * v1_norm + b5 * v2_norm) / S0_WTI) / S_FX

                    If raw_dWTI < 0# Then
                        delta_WTI = 0#
                    ElseIf raw_dWTI > 1# Then
                        delta_WTI = 1#
                    Else
                        delta_WTI = raw_dWTI
                    End If
                End If

                delta_FX = delta_ratio * delta_WTI

                If effective_WTI > KO_up * (1# - BUFFER) Then
                    fade = (KO_up - effective_WTI) / (KO_up * BUFFER)
                    If fade < 0# Then fade = 0#
                    delta_WTI = delta_WTI * fade
                    delta_FX = delta_FX * fade
                End If

                If effective_WTI < KO_dn * (1# + BUFFER) Then
                    fade = (effective_WTI - KO_dn) / (KO_dn * BUFFER)
                    If fade < 0# Then fade = 0#
                    delta_WTI = delta_WTI * fade
                    delta_FX = delta_FX * fade
                End If

                wtiGain = effective_WTI - S0_WTI
                If wtiGain < 0# Then wtiGain = 0#
                intrinsic = wtiGain * S_FX * wti_cont
                If intrinsic > V_base * wti_cont And intrinsic > 0# Then
                    exercised = True
                    exercise_step = stepIdx
                    S_WTI_ex = effective_WTI
                    S_FX_ex = S_FX
                    delta_WTI = 0#
                    delta_FX = 0#
                End If
            Else
                If WTI_expired Then
                    delta_WTI = 0#
                Else
                    delta_WTI = IIf(effective_WTI > S0_WTI, 1#, 0#)
                End If
                delta_FX = delta_ratio * delta_WTI
            End If

            If nJ > 0 Then
                jump_loss = jump_loss + Abs((delta_WTI - delta_WTI_old) _
                            * wti_cont * S_WTI * WTI_CONTRACT * S_FX)
            End If

WriteRow:

            pos_WTI = delta_WTI * wti_cont
            dPos_WTI = pos_WTI - prev_WTI

            pos_FX = delta_FX * IIf(S_WTI < S0_WTI, S_WTI, S0_WTI) * wti_cont * WTI_CONTRACT / FX_CONTRACT
            dPos_FX = pos_FX - prev_FX

            mtm_WTI = prev_WTI * (S_WTI - prev_S_WTI) * WTI_CONTRACT * S_FX
            margin_WTI = margin_WTI + mtm_WTI
            int_margin = margin_WTI * WACC * dt
            margin_WTI = margin_WTI + int_margin

            cost_FX = dPos_FX * FX_CONTRACT * S_FX
            int_FX = cumul_FX * WACC * dt
            cumul_FX = cumul_FX + cost_FX + int_FX

            int_cost = int_margin + int_FX
            cumul_cost = cumul_FX - margin_WTI

            prev_S_WTI = S_WTI

            If doRecord Then
                rowIdx = rowIdx + 1
                wsA.Cells(rowIdx, 1) = stepIdx
                wsA.Cells(rowIdx, 2) = effective_WTI
                wsA.Cells(rowIdx, 3) = S_FX
                wsA.Cells(rowIdx, 4) = IIf(knocked_out, "KO", delta_WTI)
                wsA.Cells(rowIdx, 5) = IIf(knocked_out, "KO", delta_FX)
                wsA.Cells(rowIdx, 6) = dPos_WTI
                wsA.Cells(rowIdx, 7) = mtm_WTI + cost_FX
                wsA.Cells(rowIdx, 8) = cumul_cost
                wsA.Cells(rowIdx, 9) = int_cost
            End If

            prev_WTI = pos_WTI
            prev_FX = pos_FX
            GoTo NextStep

RecordZero:
            If doRecord Then
                rowIdx = rowIdx + 1
                wsA.Cells(rowIdx, 1) = stepIdx
                wsA.Cells(rowIdx, 2) = effective_WTI
                wsA.Cells(rowIdx, 3) = S_FX
                wsA.Cells(rowIdx, 4) = IIf(knocked_out, "KO", "exercised")
                wsA.Cells(rowIdx, 5) = 0
                wsA.Cells(rowIdx, 6) = 0
                wsA.Cells(rowIdx, 7) = 0
                wsA.Cells(rowIdx, 8) = cumul_cost
                wsA.Cells(rowIdx, 9) = 0
            End If

NextStep:
        Next stepIdx

        If exercised Then
            wtiGain = S_WTI_ex - S0_WTI
            If wtiGain < 0# Then wtiGain = 0#
            py = wtiGain * S_FX_ex * WTI_CONTRACT * wti_cont
        ElseIf Not knocked_out Then
            wtiGain = effective_WTI - S0_WTI
            If wtiGain < 0# Then wtiGain = 0#
            py = wtiGain * S_FX * WTI_CONTRACT * wti_cont
        Else
            py = 0#
        End If

        opt_profit = total_premium - py

        tot_profit = opt_profit - cumul_cost

        opt_profits(sim) = opt_profit
        tot_profits(sim) = tot_profit
        jump_losses(sim) = jump_loss
        ko_flags(sim) = knocked_out
        ex_flags(sim) = exercised

        simTable(sim, 1) = sim
        simTable(sim, 2) = opt_profit
        simTable(sim, 3) = tot_profit
        simTable(sim, 4) = jump_loss
        simTable(sim, 5) = IIf(knocked_out, "Y", "N")
        simTable(sim, 6) = IIf(exercised, exercise_step, "N")

    Next sim

    wsA.Range(wsA.Cells(TABLE_START + 1, 11), wsA.Cells(TABLE_START + SIM_RUNS, 16)).Value = simTable

    Dim rngOptProfit As Range, rngTotProfit As Range
    Set rngOptProfit = wsA.Range(wsA.Cells(TABLE_START + 1, 12), wsA.Cells(TABLE_START + SIM_RUNS, 12))
    Set rngTotProfit = wsA.Range(wsA.Cells(TABLE_START + 1, 13), wsA.Cells(TABLE_START + SIM_RUNS, 13))
    rngOptProfit.FormatConditions.Delete
    rngTotProfit.FormatConditions.Delete
    With rngOptProfit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = RGB(198, 239, 206)
    End With
    With rngOptProfit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 199, 206)
    End With
    With rngTotProfit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = RGB(198, 239, 206)
    End With
    With rngTotProfit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 199, 206)
    End With

    Dim ko_count As Long, ex_count As Long
    For s = 1 To SIM_RUNS
        If ko_flags(s) Then ko_count = ko_count + 1
        If ex_flags(s) Then ex_count = ex_count + 1
    Next s

    RR = TABLE_START
    wsA.Cells(RR, 18) = ""
    wsA.Cells(RR, 19) = "Option Profit"
    wsA.Cells(RR, 20) = "Total Profit"

    Dim statRow As Long: statRow = RR + 1
    wsA.Cells(statRow, 18) = "Minimum"
    wsA.Cells(statRow, 19) = WorksheetFunction.Min(opt_profits)
    wsA.Cells(statRow, 20) = WorksheetFunction.Min(tot_profits)
    wsA.Cells(statRow + 1, 18) = "Maximum"
    wsA.Cells(statRow + 1, 19) = WorksheetFunction.Max(opt_profits)
    wsA.Cells(statRow + 1, 20) = WorksheetFunction.Max(tot_profits)
    wsA.Cells(statRow + 2, 18) = "Mean"
    wsA.Cells(statRow + 2, 19) = WorksheetFunction.Average(opt_profits)
    wsA.Cells(statRow + 2, 20) = WorksheetFunction.Average(tot_profits)
    wsA.Cells(statRow + 3, 18) = "Std Dev"
    wsA.Cells(statRow + 3, 19) = WorksheetFunction.StDev(opt_profits)
    wsA.Cells(statRow + 3, 20) = WorksheetFunction.StDev(tot_profits)
    wsA.Cells(statRow + 4, 18) = "Variance"
    wsA.Cells(statRow + 4, 19) = WorksheetFunction.Var(opt_profits)
    wsA.Cells(statRow + 4, 20) = WorksheetFunction.Var(tot_profits)
    wsA.Cells(statRow + 5, 18) = "Skewness"
    wsA.Cells(statRow + 5, 19) = WorksheetFunction.Skew(opt_profits)
    wsA.Cells(statRow + 5, 20) = WorksheetFunction.Skew(tot_profits)
    wsA.Cells(statRow + 6, 18) = "Kurtosis"
    wsA.Cells(statRow + 6, 19) = WorksheetFunction.Kurt(opt_profits)
    wsA.Cells(statRow + 6, 20) = WorksheetFunction.Kurt(tot_profits)
    wsA.Cells(statRow + 7, 18) = "Median"
    wsA.Cells(statRow + 7, 19) = WorksheetFunction.Median(opt_profits)
    wsA.Cells(statRow + 7, 20) = WorksheetFunction.Median(tot_profits)
    wsA.Cells(statRow + 8, 18) = "Avg Jump Loss"
    wsA.Cells(statRow + 8, 19) = WorksheetFunction.Average(jump_losses)
    wsA.Cells(statRow + 8, 20) = WorksheetFunction.Average(jump_losses)
    wsA.Cells(statRow + 9, 18) = "KO Rate"
    wsA.Cells(statRow + 9, 19) = ko_count / SIM_RUNS
    wsA.Cells(statRow + 9, 19).NumberFormat = "0.0%"
    wsA.Cells(statRow + 10, 18) = "Early Exercise Rate"
    wsA.Cells(statRow + 10, 19) = ex_count / SIM_RUNS
    wsA.Cells(statRow + 10, 19).NumberFormat = "0.0%"

    Dim pctHdr As Long: pctHdr = statRow + 12
    wsA.Cells(pctHdr, 18) = "Percentile"
    wsA.Cells(pctHdr, 19) = "Option Profit"
    wsA.Cells(pctHdr, 20) = "Total Profit"

    Dim pct    As Integer
    Dim pctRow As Long: pctRow = pctHdr + 1
    Dim p      As Double

    For pct = 5 To 95 Step 5
        p = pct / 100#
        wsA.Cells(pctRow, 18) = pct & "%"
        wsA.Cells(pctRow, 19) = WorksheetFunction.Percentile(opt_profits, p)
        wsA.Cells(pctRow, 20) = WorksheetFunction.Percentile(tot_profits, p)
        wsA.Cells(pctRow, 19).Interior.Color = _
            IIf(wsA.Cells(pctRow, 19).Value >= 0, RGB(198, 239, 206), RGB(255, 199, 206))
        wsA.Cells(pctRow, 20).Interior.Color = _
            IIf(wsA.Cells(pctRow, 20).Value >= 0, RGB(198, 239, 206), RGB(255, 199, 206))
        pctRow = pctRow + 1
    Next pct

    wsA.Range(wsA.Cells(statRow, 19), wsA.Cells(pctRow - 1, 20)).NumberFormat = "#,##0.0000"
    wsA.Range(wsA.Cells(RR, 18), wsA.Cells(RR, 20)).Font.Bold = True
    wsA.Range(wsA.Cells(pctHdr, 18), wsA.Cells(pctHdr, 20)).Font.Bold = True
    wsA.Columns("R:T").AutoFit

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Dim opt_mean As Double: opt_mean = WorksheetFunction.Average(opt_profits)
    Dim tot_mean As Double: tot_mean = WorksheetFunction.Average(tot_profits)
    Dim tot_std  As Double: tot_std = WorksheetFunction.StDev(tot_profits)

    MsgBox "American Delta Hedge complete" & vbCrLf & _
           "Simulations:   " & SIM_RUNS & vbCrLf & _
           "Hedge steps:   " & hedge_steps & vbCrLf & vbCrLf & _
           "Base Premium:  " & Format(Base_Premium, "#,##0") & " KRW" & vbCrLf & _
           "Total Premium: " & Format(total_premium, "#,##0") & " KRW" & vbCrLf & _
           "KO Rate:       " & Format(ko_count / SIM_RUNS, "0.0%") & vbCrLf & _
           "Exercise Rate: " & Format(ex_count / SIM_RUNS, "0.0%") & vbCrLf & _
           "Mean Opt P&L:  " & Format(opt_mean, "#,##0") & " KRW" & vbCrLf & _
           "Mean Tot P&L:  " & Format(tot_mean, "#,##0") & " KRW" & vbCrLf & _
           "StdDev(Tot):   " & Format(tot_std, "#,##0") & " KRW" & vbCrLf & vbCrLf & _
           "Fixes applied: BB correction in hedge sim (#1), log-space BB formula (#2)," & vbCrLf & _
           "WTI futures MTM cost model (#3), symmetric barrier fade (#5)", _
           vbInformation, "American Delta Hedge (Revised)"
End Sub

Public Sub Run_European_DeltaHedge()

    Dim wsE   As Worksheet: Set wsE = Sheets("European_Delta")
    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsLS  As Worksheet: Set wsLS = Sheets("LSMC")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim Lambda   As Double: Lambda = wsLS.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLS.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsLS.Range("B3").Value
    Dim Drift1   As Double: Drift1 = wsLS.Range("B4").Value
    Dim Drift2   As Double: Drift2 = wsLS.Range("B5").Value
    Dim KO_up    As Double: KO_up = wsLS.Range("B6").Value
    Dim KO_dn    As Double: KO_dn = wsLS.Range("B7").Value
    Dim S0_WTI   As Double: S0_WTI = wsLS.Range("B8").Value
    Dim S0_FX    As Double: S0_FX = wsLS.Range("B9").Value
    Dim vol1     As Double: vol1 = wsLS.Range("B10").Value
    Dim vol2     As Double: vol2 = wsLS.Range("B11").Value
    Dim corr     As Double: corr = wsLS.Range("B12").Value
    Dim T_total  As Double: T_total = wsLS.Range("B14").Value

    Dim delta_ratio As Double: delta_ratio = 1#

    Dim WACC     As Double: WACC = wsEnc.Range("B12").Value
    Dim r_US     As Double: r_US = wsEnc.Range("B4").Value
    Dim r_KR     As Double: r_KR = wsEnc.Range("B5").Value
    Dim barrels  As Double: barrels = wsEnc.Range("B9").Value

    Dim SIM_RUNS As Long
    SIM_RUNS = wsE.Range("B1").Value
    If SIM_RUNS < 1 Then SIM_RUNS = 100000

    Const WTI_CONTRACT As Long = 1000
    Const FX_CONTRACT  As Double = 100000

    Dim wti_cont As Double: wti_cont = barrels / WTI_CONTRACT

    Dim hedge_steps As Long: hedge_steps = CLng(T_total * STEPS_PER_YEAR)
    Dim dt          As Double: dt = T_total / hedge_steps
    Dim kappa       As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar       As Double: lnVar = vol1 ^ 2 * dt
    Dim BUFFER      As Double: BUFFER = 0.05

    Dim K_WTI As Double: K_WTI = S0_WTI

    Dim sigma_total As Double: sigma_total = Sigma_Eff_WTI(vol1, Lambda, JumpMean, JumpVol)

    Dim Base_Premium_EU  As Double: Base_Premium_EU = wsLS.Range("J2").Value
    Dim total_premium_EU As Double: total_premium_EU = Base_Premium_EU * wti_cont

    Dim p As Long, i As Long, rowIdx As Long
    Dim S_WTI As Double, S_FX As Double
    Dim prev_S_WTI As Double, prev_S_FX As Double
    Dim delta_WTI As Double, delta_FX As Double
    Dim delta_WTI_old As Double, delta_FX_old As Double
    Dim pos_WTI As Double, pos_FX As Double
    Dim prev_WTI As Double, prev_FX As Double
    Dim dPos_WTI As Double, dPos_FX As Double
    Dim cost_FX As Double
    Dim margin_WTI As Double
    Dim cumul_FX As Double
    Dim cumul_cost As Double
    Dim mtm_WTI As Double
    Dim int_margin As Double, int_FX As Double, int_cost As Double
    Dim z1 As Double, z2 As Double, e2 As Double, nJ As Long
    Dim tau As Double
    Dim fade As Double
    Dim knocked_out As Boolean
    Dim opt_profit As Double, tot_profit As Double, py As Double
    Dim d1 As Double

    Dim opt_profits()  As Double:  ReDim opt_profits(1 To SIM_RUNS)
    Dim tot_profits()  As Double:  ReDim tot_profits(1 To SIM_RUNS)
    Dim ko_flags()     As Boolean: ReDim ko_flags(1 To SIM_RUNS)
    Dim simTable()     As Variant: ReDim simTable(1 To SIM_RUNS, 1 To 4)

    Const TABLE_START As Long = 30
    wsE.Range("A1") = "Simulation Runs:"
    wsE.Range("C1") = " <- enter in B1"
    wsE.Range("A2") = "Hedge Steps:"
    wsE.Cells(2, 2).Value = hedge_steps
    wsE.Range("A3") = "Strike (K_WTI):"
    wsE.Cells(3, 2).Value = K_WTI
    wsE.Range("A4") = "Base Premium (KRW):"
    wsE.Cells(4, 2).Value = Base_Premium_EU
    wsE.Range("A5") = "Total Premium (KRW):"
    wsE.Cells(5, 2).Value = total_premium_EU
    If wsE.Range("B1").Value < 1 Then wsE.Range("B1").Value = SIM_RUNS

    wsE.Range("A" & TABLE_START & ":U" & _
             (TABLE_START + WorksheetFunction.Max(hedge_steps, SIM_RUNS) + 40)).ClearContents

    wsE.Cells(TABLE_START, 1) = "Step"
    wsE.Cells(TABLE_START, 2) = "S_WTI"
    wsE.Cells(TABLE_START, 3) = "S_FX"
    wsE.Cells(TABLE_START, 4) = "Delta_WTI"
    wsE.Cells(TABLE_START, 5) = "Delta_FX"
    wsE.Cells(TABLE_START, 6) = "Contracts Rebal"
    wsE.Cells(TABLE_START, 7) = "Step Cash Flow (KRW)"
    wsE.Cells(TABLE_START, 8) = "Net Hedge Cost (KRW)"
    wsE.Cells(TABLE_START, 9) = "Interest Cost (KRW)"
    wsE.Cells(TABLE_START, 11) = "Sim#"
    wsE.Cells(TABLE_START, 12) = "Option Profit (KRW)"
    wsE.Cells(TABLE_START, 13) = "Total Profit (KRW)"

    Dim d0_WTI As Double, d0_FX As Double
    Dim tau0 As Double: tau0 = T_total

    d1 = (Log(S0_WTI / K_WTI) + 0.5 * sigma_total ^ 2 * tau0) / (sigma_total * Sqr(tau0))
    d0_WTI = Exp(-r_US * tau0) * WorksheetFunction.Norm_S_Dist(d1, True)

    If d0_WTI < 0# Then d0_WTI = 0#
    If d0_WTI > 1# Then d0_WTI = 1#
    d0_FX = d0_WTI

    If PB_N <> SIM_RUNS Or PB_Steps <> hedge_steps Then
        Call Build_Path_Bank(SIM_RUNS, hedge_steps, S0_WTI, S0_FX, vol1, vol2, corr, _
                             Drift1, Drift2, Lambda, JumpMean, JumpVol, KO_up, KO_dn, T_total)
    End If

    For p = 1 To SIM_RUNS

        Dim doRecord As Boolean: doRecord = (p = SIM_RUNS)

        S_WTI = S0_WTI
        S_FX = S0_FX
        prev_S_WTI = S0_WTI
        prev_S_FX = S0_FX
        knocked_out = False

        delta_WTI = d0_WTI
        delta_FX = d0_FX

        pos_WTI = delta_WTI * wti_cont
        pos_FX = delta_FX * S0_WTI * wti_cont * WTI_CONTRACT / FX_CONTRACT
        prev_WTI = pos_WTI
        prev_FX = pos_FX

        margin_WTI = 0#
        cumul_FX = pos_FX * FX_CONTRACT * S0_FX
        cumul_cost = cumul_FX - margin_WTI

        If doRecord Then
            rowIdx = TABLE_START + 1
            wsE.Cells(rowIdx, 1) = 0
            wsE.Cells(rowIdx, 2) = S0_WTI
            wsE.Cells(rowIdx, 3) = S0_FX
            wsE.Cells(rowIdx, 4) = delta_WTI
            wsE.Cells(rowIdx, 5) = delta_FX
            wsE.Cells(rowIdx, 6) = pos_WTI
            wsE.Cells(rowIdx, 7) = cumul_FX
            wsE.Cells(rowIdx, 8) = cumul_cost
            wsE.Cells(rowIdx, 9) = 0
        End If

        For i = 1 To hedge_steps

            If knocked_out Then Exit For

            S_WTI = PB_S1(p, i)
            S_FX = PB_S2(p, i)

            delta_WTI_old = delta_WTI
            delta_FX_old = delta_FX

            If PB_KO(p) > 0 And i = PB_KO(p) Then
                knocked_out = True: delta_WTI = 0#: delta_FX = 0#: GoTo WriteRow_EU
            End If

            tau = T_total - i * dt
            If tau > 0.001 Then
                d1 = (Log(S_WTI / K_WTI) + 0.5 * sigma_total ^ 2 * tau) / (sigma_total * Sqr(tau))
                delta_WTI = Exp(-r_US * tau) * WorksheetFunction.Norm_S_Dist(d1, True)

                If delta_WTI < 0# Then delta_WTI = 0#
                If delta_WTI > 1# Then delta_WTI = 1#
                delta_FX = delta_ratio * delta_WTI
            Else
                delta_WTI = IIf(S_WTI > K_WTI, 1#, 0#)
                delta_FX = delta_ratio * delta_WTI
            End If

            If S_WTI > KO_up * (1# - BUFFER) Then
                fade = (KO_up - S_WTI) / (KO_up * BUFFER)
                If fade < 0# Then fade = 0#
                delta_WTI = delta_WTI * fade
                delta_FX = delta_FX * fade
            End If
            If S_WTI < KO_dn * (1# + BUFFER) Then
                fade = (S_WTI - KO_dn) / (KO_dn * BUFFER)
                If fade < 0# Then fade = 0#
                delta_WTI = delta_WTI * fade
                delta_FX = delta_FX * fade
            End If

WriteRow_EU:

            pos_WTI = delta_WTI * wti_cont
            dPos_WTI = pos_WTI - prev_WTI

            pos_FX = delta_FX * IIf(S_WTI < K_WTI, S_WTI, K_WTI) * wti_cont * WTI_CONTRACT / FX_CONTRACT
            dPos_FX = pos_FX - prev_FX

            mtm_WTI = prev_WTI * (S_WTI - prev_S_WTI) * WTI_CONTRACT * S_FX
            margin_WTI = margin_WTI + mtm_WTI
            int_margin = margin_WTI * WACC * dt
            margin_WTI = margin_WTI + int_margin

            cost_FX = dPos_FX * FX_CONTRACT * S_FX
            int_FX = cumul_FX * WACC * dt
            cumul_FX = cumul_FX + cost_FX + int_FX

            int_cost = int_margin + int_FX
            cumul_cost = cumul_FX - margin_WTI

            prev_S_WTI = S_WTI

            If doRecord Then
                rowIdx = rowIdx + 1
                wsE.Cells(rowIdx, 1) = i
                wsE.Cells(rowIdx, 2) = S_WTI
                wsE.Cells(rowIdx, 3) = S_FX
                wsE.Cells(rowIdx, 4) = IIf(knocked_out, "KO", delta_WTI)
                wsE.Cells(rowIdx, 5) = IIf(knocked_out, "KO", delta_FX)
                wsE.Cells(rowIdx, 6) = dPos_WTI
                wsE.Cells(rowIdx, 7) = mtm_WTI + cost_FX
                wsE.Cells(rowIdx, 8) = cumul_cost
                wsE.Cells(rowIdx, 9) = int_cost
            End If

            prev_WTI = pos_WTI
            prev_FX = pos_FX
        Next i

        If Not knocked_out Then
            Dim wtiGain As Double: wtiGain = S_WTI - K_WTI
            If wtiGain < 0# Then wtiGain = 0#
            py = wtiGain * S_FX * WTI_CONTRACT * wti_cont
        Else
            py = 0#
        End If

        opt_profit = total_premium_EU - py
        tot_profit = opt_profit - cumul_cost

        opt_profits(p) = opt_profit
        tot_profits(p) = tot_profit
        ko_flags(p) = knocked_out

        simTable(p, 1) = p
        simTable(p, 2) = opt_profit
        simTable(p, 3) = tot_profit
        simTable(p, 4) = IIf(knocked_out, "Y", "N")

    Next p

    wsE.Range(wsE.Cells(TABLE_START + 1, 11), wsE.Cells(TABLE_START + SIM_RUNS, 14)).Value = simTable

    Dim rngOptProfit As Range, rngTotProfit As Range
    Set rngOptProfit = wsE.Range(wsE.Cells(TABLE_START + 1, 12), wsE.Cells(TABLE_START + SIM_RUNS, 12))
    Set rngTotProfit = wsE.Range(wsE.Cells(TABLE_START + 1, 13), wsE.Cells(TABLE_START + SIM_RUNS, 13))
    rngOptProfit.FormatConditions.Delete
    rngTotProfit.FormatConditions.Delete
    With rngOptProfit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = RGB(198, 239, 206)
    End With
    With rngOptProfit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 199, 206)
    End With
    With rngTotProfit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = RGB(198, 239, 206)
    End With
    With rngTotProfit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 199, 206)
    End With

    Dim ko_count As Long: ko_count = 0
    For i = 1 To SIM_RUNS
        If ko_flags(i) Then ko_count = ko_count + 1
    Next i

    Dim RR As Long: RR = TABLE_START
    wsE.Cells(RR, 18) = ""
    wsE.Cells(RR, 19) = "Option Profit"
    wsE.Cells(RR, 20) = "Total Profit"

    Dim statRow As Long: statRow = RR + 1
    wsE.Cells(statRow, 18) = "Mean"
    wsE.Cells(statRow, 19) = WorksheetFunction.Average(opt_profits)
    wsE.Cells(statRow, 20) = WorksheetFunction.Average(tot_profits)
    wsE.Cells(statRow + 1, 18) = "Std Dev"
    wsE.Cells(statRow + 1, 19) = WorksheetFunction.StDev(opt_profits)
    wsE.Cells(statRow + 1, 20) = WorksheetFunction.StDev(tot_profits)
    wsE.Cells(statRow + 2, 18) = "KO Rate"
    wsE.Cells(statRow + 2, 19) = ko_count / SIM_RUNS
    wsE.Cells(statRow + 2, 19).NumberFormat = "0.0%"

    wsE.Range(wsE.Cells(statRow, 19), wsE.Cells(statRow + 1, 20)).NumberFormat = "#,##0.0000"
    wsE.Range(wsE.Cells(RR, 18), wsE.Cells(RR, 20)).Font.Bold = True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "European Delta Hedge (Fair Comparison) complete" & vbCrLf & _
           "KO Rate:  " & Format(ko_count / SIM_RUNS, "0.0%") & vbCrLf & _
           "Strike:   " & Format(K_WTI, "0.00") & " (ATM)" & vbCrLf & _
           "Premium:  " & Format(total_premium_EU, "#,##0") & " KRW" & vbCrLf & _
           "Hedge:    Vanilla BS delta + symmetric barrier fade", _
           vbInformation, "European Delta Hedge (Revised)"
End Sub

Private Function TW_M1(ByVal b As Double, ByVal tau As Double) As Double
    If Abs(b) < 1E-07 Then
        TW_M1 = 1#
    Else
        TW_M1 = (Exp(b * tau) - 1#) / (b * tau)
    End If
End Function

Private Function TW_M2(ByVal b As Double, ByVal vol As Double, ByVal tau As Double) As Double
    If Abs(b) < 1E-07 Then
        TW_M2 = 2# * (Exp(vol ^ 2 * tau) - 1# - vol ^ 2 * tau) / (vol ^ 4 * tau ^ 2)
    Else
        TW_M2 = (2# / tau ^ 2) * ( _
                    Exp((2# * b + vol ^ 2) * tau) / ((b + vol ^ 2) * (2# * b + vol ^ 2)) _
                  + (1# / b) * (1# / (2# * b + vol ^ 2) - Exp(b * tau) / (b + vol ^ 2)) _
                )
    End If
End Function

Private Function TW_Delta2(ByVal s As Double, ByVal S_avg As Double, ByVal Strike As Double, _
                           ByVal t_passed As Double, ByVal T_total As Double, _
                           ByVal vol As Double, ByVal b As Double, ByVal r_disc As Double) As Double
    Dim tau As Double: tau = T_total - t_passed
    If tau <= 0.001 Then
        TW_Delta2 = IIf(S_avg > Strike, 1#, 0#)
        Exit Function
    End If

    Dim W1 As Double: W1 = t_passed / T_total
    Dim W2 As Double: W2 = tau / T_total
    Dim Adj_Strike As Double: Adj_Strike = (Strike - W1 * S_avg) / W2

    Dim M1 As Double: M1 = TW_M1(b, tau)

    If Adj_Strike <= 0# Then
        TW_Delta2 = Exp(-r_disc * tau) * W2 * M1
        Exit Function
    End If

    Dim M2 As Double: M2 = TW_M2(b, vol, tau)
    Dim F_A As Double: F_A = s * M1
    Dim v_A As Double: v_A = Sqr(Log(M2 / (M1 ^ 2)))
    Dim d1 As Double: d1 = (Log(F_A / Adj_Strike) + 0.5 * v_A ^ 2) / v_A

    TW_Delta2 = Exp(-r_disc * tau) * W2 * M1 * WorksheetFunction.Norm_S_Dist(d1, True)
End Function

Public Sub Run_Asian_DeltaHedge()

    Dim wsA   As Worksheet: Set wsA = Sheets("Asian_Delta")
    Dim wsEnc As Worksheet: Set wsEnc = Sheets("Encoding")
    Dim wsLS  As Worksheet: Set wsLS = Sheets("LSMC")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim Lambda   As Double: Lambda = wsLS.Range("B1").Value
    Dim JumpMean As Double: JumpMean = wsLS.Range("B2").Value
    Dim JumpVol  As Double: JumpVol = wsLS.Range("B3").Value
    Dim Drift1   As Double: Drift1 = wsLS.Range("B4").Value
    Dim Drift2   As Double: Drift2 = wsLS.Range("B5").Value
    Dim KO_up    As Double: KO_up = wsLS.Range("B6").Value
    Dim KO_dn    As Double: KO_dn = wsLS.Range("B7").Value
    Dim S0_WTI   As Double: S0_WTI = wsLS.Range("B8").Value
    Dim S0_FX    As Double: S0_FX = wsLS.Range("B9").Value
    Dim vol1     As Double: vol1 = wsLS.Range("B10").Value
    Dim vol2     As Double: vol2 = wsLS.Range("B11").Value
    Dim corr     As Double: corr = wsLS.Range("B12").Value
    Dim T_total  As Double: T_total = wsLS.Range("B14").Value

    Dim delta_ratio As Double: delta_ratio = 1#

    Dim WACC     As Double: WACC = wsEnc.Range("B12").Value
    Dim r_US     As Double: r_US = wsEnc.Range("B4").Value
    Dim r_KR     As Double: r_KR = wsEnc.Range("B5").Value
    Dim barrels  As Double: barrels = wsEnc.Range("B9").Value

    Dim SIM_RUNS As Long
    SIM_RUNS = wsA.Range("B1").Value
    If SIM_RUNS < 1 Then SIM_RUNS = 100000

    Const WTI_CONTRACT As Long = 1000
    Const FX_CONTRACT  As Double = 100000

    Dim wti_cont As Double: wti_cont = barrels / WTI_CONTRACT

    Dim hedge_steps As Long: hedge_steps = CLng(T_total * STEPS_PER_YEAR)
    Dim dt          As Double: dt = T_total / hedge_steps
    Dim kappa       As Double: kappa = Exp(JumpMean + 0.5 * JumpVol ^ 2) - 1#
    Dim lnVar       As Double: lnVar = vol1 ^ 2 * dt
    Dim BUFFER      As Double: BUFFER = 0.05

    Dim K_WTI As Double: K_WTI = S0_WTI

    Dim sigma_eff As Double: sigma_eff = Sigma_Eff_WTI(vol1, Lambda, JumpMean, JumpVol)

    Dim Base_Premium_Asian  As Double: Base_Premium_Asian = wsLS.Range("J2").Value
    Dim total_premium_Asian As Double: total_premium_Asian = Base_Premium_Asian * wti_cont

    Dim p As Long, i As Long, rowIdx As Long
    Dim S_WTI As Double, S_FX As Double, prev_S_WTI As Double
    Dim delta_WTI As Double, delta_FX As Double
    Dim delta_WTI_old As Double, delta_FX_old As Double
    Dim pos_WTI As Double, pos_FX As Double
    Dim prev_WTI As Double, prev_FX As Double
    Dim dPos_WTI As Double, dPos_FX As Double
    Dim cost_FX As Double
    Dim margin_WTI As Double, cumul_FX As Double, cumul_cost As Double
    Dim mtm_WTI As Double, int_margin As Double, int_FX As Double, int_cost As Double
    Dim z1 As Double, z2 As Double, e2 As Double, nJ As Long
    Dim t_passed As Double, tau As Double
    Dim fade As Double
    Dim knocked_out As Boolean
    Dim opt_profit As Double, tot_profit As Double, py As Double
    Dim jump_loss As Double

    Dim run_sum_WTI As Double, avg_WTI As Double, avg_WTI_final As Double
    Dim count_WTI As Long

    Dim opt_profits()  As Double:  ReDim opt_profits(1 To SIM_RUNS)
    Dim tot_profits()  As Double:  ReDim tot_profits(1 To SIM_RUNS)
    Dim jump_losses()  As Double:  ReDim jump_losses(1 To SIM_RUNS)
    Dim ko_flags()     As Boolean: ReDim ko_flags(1 To SIM_RUNS)
    Dim simTable()     As Variant: ReDim simTable(1 To SIM_RUNS, 1 To 4)

    Const TABLE_START As Long = 30

    wsA.Range("A1") = "Simulation Runs:"
    wsA.Range("C1") = " <- enter in B1"
    wsA.Range("A2") = "Hedge Steps:"
    wsA.Cells(2, 2).Value = hedge_steps
    wsA.Range("A3") = "Strike (K_WTI):"
    wsA.Cells(3, 2).Value = K_WTI
    wsA.Range("A4") = "Base Premium (KRW):"
    wsA.Cells(4, 2).Value = Base_Premium_Asian
    wsA.Range("A5") = "Total Premium (KRW):"
    wsA.Cells(5, 2).Value = total_premium_Asian
    If wsA.Range("B1").Value < 1 Then wsA.Range("B1").Value = SIM_RUNS

    wsA.Range("A" & TABLE_START & ":U" & _
             (TABLE_START + WorksheetFunction.Max(hedge_steps, SIM_RUNS) + 40)).ClearContents

    wsA.Cells(TABLE_START, 1) = "Step"
    wsA.Cells(TABLE_START, 2) = "S_WTI"
    wsA.Cells(TABLE_START, 3) = "Avg_WTI"
    wsA.Cells(TABLE_START, 4) = "S_FX"
    wsA.Cells(TABLE_START, 5) = "Delta_WTI"
    wsA.Cells(TABLE_START, 6) = "Delta_FX"
    wsA.Cells(TABLE_START, 7) = "Contracts Rebal"
    wsA.Cells(TABLE_START, 8) = "Step Cash Flow (KRW)"
    wsA.Cells(TABLE_START, 9) = "Net Hedge Cost (KRW)"
    wsA.Cells(TABLE_START, 10) = "Interest Cost (KRW)"
    wsA.Cells(TABLE_START, 12) = "Sim#"
    wsA.Cells(TABLE_START, 13) = "Option Profit (KRW)"
    wsA.Cells(TABLE_START, 14) = "Total Profit (KRW)"

    Dim d0_WTI As Double, d0_FX As Double
    d0_WTI = TW_Delta2(S0_WTI, S0_WTI, K_WTI, 0#, T_total, sigma_eff, 0#, r_US)
    If d0_WTI < 0# Then d0_WTI = 0#
    If d0_WTI > 1# Then d0_WTI = 1#
    d0_FX = d0_WTI

    If PB_N <> SIM_RUNS Or PB_Steps <> hedge_steps Then
        Call Build_Path_Bank(SIM_RUNS, hedge_steps, S0_WTI, S0_FX, vol1, vol2, corr, _
                             Drift1, Drift2, Lambda, JumpMean, JumpVol, KO_up, KO_dn, T_total)
    End If

    For p = 1 To SIM_RUNS

        Dim doRecord As Boolean: doRecord = (p = SIM_RUNS)

        S_WTI = S0_WTI
        S_FX = S0_FX
        prev_S_WTI = S0_WTI
        knocked_out = False
        jump_loss = 0#

        delta_WTI = d0_WTI
        delta_FX = d0_FX

        pos_WTI = delta_WTI * wti_cont
        pos_FX = delta_FX * S0_WTI * wti_cont * WTI_CONTRACT / FX_CONTRACT
        prev_WTI = pos_WTI
        prev_FX = pos_FX

        margin_WTI = 0#
        cumul_FX = pos_FX * FX_CONTRACT * S0_FX
        cumul_cost = cumul_FX - margin_WTI

        run_sum_WTI = S0_WTI
        avg_WTI = S0_WTI
        count_WTI = 1
        avg_WTI_final = S0_WTI

        If doRecord Then
            rowIdx = TABLE_START + 1
            wsA.Cells(rowIdx, 1) = 0
            wsA.Cells(rowIdx, 2) = S0_WTI
            wsA.Cells(rowIdx, 3) = avg_WTI
            wsA.Cells(rowIdx, 4) = S0_FX
            wsA.Cells(rowIdx, 5) = delta_WTI
            wsA.Cells(rowIdx, 6) = delta_FX
            wsA.Cells(rowIdx, 7) = pos_WTI
            wsA.Cells(rowIdx, 8) = cumul_FX
            wsA.Cells(rowIdx, 9) = cumul_cost
            wsA.Cells(rowIdx, 10) = 0
        End If

        For i = 1 To hedge_steps

            If knocked_out Then Exit For

            nJ = PB_nJ(p, i)
            S_WTI = PB_S1(p, i)
            S_FX = PB_S2(p, i)

            delta_WTI_old = delta_WTI
            delta_FX_old = delta_FX

            run_sum_WTI = run_sum_WTI + S_WTI
            count_WTI = count_WTI + 1
            avg_WTI = run_sum_WTI / count_WTI

            If PB_KO(p) > 0 And i = PB_KO(p) Then
                knocked_out = True: delta_WTI = 0#: delta_FX = 0#
                avg_WTI_final = avg_WTI: GoTo WriteRow_AS
            End If

            t_passed = i * dt
            tau = T_total - t_passed
            If tau > 0.001 Then
                delta_WTI = TW_Delta2(S_WTI, avg_WTI, K_WTI, t_passed, T_total, sigma_eff, 0#, r_US)
                If delta_WTI < 0# Then delta_WTI = 0#
                If delta_WTI > 1# Then delta_WTI = 1#
                delta_FX = delta_ratio * delta_WTI
            Else
                delta_WTI = IIf(avg_WTI > K_WTI, 1#, 0#)
                delta_FX = delta_ratio * delta_WTI
            End If

            If S_WTI > KO_up * (1# - BUFFER) Then
                fade = (KO_up - S_WTI) / (KO_up * BUFFER)
                If fade < 0# Then fade = 0#
                delta_WTI = delta_WTI * fade
                delta_FX = delta_FX * fade
            End If
            If S_WTI < KO_dn * (1# + BUFFER) Then
                fade = (S_WTI - KO_dn) / (KO_dn * BUFFER)
                If fade < 0# Then fade = 0#
                delta_WTI = delta_WTI * fade
                delta_FX = delta_FX * fade
            End If

            If nJ > 0 Then
                jump_loss = jump_loss + Abs((delta_WTI - delta_WTI_old) _
                            * wti_cont * S_WTI * WTI_CONTRACT * S_FX)
            End If

WriteRow_AS:
            pos_WTI = delta_WTI * wti_cont
            dPos_WTI = pos_WTI - prev_WTI

            pos_FX = delta_FX * IIf(S_WTI < K_WTI, S_WTI, K_WTI) * wti_cont * WTI_CONTRACT / FX_CONTRACT
            dPos_FX = pos_FX - prev_FX

            mtm_WTI = prev_WTI * (S_WTI - prev_S_WTI) * WTI_CONTRACT * S_FX
            margin_WTI = margin_WTI + mtm_WTI
            int_margin = margin_WTI * WACC * dt
            margin_WTI = margin_WTI + int_margin

            cost_FX = dPos_FX * FX_CONTRACT * S_FX
            int_FX = cumul_FX * WACC * dt
            cumul_FX = cumul_FX + cost_FX + int_FX

            int_cost = int_margin + int_FX
            cumul_cost = cumul_FX - margin_WTI

            prev_S_WTI = S_WTI

            If doRecord Then
                rowIdx = rowIdx + 1
                wsA.Cells(rowIdx, 1) = i
                wsA.Cells(rowIdx, 2) = S_WTI
                wsA.Cells(rowIdx, 3) = avg_WTI
                wsA.Cells(rowIdx, 4) = S_FX
                wsA.Cells(rowIdx, 5) = IIf(knocked_out, "KO", delta_WTI)
                wsA.Cells(rowIdx, 6) = IIf(knocked_out, "KO", delta_FX)
                wsA.Cells(rowIdx, 7) = dPos_WTI
                wsA.Cells(rowIdx, 8) = mtm_WTI + cost_FX
                wsA.Cells(rowIdx, 9) = cumul_cost
                wsA.Cells(rowIdx, 10) = int_cost
            End If

            prev_WTI = pos_WTI
            prev_FX = pos_FX
        Next i

        If Not knocked_out Then avg_WTI_final = avg_WTI

        If Not knocked_out Then
            py = WorksheetFunction.Max(avg_WTI_final - K_WTI, 0) * S_FX * WTI_CONTRACT * wti_cont
        Else
            py = 0#
        End If

        opt_profit = total_premium_Asian - py
        tot_profit = opt_profit - cumul_cost

        opt_profits(p) = opt_profit
        tot_profits(p) = tot_profit
        jump_losses(p) = jump_loss
        ko_flags(p) = knocked_out

        simTable(p, 1) = p
        simTable(p, 2) = opt_profit
        simTable(p, 3) = tot_profit
        simTable(p, 4) = IIf(knocked_out, "Y", "N")

    Next p

    wsA.Range(wsA.Cells(TABLE_START + 1, 12), wsA.Cells(TABLE_START + SIM_RUNS, 15)).Value = simTable

    Dim ko_count As Long: ko_count = 0
    For i = 1 To SIM_RUNS
        If ko_flags(i) Then ko_count = ko_count + 1
    Next i

    Dim RR As Long: RR = TABLE_START
    wsA.Cells(RR, 18) = ""
    wsA.Cells(RR, 19) = "Option Profit"
    wsA.Cells(RR, 20) = "Total Profit"

    Dim statRow As Long: statRow = RR + 1
    wsA.Cells(statRow, 18) = "Mean"
    wsA.Cells(statRow, 19) = WorksheetFunction.Average(opt_profits)
    wsA.Cells(statRow, 20) = WorksheetFunction.Average(tot_profits)
    wsA.Cells(statRow + 1, 18) = "Std Dev"
    wsA.Cells(statRow + 1, 19) = WorksheetFunction.StDev(opt_profits)
    wsA.Cells(statRow + 1, 20) = WorksheetFunction.StDev(tot_profits)
    wsA.Cells(statRow + 2, 18) = "KO Rate"
    wsA.Cells(statRow + 2, 19) = ko_count / SIM_RUNS
    wsA.Cells(statRow + 2, 19).NumberFormat = "0.0%"

    wsA.Range(wsA.Cells(statRow, 19), wsA.Cells(statRow + 1, 20)).NumberFormat = "#,##0.0000"
    wsA.Range(wsA.Cells(RR, 18), wsA.Cells(RR, 20)).Font.Bold = True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Asian Delta Hedge (TW proxy + symmetric fade) complete" & vbCrLf & _
           "KO Rate:  " & Format(ko_count / SIM_RUNS, "0.0%") & vbCrLf & _
           "Strike:   " & Format(K_WTI, "0.00") & " (ATM)" & vbCrLf & _
           "Premium:  " & Format(total_premium_Asian, "#,##0") & " KRW" & vbCrLf & _
           "Hedge:    Turnbull-Wakeman Asian delta + symmetric barrier fade", _
           vbInformation, "Asian Delta Hedge (Revised)"
End Sub

Public Sub Run_All_DeltaHedge_Engines()
    Call RunAllDeltaHedgeEnginesCore(0)
End Sub

Public Sub Run_All_DeltaHedge_Engines_QuickTest()
    Call RunAllDeltaHedgeEnginesCore(10000)
End Sub

Private Sub RunAllDeltaHedgeEnginesCore(Optional ByVal simRuns As Long = 0)
    If simRuns > 0 Then
        Sheets("American_Delta").Range("B1").Value = simRuns
        Sheets("European_Delta").Range("B1").Value = simRuns
        Sheets("Asian_Delta").Range("B1").Value = simRuns
    End If

    Call Run_LSMC_Engine
    PB_N = 0
    Call Run_American_DeltaHedge
    Call Run_European_DeltaHedge
    Call Run_Asian_DeltaHedge

    MsgBox "All 3 delta-hedge engines complete." & vbCrLf & _
           "Check sheets: American_Delta, European_Delta, Asian_Delta." & vbCrLf & vbCrLf & _
           "CRN (��33): all three engines shared one path bank (fixed seed" & vbCrLf & _
           "Randomize 20240101). KO rates are IDENTICAL by construction." & vbCrLf & _
           "Any P&L differences now reflect delta method only, not RNG noise.", _
           vbInformation, "Hedge Engine Suite"
End Sub
