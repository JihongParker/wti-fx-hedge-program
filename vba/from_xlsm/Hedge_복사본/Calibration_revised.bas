Attribute VB_Name = "Calibration_revised"
Option Explicit

' =============================================================================
' MERGED 2026-06-22: this file is the "vol decomposition" bucket of the
' 3-file consolidation (Calibration_revised.bas / DeltaHedging_revised.bas /
' PaperVerification_revised.bas), per user request to reduce the project from
' 9 .bas files down to 3. Content is WTI_Jump_Calibration.bas, renamed only
' (Attribute VB_Name changed; no code changes). See MODEL_SPEC ��25.
' =============================================================================

' =============================================================================
' WTI jump/diffusion variance decomposition (2026-06-2x)
'
' Problem: Raw_Timeseries!H2 (-> LSMC!B10, vol1) is currently computed as
'   STDEV.S(all WTI log returns) * SQRT(252)
' i.e. naive total historical volatility, which already includes whatever
' jumps occurred in the historical window. LSMC!B1:B3 (Lambda, JumpMean,
' JumpVol) are separately hardcoded assumptions, not derived by removing the
' jump contribution from that same series. Combining them double-counts
' variance in the Merton SDE:
'   total model variance = vol1^2 + Lambda*(JumpMean^2 + JumpVol^2)
' Verified against the current inputs: this roughly doubles total annualized
' vol vs. the historical estimate (see MODEL_SPEC ��16).
'
' Fix: robust iterative k-sigma jump/diffusion split on the SAME WTI
' log-return series (Raw_Timeseries!F), so each observation contributes to
' exactly one component (diffusion OR jump), not both:
'   1. sigma_0 = robust scale estimate (MAD * 1.4826) -- avoids the masking
'      problem of starting from the jump-contaminated naive STDEV.
'   2. classify |r - mean(non-jump set)| > kSigma * sigma as "jump"
'   3. recompute sigma from non-jump days only
'   4. repeat 2-3 until the jump/non-jump split stops changing
'   5. vol1_new    = STDEV.S(non-jump days) * SQRT(252)
'      Lambda_new   = (jump day count / total days) * 252
'      JumpMean_new = AVERAGE(jump day returns)
'      JumpVol_new  = STDEV.S(jump day returns)
'
' Calibrate_WTI_JumpDiffusion writes a single-k result to Raw_Timeseries!K1:L8
' for review. Run_KSigma_Robustness_Check runs k=2.5/3.0/3.5 and writes a
' comparison table (paper-defense appendix material: "core conclusions are
' insensitive to the exact threshold choice"). Neither touches LSMC directly.
' Apply_WTI_JumpCalibration pushes the reviewed k=3 numbers into LSMC after an
' explicit confirmation prompt.
' =============================================================================

' =============================================================================
' Run_Full_WTI_Calibration_Pipeline -- orchestrator (2026-06-22)
' Chains the three steps below into one macro, so a single Excel shape/button
' can run the whole calibration pipeline instead of three separate manual
' clicks. Apply_WTI_JumpCalibration's own confirmation prompt (overwriting
' LSMC!B1:B3/B10) is NOT bypassed here -- it still asks before touching
' production values; this orchestrator only saves clicks for the two review
' steps that precede it. See MODEL_SPEC ��23.
' =============================================================================
Public Sub Run_Full_WTI_Calibration_Pipeline()
    Call Calibrate_WTI_JumpDiffusion
    Call Run_KSigma_Robustness_Check
    Call Apply_WTI_JumpCalibration
End Sub

Public Sub Calibrate_WTI_JumpDiffusion()

    Dim kSigma As Double: kSigma = 3#
    Dim wsRaw As Worksheet: Set wsRaw = Sheets("Raw_Timeseries")

    Dim vol1_new As Double, Lambda_new As Double, JumpMean_new As Double, JumpVol_new As Double
    Dim cntJump As Long, nTotal As Long, iterUsed As Long

    If Not DecomposeJumpDiffusion(kSigma, vol1_new, Lambda_new, JumpMean_new, JumpVol_new, cntJump, nTotal, iterUsed) Then
        MsgBox "Not enough WTI return data to calibrate.", vbExclamation
        Exit Sub
    End If

    ' --- Old (naive) values for comparison ---
    Dim vol1_old As Double: vol1_old = wsRaw.Range("H2").Value
    Dim Lambda_old As Double, JumpMean_old As Double, JumpVol_old As Double
    On Error Resume Next
    Lambda_old = Sheets("LSMC").Range("B1").Value
    JumpMean_old = Sheets("LSMC").Range("B2").Value
    JumpVol_old = Sheets("LSMC").Range("B3").Value
    On Error GoTo 0

    Dim totalVar_old As Double: totalVar_old = vol1_old ^ 2 + Lambda_old * (JumpMean_old ^ 2 + JumpVol_old ^ 2)
    Dim totalVar_new As Double: totalVar_new = vol1_new ^ 2 + Lambda_new * (JumpMean_new ^ 2 + JumpVol_new ^ 2)

    ' --- Write results for review (LSMC sheet is NOT touched here) ---
    wsRaw.Range("K1").Value = "Diffusion-only vol1 (new)"
    wsRaw.Range("L1").Value = vol1_new
    wsRaw.Range("K2").Value = "Lambda (new, /yr)"
    wsRaw.Range("L2").Value = Lambda_new
    wsRaw.Range("K3").Value = "JumpMean (new)"
    wsRaw.Range("L3").Value = JumpMean_new
    wsRaw.Range("K4").Value = "JumpVol (new)"
    wsRaw.Range("L4").Value = JumpVol_new
    wsRaw.Range("K5").Value = "Jump days detected"
    wsRaw.Range("L5").Value = cntJump
    wsRaw.Range("K6").Value = "Total trading days"
    wsRaw.Range("L6").Value = nTotal
    wsRaw.Range("K7").Value = "Implied total annual vol (old, naive)"
    wsRaw.Range("L7").Value = Sqr(totalVar_old)
    wsRaw.Range("K8").Value = "Implied total annual vol (new, decomposed)"
    wsRaw.Range("L8").Value = Sqr(totalVar_new)
    wsRaw.Range("K9").Value = "k-sigma used"
    wsRaw.Range("L9").Value = kSigma
    wsRaw.Columns("K:L").AutoFit

    MsgBox "WTI jump/diffusion calibration complete (robust " & kSigma & "-sigma, " & iterUsed & " iterations)." & vbCrLf & vbCrLf & _
           "Jump days: " & cntJump & " / " & nTotal & "  (implied Lambda = " & Format(Lambda_new, "0.00") & "/yr)" & vbCrLf & vbCrLf & _
           "                    Old (naive)    New (decomposed)" & vbCrLf & _
           "vol1 (diffusion):   " & Format(vol1_old, "0.0000") & "        " & Format(vol1_new, "0.0000") & vbCrLf & _
           "Lambda:             " & Format(Lambda_old, "0.0000") & "        " & Format(Lambda_new, "0.0000") & vbCrLf & _
           "JumpMean:           " & Format(JumpMean_old, "0.0000") & "       " & Format(JumpMean_new, "0.0000") & vbCrLf & _
           "JumpVol:            " & Format(JumpVol_old, "0.0000") & "        " & Format(JumpVol_new, "0.0000") & vbCrLf & vbCrLf & _
           "Implied TOTAL annual vol:  old " & Format(Sqr(totalVar_old), "0.0%") & "  ->  new " & Format(Sqr(totalVar_new), "0.0%") & vbCrLf & vbCrLf & _
           "Results written to Raw_Timeseries!K1:L9 for review." & vbCrLf & _
           "LSMC!B1:B3/B10 were NOT modified. Review the numbers, then run" & vbCrLf & _
           "Apply_WTI_JumpCalibration to push them into LSMC.", _
           vbInformation, "WTI Jump/Diffusion Calibration"

End Sub

Public Sub Run_KSigma_Robustness_Check()
    ' Paper-defense material: re-runs the decomposition at k = 2.5, 3.0, 3.5
    ' and tabulates the results side by side, so the appendix can show the
    ' core calibration is not sensitive to the exact threshold choice.
    Dim wsRaw As Worksheet: Set wsRaw = Sheets("Raw_Timeseries")

    wsRaw.Range("K11").Value = "k-sigma"
    wsRaw.Range("L11").Value = "vol1 (diffusion)"
    wsRaw.Range("M11").Value = "Lambda (/yr)"
    wsRaw.Range("N11").Value = "JumpMean"
    wsRaw.Range("O11").Value = "JumpVol"
    wsRaw.Range("P11").Value = "Jump days"
    wsRaw.Range("Q11").Value = "Implied total vol"

    Dim ks() As Variant: ks = Array(2.5, 3#, 3.5)
    Dim row As Long: row = 12
    Dim idx As Long
    For idx = LBound(ks) To UBound(ks)
        Dim k As Double: k = ks(idx)
        Dim vol1_new As Double, Lambda_new As Double, JumpMean_new As Double, JumpVol_new As Double
        Dim cntJump As Long, nTotal As Long, iterUsed As Long

        If DecomposeJumpDiffusion(k, vol1_new, Lambda_new, JumpMean_new, JumpVol_new, cntJump, nTotal, iterUsed) Then
            Dim totalVar As Double: totalVar = vol1_new ^ 2 + Lambda_new * (JumpMean_new ^ 2 + JumpVol_new ^ 2)
            wsRaw.Cells(row, 11).Value = k
            wsRaw.Cells(row, 12).Value = vol1_new
            wsRaw.Cells(row, 13).Value = Lambda_new
            wsRaw.Cells(row, 14).Value = JumpMean_new
            wsRaw.Cells(row, 15).Value = JumpVol_new
            wsRaw.Cells(row, 16).Value = cntJump
            wsRaw.Cells(row, 17).Value = Sqr(totalVar)
        End If
        row = row + 1
    Next idx

    wsRaw.Columns("K:Q").AutoFit

    MsgBox "k-sigma robustness check complete (k = 2.5, 3.0, 3.5)." & vbCrLf & _
           "Results written to Raw_Timeseries!K11:Q14." & vbCrLf & vbCrLf & _
           "For the paper appendix: confirm vol1/Lambda/JumpMean/JumpVol don't move" & vbCrLf & _
           "drastically across these three thresholds before citing '3-sigma, the" & vbCrLf & _
           "conventional outlier-detection bound' as the headline choice.", _
           vbInformation, "k-Sigma Robustness Check"
End Sub

Public Sub Apply_WTI_JumpCalibration()
    ' Pushes the reviewed Raw_Timeseries!L1:L4 values (from the most recent
    ' Calibrate_WTI_JumpDiffusion call) into LSMC!B1:B3,B10.
    Dim wsRaw As Worksheet: Set wsRaw = Sheets("Raw_Timeseries")
    Dim wsLSMC As Worksheet: Set wsLSMC = Sheets("LSMC")

    If wsRaw.Range("L1").Value = 0 And wsRaw.Range("L2").Value = 0 Then
        MsgBox "Run Calibrate_WTI_JumpDiffusion first and review Raw_Timeseries!K1:L9.", vbExclamation
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox("This will overwrite:" & vbCrLf & _
                  "  LSMC!B1 (Lambda)    -> " & Format(wsRaw.Range("L2").Value, "0.0000") & vbCrLf & _
                  "  LSMC!B2 (JumpMean)  -> " & Format(wsRaw.Range("L3").Value, "0.0000") & vbCrLf & _
                  "  LSMC!B3 (JumpVol)   -> " & Format(wsRaw.Range("L4").Value, "0.0000") & vbCrLf & _
                  "  LSMC!B10 (vol1)     -> " & Format(wsRaw.Range("L1").Value, "0.0000") & vbCrLf & vbCrLf & _
                  "LSMC!B10 currently links to Raw_Timeseries!H2 (the naive formula)." & vbCrLf & _
                  "This will replace it with a static decomposed value, breaking that link." & vbCrLf & vbCrLf & _
                  "Continue?", vbYesNo + vbQuestion, "Apply WTI Jump Calibration")
    If resp <> vbYes Then Exit Sub

    wsLSMC.Range("B1").Value = wsRaw.Range("L2").Value    ' Lambda
    wsLSMC.Range("B2").Value = wsRaw.Range("L3").Value    ' JumpMean
    wsLSMC.Range("B3").Value = wsRaw.Range("L4").Value    ' JumpVol
    wsLSMC.Range("B10").Value = wsRaw.Range("L1").Value   ' vol1 (static value, overwrites the =Raw_Timeseries!H2 link)

    MsgBox "LSMC!B1:B3 and B10 updated with the decomposed jump/diffusion calibration." & vbCrLf & _
           "Re-run Run_LSMC_Engine (and Run_American_DeltaHedge etc.) to re-price under the corrected parameters.", _
           vbInformation
End Sub

' =============================================================================
' DecomposeJumpDiffusion -- core robust iterative k-sigma decomposition.
' Shared by Calibrate_WTI_JumpDiffusion (single k) and
' Run_KSigma_Robustness_Check (k sensitivity table). Returns False if there
' isn't enough data.
' =============================================================================
Private Function DecomposeJumpDiffusion(ByVal kSigma As Double, _
                                         ByRef vol1_new As Double, ByRef Lambda_new As Double, _
                                         ByRef JumpMean_new As Double, ByRef JumpVol_new As Double, _
                                         ByRef cntJump As Long, ByRef nTotal As Long, ByRef iterUsed As Long) As Boolean

    Dim wsRaw As Worksheet: Set wsRaw = Sheets("Raw_Timeseries")
    Const MAX_ITER As Long = 10

    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, "B").End(xlUp).row

    Dim n As Long: n = lastRow - 2   ' returns start at row 3 (F3 = LN(B3/B2))
    If n < 30 Then
        DecomposeJumpDiffusion = False
        Exit Function
    End If
    nTotal = n

    Dim ret() As Double: ReDim ret(1 To n)
    Dim isJump() As Boolean: ReDim isJump(1 To n)
    Dim i As Long
    For i = 1 To n
        ret(i) = wsRaw.Cells(i + 2, "F").Value
        isJump(i) = False
    Next i

    ' Robust starting scale: MAD * 1.4826 (avoids masking by the jumps themselves)
    Dim sigma As Double
    sigma = RobustMAD(ret, n) * 1.4826

    Dim meanNormal As Double
    Dim prevJumpCount As Long: prevJumpCount = -1
    Dim jumpCount As Long

    Dim iter As Long
    For iter = 1 To MAX_ITER
        iterUsed = iter

        Dim sumN As Double: sumN = 0#
        Dim cntN As Long: cntN = 0
        For i = 1 To n
            If Not isJump(i) Then
                sumN = sumN + ret(i)
                cntN = cntN + 1
            End If
        Next i
        meanNormal = IIf(cntN > 0, sumN / cntN, 0#)

        jumpCount = 0
        For i = 1 To n
            isJump(i) = (Abs(ret(i) - meanNormal) > kSigma * sigma)
            If isJump(i) Then jumpCount = jumpCount + 1
        Next i

        If jumpCount = prevJumpCount Then Exit For   ' converged
        prevJumpCount = jumpCount

        Dim sumsq As Double: sumsq = 0#
        cntN = 0
        For i = 1 To n
            If Not isJump(i) Then
                sumsq = sumsq + (ret(i) - meanNormal) ^ 2
                cntN = cntN + 1
            End If
        Next i
        If cntN > 1 Then sigma = Sqr(sumsq / (cntN - 1))
    Next iter

    ' Final split statistics
    Dim sumNorm As Double, sumJump As Double
    Dim cntNorm As Long
    sumNorm = 0#: sumJump = 0#: cntNorm = 0: cntJump = 0

    For i = 1 To n
        If isJump(i) Then
            sumJump = sumJump + ret(i)
            cntJump = cntJump + 1
        Else
            sumNorm = sumNorm + ret(i)
            cntNorm = cntNorm + 1
        End If
    Next i

    Dim meanNorm As Double: meanNorm = IIf(cntNorm > 0, sumNorm / cntNorm, 0#)
    Dim meanJump As Double: meanJump = IIf(cntJump > 0, sumJump / cntJump, 0#)

    Dim sumNormSq As Double, sumJumpSq As Double
    sumNormSq = 0#: sumJumpSq = 0#
    For i = 1 To n
        If isJump(i) Then
            sumJumpSq = sumJumpSq + (ret(i) - meanJump) ^ 2
        Else
            sumNormSq = sumNormSq + (ret(i) - meanNorm) ^ 2
        End If
    Next i

    vol1_new = IIf(cntNorm > 1, Sqr(sumNormSq / (cntNorm - 1)) * Sqr(252), 0#)
    Lambda_new = (cntJump / n) * 252
    JumpMean_new = meanJump
    JumpVol_new = IIf(cntJump > 1, Sqr(sumJumpSq / (cntJump - 1)), 0#)

    DecomposeJumpDiffusion = True
End Function

' =============================================================================
' RobustMAD -- Median Absolute Deviation (unscaled). Used as a jump-resistant
' starting scale estimate, avoiding the masking problem of seeding the
' threshold from the jump-contaminated naive STDEV.
' =============================================================================
Private Function RobustMAD(ByRef ret() As Double, ByVal n As Long) As Double
    Dim sorted() As Double: ReDim sorted(1 To n)
    Dim i As Long
    For i = 1 To n
        sorted(i) = ret(i)
    Next i

    Dim med As Double: med = MedianOf(sorted, n)

    Dim dev() As Double: ReDim dev(1 To n)
    For i = 1 To n
        dev(i) = Abs(ret(i) - med)
    Next i

    RobustMAD = MedianOf(dev, n)
End Function

Private Function MedianOf(ByRef arr() As Double, ByVal n As Long) As Double
    Dim tmp As Variant
    tmp = arr
    MedianOf = WorksheetFunction.Median(tmp)
End Function
