/*
 * Copyright 2006-2017 CIRDLES.org.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.cirdles.ludwig;

import java.util.Arrays;

/**
 * double implementations of Ken Ludwig's Isoplot.Pub VBA code for use with
 * Shrimp prawn files data reduction. Each function returns a two dimensional
 * array of double.
 *
 * @see
 * https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/Pub.bas
 *
 * @author James F. Bowring
 */
public class IsoplotMeans {

    private IsoplotMeans() {
    }


    /*
     * Sub WeightedAv(W(), ValuesErrs, Optional PercentOut = False, Optional PercentIn = False, _
  Optional SigmaLevelIn = 2, Optional CanRej = False, Optional ConstExtErr = False, _
  Optional AltRej = False, Optional SigmaLevelOut = 2)
' Array function to calculate weighted averages.  ValuesErrs is a 2-column range
'  with values & errors, PercentOut specifies whether output errors are absolute
'  (default) or %; PercentIn specifies input errors as absolute (default) or %;
'  SigmaLevelIn is sigma-level of inputerrors (default is 2-sigma); OUTPUT ALWAYS 2-SIGMA.
'  CanRej permits rejection of outliers (default is False); If Probability<15%,
'  ConstExtErr specifies weighting by assigned errors plus a constant-external error
'  (default is weighting by assigned errors only).
' Errors are expanded by t-sigma-Sqrt(MSWD) if probability<10%.
' Output is a 2 columns of 5 (ConstExtErr=FALSE) or 6 (ConstExtErr=TRUE) rows, where
'  the left column contains values & the right captions.
Dim i&, Nrej&, j&, nR&, v#(), wt#(), IntErr68#, ExtErr68#
Dim Sigma#, v1#, v2#, tB As Boolean, Wrejected(), Nareas%, Co%
Dim StrkThru As Boolean, ww As wWtdAver, SL$(2), Cnf$(2), rj$, NoColZero As Boolean
ViM PercentOut, False
ViM PercentIn, False
ViM SigmaLevelIn, 2
ViM CanRej, False
ViM ConstExtErr, False
ViM AltRej, False
If SigmaLevelIn <= 0 Then SigmaLevelIn = 2
NoColZero = False: CanReject = CanRej
If TypeName(ValuesErrs) = "Range" Then
  Nareas = ValuesErrs.Areas.Count
Else
  Nareas = 1
End If
If AltRej Then ' If AltRej AND can't read font attribs, rejected values are indicated
  i = 99       '  by "rej" in column to left of values; otherwise by StrkThru.
  On Error Resume Next
  i = ValuesErrs(1, 1).Font.Strikethrough
  On Error GoTo 0
  If i <> 99 Then AltRej = False
End If
If AltRej Then   ' Is values-column in column A (so no "rej" possible)?
  i = 9
  On Error Resume Next
  i = IsNumeric(ValuesErrs(1, 0))
  On Error GoTo 0
  NoColZero = (i = 9)
End If
SigLev = 2: AbsErrs = True
nR = ValuesErrs.Rows.Count
ReDim v(nR, 2)
SL$(1) = "1-sigma": SL$(2) = "2-sigma"
Cnf$(1) = "68%-conf.": Cnf$(2) = "95%-conf."
N = 0
For i = 1 To nR
  v1 = Val(ValuesErrs(i, 1))
  v2 = Val(ValuesErrs.Areas(Nareas)(i, IIf(Nareas = 2, 1, 2)))
  StrkThru = False
  On Error Resume Next
  If AltRej Then
    If Not NoColZero Then StrkThru = (ValuesErrs(i, 0) = "rej")
  Else
    StrkThru = ValuesErrs(i, 1).Font.Strikethrough
  End If
  On Error GoTo 0
  If IsNumeric(v1) And IsNumeric(v2) Then
    If v1 <> 0 And v2 <> 0 And Not StrkThru Then
      N = 1 + N
      Sigma = v2 / SigmaLevelIn
      If PercentIn Then Sigma = Abs(Sigma / Hun * v1)
      v(N, 1) = v1: v(N, 2) = Sigma
    End If
  End If
Next i
ReDim wt(N), Wrejected(N, 2), InpDat(N, 2)
For i = 1 To N
  InpDat(i, 1) = v(i, 1): InpDat(i, 2) = v(i, 2)
Next i
Erase v
    
    
WeightedAverage N, ww, Nrej, Wrejected(), False, (SigmaLevelOut = 1), IntErr68, ExtErr68
With ww
  W(1, 1) = .IntMean: W(1, 2) = "Wtd Mean (from internal errs)"
  If ConstExtErr Then
    W(3, 2) = "Required external " & SL$(SigmaLevelOut)
    W(3, 1) = 0
  End If
  If .Probability > 0.05 Then
    W(2, 1) = .IntMeanErr2sigma / 2 * SigmaLevelOut
    W(2, 2) = SL$(SigmaLevelOut)
  ElseIf ConstExtErr Then
    W(1, 1) = .ExtMean
    W(1, 2) = "Wtd Mean (using ext. err)"
    W(2, 1) = Choose(SigmaLevelOut, .ExtMeanErr68, .ExtMeanErr95)  'ExtErr68, .ExtMeanErr95)
    W(2, 2) = Cnf$(SigmaLevelOut)
    W(3, 1) = .Ext2Sigma / 2 * SigmaLevelOut
  Else
    W(2, 1) = Choose(SigmaLevelOut, IntErr68, .IntMeanErr95)
    W(2, 2) = Cnf$(SigLev)
  End If
End With
W(2, 2) = W(2, 2) & " err. of mean"
If PercentOut Then
  W(2, 1) = Abs(W(2, 1) / W(1, 1) * Hun)
  If ConstExtErr Then W(3, 1) = Abs(W(3, 1) / W(1, 1) * Hun)
  W(2, 2) = W(2, 2) & " (%)"
  For i = 2 To 3 + ConstExtErr
    j = InStr(W(i, 2), "sigma")
    If j Then W(i, 2) = Left$(W(i, 2), j - 1) & "sigma%" & Mid$(W(i, 2), j + 5)
  Next i
End If
W(3 - ConstExtErr, 1) = ww.MSWD:        W(3 - ConstExtErr, 2) = "MSWD"
W(4 - ConstExtErr, 1) = Nrej:           W(4 - ConstExtErr, 2) = "rejected"
W(5 - ConstExtErr, 1) = ww.Probability: W(5 - ConstExtErr, 2) = "Probability of fit"
If CanReject And UBound(W, 1) >= j Then
  j = 6 - ConstExtErr: rj$ = ""
  If Nrej Then      ' Create space-delimited string containing index #s
    For i = 1 To N  '  of the rejected points.
      If Not IsEmpty(Wrejected(i, 1)) Then rj$ = rj$ & Str(i)
    Next i
    W(j, 1) = rj$: W(j, 2) = "rej. item #(s)"
  Else
    W(j, 1) = "": W(j, 2) = ""
  End If
End If
End Sub
     */
    public static double[][] weightedAv(double[] values, double[] errors) {
        return null;
    }

    /*
     * Sub WeightedAverage(ByVal Npts&, ww As wWtdAver, Nrej&, Wrejected(), _
  ByVal CanTukeys As Boolean, Optional OneSigOut = False, _
  Optional IntErr68 = 0, Optional ExtErr68 = 0)
Dim i&, j&, Count&, nU&
Dim N0&, Nn&, Weight#, SumWtdRatios#
Dim q#, IntMean#, ExtSigma#, t68#, t95#, WtdAvg#, WtdAvgErr#
Dim IntSigmaMean#, ExtSigmaMean#, TotSigmaMean#, PointError#, TotalError#
Dim Tolerance#, Sums#, InverseVar#(), temp#, Tot1sig#, IntVar#
Dim ExtVar#, WtdR2#, TotVar#, Trial#, Resid#, WtdResid#
Dim yy#(), IvarY#(), tbx#(), r$(4), NoEvar As Boolean, ExtXbarSigma#
Nn = Npts: Count = 0: Nrej = 0
If MinProb = 0 Then MinProb = Val(Menus("MinProb"))
ReDim InverseVar(Nn), yy(Nn), IvarY(Nn), tbx(Nn), yf.WtdResid(Nn)
If CanReject Then ReDim Wrejected(Nn, 2)

    
    For i = 1 To Nn
  InverseVar(i) = 1 / SQ(InpDat(i, 2))
Next i
Recalc:
ExtSigma = 0:  ww.Ext2Sigma = 0: Weight = 0
SumWtdRatios = 0: q = 0: Count = Count + 1
For i = 1 To Npts
  If InpDat(i, 1) <> 0 And InpDat(i, 2) <> 0 Then
    Weight = Weight + InverseVar(i)
    SumWtdRatios = SumWtdRatios + InverseVar(i) * InpDat(i, 1)
    q = q + InverseVar(i) * SQ(InpDat(i, 1))
  End If
Next i
nU = Nn - 1 ' Deg. freedom
t68 = StudentsT(nU, 68.26)
t95 = StudentsT(nU, 95)
IntMean = SumWtdRatios / Weight  ' "Internal" error of wtd average
ww.IntMean = IntMean             ' Use double-prec. var.in calcs!
'Sums = q - Weight * Sq(IntMean) ' Sums of squares of weighted deviates
'If Sums < 0 Then                ' May be better in case of roundoff probs
  Sums = 0
  For i = 1 To Npts
    If InpDat(i, 1) <> 0 And InpDat(i, 2) <> 0 Then
      Resid = InpDat(i, 1) - IntMean  ' Simple residual
      WtdResid = Resid / InpDat(i, 2) ' Wtd residual
      WtdR2 = SQ(WtdResid)            ' Square of wtd residual
      If Nn = Npts Then yf.WtdResid(i) = WtdResid 'WtdR2
      Sums = Sums + WtdR2
    End If
  Next i
  Sums = Max(Sums, 0)
'End If
    
With ww
  .MSWD = Sums / nU  ' Mean square of weighted deviates
  IntSigmaMean = Sqr(1 / Weight)
  .IntMeanErr2sigma = 2 * IntSigmaMean
  TotSigmaMean = IntSigmaMean * Sqr(.MSWD)
  .Probability = ChiSquare(.MSWD, (nU))
  .IntMeanErr95 = IntSigmaMean * IIf(.Probability >= 0.3, 1.96, t95 * Sqr(.MSWD))
  If OneSigOut Then
    IntErr68 = IntSigmaMean * IIf(.Probability >= 0.3, 0.9998, StudentsT(nU, 68.26) * Sqr(.MSWD))
  End If
End With
If ww.Probability < MinProb And ww.MSWD > 1 Then
  'Find the MLE constant external variance
  Nn = 0
  For i = 1 To Npts
    If InpDat(i, 1) <> 0 Then
     Nn = 1 + Nn: yy(Nn) = InpDat(i, 1)
     IvarY(Nn) = SQ(InpDat(i, 2))
  End If
  Next i
  ReDim Preserve yy(Nn), IvarY(Nn)
  IntVar = Nn / Weight        ' Mean internal variance of pts
  TotVar = IntVar * Sums / nU ' Mean total variance of pts
  WtdExtRTSEC 0, 10 * SQ(IntSigmaMean), ExtVar, ww.ExtMean, ExtXbarSigma, yy(), IvarY(), (Nn), NoEvar
  With ww
    If Not NoEvar Then
      ExtSigma = Sqr(ExtVar)
      ' Knowns: N of the x(i) plus N of the SigmaX(i)
      ' Unknowns: Xbar abd ExtVar
      .ExtMeanErr95 = StudentsT(2 * Nn - 2) * ExtXbarSigma
      If NIM(ExtErr68) Then
        ExtErr68 = StudentsT(2 * Nn - 2, 68.26) * ExtXbarSigma
      End If
      .Ext2Sigma = 2 * ExtSigma
    ElseIf .MSWD > 4 Then  ' Failure of RTSEC algorithm because of extremely high MSWD
      ExtSigma = App.StDev(yy)
      .ExtMean = iAverage(yy)
      .ExtMeanErr95 = t95 * ExtSigma / Sqr(Nn)
      If NIM(ExtErr68) Then
        ExtErr68 = StudentsT(nU, 68.26) * ExtSigma
      End If
      .Ext2Sigma = 2 * ExtSigma
    Else
      .ExtMean = 0: .ExtMeanErr95 = 0: .Ext2Sigma = 0
      If NIM(ExtErr68) Then ExtErr68 = 0
    End If
    .ExtMeanErr68 = t68 / t95 * .ExtMeanErr95
  End With
End If
If CanReject And ww.Probability < MinProb Then GoSub Reject
If CanTukeys Then
  For i = 1 To Npts: tbx(i) = InpDat(i, 1): Next i
  With ww
    TukeysBiweight tbx(), Npts, 6, .BiwtMean, .BiWtSigma, .BiWtErr95
    .Median = iMedian(tbx()): .MedianConf = MedianConfLevel(Npts)
    .MedianPlusErr = MedianUpperLim(tbx) - .Median
    .MedianMinusErr = .Median - MedianLowerLim(tbx)
  End With
End If
Exit Sub
     */
    public static double[][] weightedAverage(double[] values, double[] errors) {

        double[][] retVal = new double[][]{{0, 0, 0}};

        // check precondition of same size xValues and yValues and at least 3 points
        int nPts = values.length;
        if ((nPts == errors.length) && nPts > 2) {
            // proceed
            double[] inverseVar = new double[nPts];
            double[] wtdResid = new double[nPts];

            for (int i = 0; i < nPts; i++) {
                inverseVar[i] = Math.pow(errors[i], 2);
            }

            double extSigma = 0.0;
            double ext2Sigma = 0.0;
            double weight = 0.0;
            double sumWtdRatios = 0.0;
            double q = 0.0;
            int count = 0;

            for (int i = 0; i < nPts; i++) {
                if (values[i] * errors[i] != 0.0) {
                    weight += inverseVar[i];
                    sumWtdRatios += inverseVar[i] * values[i];
                    q += inverseVar[i] * Math.pow(values[i], 2);
                }
            }

            double intMean = sumWtdRatios / weight;//  ' "Internal" error of wtd average

            double sums = 0.0;
            for (int i = 0; i < nPts; i++) {
                if (values[i] * errors[i] != 0.0) {
                    double resid = values[i] - intMean;//  ' Simple residual
                    wtdResid[i] = resid / errors[i];// ' Wtd residual
                    double wtdR2 = Math.pow(wtdResid[i], 2);//' Square of wtd residual
                    sums += wtdR2;
                }
            }
            sums = Math.max(sums, 0.0);

            /*
            With ww
            .MSWD = Sums / nU  ' Mean square of weighted deviates
            IntSigmaMean = Sqr(1 / Weight)
            .IntMeanErr2sigma = 2 * IntSigmaMean
            TotSigmaMean = IntSigmaMean * Sqr(.MSWD)
            .Probability = ChiSquare(.MSWD, (nU))
            .IntMeanErr95 = IntSigmaMean * IIf(.Probability >= 0.3, 1.96, t95 * Sqr(.MSWD))
            If OneSigOut Then
              IntErr68 = IntSigmaMean * IIf(.Probability >= 0.3, 0.9998, StudentsT(nU, 68.26) * Sqr(.MSWD))
            End If
          End With
             */
        }

        return retVal;
    }
}
