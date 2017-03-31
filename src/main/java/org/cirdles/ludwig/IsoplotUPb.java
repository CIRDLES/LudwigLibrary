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

import static org.cirdles.squid.SquidConstants.MAXEXP;
import static org.cirdles.squid.SquidConstants.lambda235;
import static org.cirdles.squid.SquidConstants.lambda238;
import static org.cirdles.squid.SquidConstants.uRatio;

/**
 *
 * @author James F. Bowring
 */
public class IsoplotUPb {

    /*
    Function PbPbAge(ByVal Pb#, Optional t1 = 0, Optional iAge, _
  Optional Err76, Optional WithLambdaErrs = False) As Double
' Calculates age in Ma from radiogenic Pb-207/206 to t1 (=0 if not passed);
'  but if 2 more params (iAge & Err76) are passed, param Pb is the
'  7/6 ratio, iAge the age for that ratio, & Err76 is the absolute
'  error in the 7/6 ratio.  Uses Newton's method.
' If WithLambdaErrs=TRUE, include decay-constant errors (at global sigma-level)
'  in the age error.
Dim Exp5#, Exp8#, Numer#, Denom#, Func#
Dim T#, term1#, term2#, Deriv#, Delta#
Dim Pb76#, BadSqrt As Boolean, Exp5t1#, Exp8t1#
Dim CalcErr As Boolean, Test5#, Test8#, test#, P#
Const Toler = 0.00001
ViM t1, 0
ViM WithLambdaErrs, False
If NIM(Err76) And NIM(iAge) Then CalcErr = True
Pb76 = Pb
GetConsts
If CalcErr Then
  T = iAge
ElseIf Pb76 > (Lambda235 / Lambda238 / Uratio) Then ' 7/6 @t=0
  T = 1000
Else
  T = -4000 ' Need a trial age to start
End If
Test5 = Lambda235 * t1: Test8 = Lambda238 * t1
If Abs(Test5) > MAXEXP Or Abs(Test8) > MAXEXP Then GoTo PbFail
Exp5t1 = Exp(Test5): Exp8t1 = Exp(Test8)
Do
  Test5 = Lambda235 * T: Test8 = Lambda238 * T
  If Abs(Test5) > MAXEXP Or Abs(Test8) > MAXEXP Then GoTo PbFail
  Exp5 = Exp(Test5):  Exp8 = Exp(Test8)
  Numer = Exp5t1 - Exp5: Denom = Exp8t1 - Exp8
  If Denom = 0 Then GoTo PbFail
  Func = Numer / Denom / Uratio
  term1 = -Lambda235 * Exp5
  term2 = Lambda238 * Exp8 * Numer / Denom
  Deriv = (term1 + term2) / Denom / Uratio
  If Deriv = 0 Then GoTo PbFail
  If CalcErr Then
    If WithLambdaErrs And t1 = 0 Then
      Numer = SQ((Exp8 - 1) * Err76) + SQ(T * Exp5 * SigLev * Lambda235err / Uratio) + _
       SQ(Pb76 * T * Exp8 * SigLev * Lambda238err)
      Denom = SQ(Pb76 * Lambda238 * Exp8 - Lambda235 * Exp5 / Uratio)
      If Denom = 0 Then GoTo PbFail
      TestSqrt Numer / Denom, P, BadSqrt
      If BadSqrt Then GoTo PbFail
      PbPbAge = P
    Else
      PbPbAge = Abs(Err76 / Deriv)
    End If
    Exit Function
  ElseIf Deriv = 0 Then
    GoTo PbFail
  End If
  Delta = (Pb76 - Func) / Deriv
  T = T + Delta
Loop Until Abs(Delta) < Toler
PbPbAge = T
Exit Function
PbFail: PbPbAge = BadT
End Function
     */
    public static double[][] pbPbAge(double pb76Rad)
            throws ArithmeticException {
        // mad toler smaller by factor of 10 from Ludwig
        double toler = 0.000001;
        double delta = 0.0;
        double t1 = 0.0;
        double t;

        // Ludwig has a dual-use for this method: either age OR uncertainty
        // age only for now
        if (pb76Rad > (lambda235 / lambda238 / uRatio)) { // 7/6 @t=0
            t = 1000000000.0;
        } else {
            t = -4000000000.0; // Need a trial age to start
        }

        double test235 = lambda235 * t1;
        double test238 = lambda238 * t1;
        double exp235t1 = Math.exp(test235);
        double exp238t1 = Math.exp(test238);

        do {
            test235 = lambda235 * t;
            test238 = lambda238 * t;
            if ((Math.abs(test235) > MAXEXP) || (Math.abs(test238) > MAXEXP)) {
                throw new ArithmeticException();
            }
            
            double exp235 = Math.exp(test235);
            double exp238 = Math.exp(test238);
            double numer = exp235t1 - exp235;
            double denom = exp238t1 - exp238;
            if (denom == 0.0) {
                throw new ArithmeticException();
            }

            double func = numer / denom / uRatio;
            double term1 = -lambda235 * exp235;
            double term2 = lambda238 * exp238 * numer / denom;
            double deriv = (term1 + term2) / denom / uRatio;
            if (deriv == 0) {
                throw new ArithmeticException();
            }

            delta = (pb76Rad - func) / deriv;
            t += delta;
            
            System.out.println (t + "   " + delta);

        } while (Math.abs(delta) >= toler);

        return new double[][]{{t, 0.0}};
    }
}
