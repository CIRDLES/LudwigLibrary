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
 * double implementations of Ken Ludwig's Isoplot.UPb VBA code for use
 * with Shrimp prawn files data reduction. Each public function returns a two
 * dimensional array of double.
 *
 * @see
 * https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/UPb.bas
 * @author James F. Bowring
 */
public class IsoplotUPb {

    /**
     * Calculates age in annum from radiogenic Pb-207/206 and the absolute
     * 1-sigma uncertainty in the 7/6 ratio. Uses Newton's method. Does not
     * handle Ludwig's case of lambda uncertainties.
     *
     * @param r207_206r
     * @param r207_206r_1sigmaAbs is 1-sigma absolute
     * @return double[1][2] where [0][0] = age in annum and [0][1] = 1 sigma
     * uncertainty
     * @throws ArithmeticException
     */
    public static double[][] pbPbAge(double r207_206r, double r207_206r_1sigmaAbs)
            throws ArithmeticException {
        // made toler smaller by factor of 10 from Ludwig
        double toler = 0.000001;
        double delta = 0.0;
        double t1 = 0.0;
        double t;

        // adding an iteration counter to prevent thrashing
        int iterationMax = 100;
        int iterations = 0;

        // Need a trial age to start
        if (r207_206r > (lambda235 / lambda238 / uRatio)) {
            t = 1000000000.0;
        } else {
            t = -4000000000.0;
        }

        // intialize Newton's method
        double test235 = lambda235 * t1;
        double test238 = lambda238 * t1;
        double exp235t1 = Math.exp(test235);
        double exp238t1 = Math.exp(test238);

        double deriv = 0.0;

        do {
            iterations++;

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
            deriv = (term1 + term2) / denom / uRatio;
            if (deriv == 0) {
                throw new ArithmeticException();
            }

            delta = (r207_206r - func) / deriv;
            t += delta;

        } while ((Math.abs(delta) >= toler) && (iterations < iterationMax));

        return new double[][]{{t, Math.abs(r207_206r_1sigmaAbs / deriv)}};
    }
}
