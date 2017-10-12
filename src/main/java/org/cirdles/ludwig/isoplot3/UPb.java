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
package org.cirdles.ludwig.isoplot3;

import static org.cirdles.ludwig.isoplot3.U_2.inv2x2;
import org.cirdles.ludwig.squid25.SquidConstants;
import static org.cirdles.ludwig.squid25.SquidConstants.MAXEXP;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda235;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda238;
import static org.cirdles.ludwig.squid25.SquidConstants.uRatio;

/**
 * double implementations of Ken Ludwig's Isoplot.UPb VBA code for use with
 * Shrimp prawn files data reduction. Each public function returns an array of
 * double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/UPb.bas" target="_blank">Isoplot.UPb</a>
 * @author James F. Bowring
 */
public class UPb {

    private UPb() {
    }

    /**
     * Calculates age in annum from radiogenic Pb-207/206 and the absolute
     * 1-sigma uncertainty in the 7/6 ratio. Uses Newton's method. Does not
     * handle Ludwig's case of lambda uncertainties.
     *
     * @param r207_206r
     * @param r207_206r_1sigmaAbs is 1-sigma absolute
     * @return double[2] where [0] = age in annum and [1] = 1 sigma uncertainty
     * @throws ArithmeticException
     */
    public static double[] pbPbAge(double r207_206r, double r207_206r_1sigmaAbs)
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

        return new double[]{t, Math.abs(r207_206r_1sigmaAbs / deriv)};
    }

    /**
     * Ludwig: Calculate the sums of the squares of the weighted residuals for a
     * single Conv.-Conc. X-Y data point, where the true value of each of the
     * data pts is assumed to be on the same point on the Concordia curve, &
     * where the decay constants that describe the Concordia curve have known
     * uncertainties. See GCA 62, p. 665-676, 1998 for explanation.
     *
     * @note this implementation ignores lambda uncertainties.
     *
     * @param xConc double Concordia x-axis ratio
     * @param yConc double Concordia y-axis ratio
     * @param covariance double [2][2] matrix of age uncertainty covariances
     * @param t Age in annum
     * @return double[1] {concordSums}
     */
    public static double[] concordSums(double xConc, double yConc, double[][] covariance, double t) {

        double[] retVal = new double[]{0.0};

        double e5 = lambda235 * t;

        if (Math.abs(e5) <= MAXEXP) {
            e5 = Math.expm1(e5);
            double e8 = Math.expm1(lambda238 * t);
            double Ee5 = e5;// - 1.0;
            double Ee8 = e8;// - 1.0;
            double Rx = xConc - Ee5;
            double Ry = yConc - Ee8;

            double[] inverted = inv2x2(covariance[0][0], covariance[1][1], covariance[0][1]);

            retVal[0] = Rx * Rx * inverted[0] + Ry * Ry * inverted[1] + 2 * Rx * Ry * inverted[2];
        }
        return retVal;
    }

    /**
     * Ludwig: Calculate the variance in age for a single assumed-concordant
     * data point on the Conv. U/Pb concordia diagram (with or without taking
     * into account the uranium decay-constant errors). See GCA v62, p665-676,
     * 1998 for explanation.
     *
     * @param covariance double[2][2] matrix
     * @param t age in annum
     * @return double[1] {sigmaT 1-sigma abs uncertainty in age t}
     */
    public static double[] varTcalc(double[][] covariance, double t) {

        double[] retVal = new double[]{0.0};

        double e5 = lambda235 * t;
        if (Math.abs(e5) <= MAXEXP) {
            e5 = Math.exp(e5);
            double e8 = Math.exp(lambda238 * t);
            double Q5 = lambda235 * e5;
            double Q8 = lambda238 * e8;
//            double Xvar = covariance[0][0];
//            double Yvar = covariance[1][1];

//            double Cov = covariance[0][1];
            double[] inverted = inv2x2(covariance[0][0], covariance[1][1], covariance[0][1]);

            // Fisher is the expected second derivative with respect to T of the
            //  sums-of-squares of the weighted residuals.
            double Fisher = Q5 * Q5 * inverted[0] + Q8 * Q8 * inverted[1] + 2.0 * Q5 * Q8 * inverted[2];

            if (Fisher > 0.0) {
                retVal[0] = Math.sqrt(1.0 / Fisher);
            }
        }

        return retVal;
    }

}
