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

import java.util.Arrays;
import org.apache.commons.math3.distribution.FDistribution;
import static org.cirdles.ludwig.isoplot3.CMC.concConvert;
import static org.cirdles.ludwig.isoplot3.UPb.concordSums;
import static org.cirdles.ludwig.isoplot3.UPb.varTcalc;
import static org.cirdles.ludwig.isoplot3.U_2.inv2x2;
import static org.cirdles.ludwig.squid25.SquidConstants.MAXEXP;
import static org.cirdles.ludwig.squid25.SquidConstants.MAXLOG;
import static org.cirdles.ludwig.squid25.SquidConstants.MINLOG;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda235;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda238;

/**
 * double implementations of Ken Ludwig's Isoplot.Pub VBA code for use with
 * Shrimp prawn files data reduction. Each function returns an array of double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/Pub.bas" target="_blank">Isoplot.Pub</a>
 *
 * @author James F. Bowring
 */
public class Pub {

    private Pub() {
    }

    /**
     * Ludwig: Robust linear regression using median of all pairwise
     * slopes/intercepts, after Hoaglin, Mosteller & Tukey, Understanding Robust
     * & Exploratory Data Analysis, John Wiley & Sons, 1983, p. 160, with errors
     * from code in Rock & Duffy, 1986 (Comp. Geosci. 12, 807-818), derived from
     * Vugrinovich (1981), J. Math. Geol. 13, 443-454). Has simple, rapid
     * solution for errors. Ludwig used flags and our approach is to do all the
     * math and return all possible values available as if those flags were
     * true.
     *
     * @param xValues double [] array length n
     * @param yValues double [] array length n
     * @return double [9] containing slope, lSlope, uSlope, yInt, xInt, lYint,
     * uYint, lXint, uXint
     */
    public static double[] robustReg2(double[] xValues, double[] yValues) {

        double[] retVal = new double[]{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0};

        // check precondition of same size xValues and yValues and at least 3 points
        int n = xValues.length;
        if ((n == yValues.length) && n > 2) {
            // proceed
            double[][] slopeCalcs = RobustReg.getRobSlope(xValues, yValues);
            double slope = slopeCalcs[0][0];
            double yInt = slopeCalcs[0][1];
            double xInt = slopeCalcs[0][2];

            double[] slp = slopeCalcs[1];
            Arrays.sort(slp);
            double[] yInter = slopeCalcs[2];
            Arrays.sort(yInter);
            double[] xInter = slopeCalcs[3];
            Arrays.sort(xInter);

            double[] conf95Calcs = RobustReg.conf95(n, slp.length);
            // reduce indices by 1 to zero-based - this did not work but keeping them did
            // TODO: understand why - probably integer division related
            int lwrInd = (int) conf95Calcs[0] - 0;
            int upprInd = (int) conf95Calcs[1] - 0;

            double lSlope = slp[lwrInd];
            double uSlope = slp[upprInd];

            double lYint = yInter[lwrInd];
            double uYint = yInter[upprInd];

            double lXint = xInter[lwrInd];
            double uXint = xInter[upprInd];

            retVal = new double[]{slope, lSlope, uSlope, yInt, xInt, lYint, uYint, lXint, uXint};
        }

        return retVal;
    }

    /**
     * Ludwig: (Age No Lambda Errors) Using a 2-D Newton's method, find the age
     * for a presumed-concordant point on the U-Pb Concordia diagram that
     * minimizes Sums, assuming no decay-constant errors. See GCA 62, p.
     * 665-676, 1998 for explanation.
     *
     * @param xVal x-axis ratio from Concordia
     * @param yVal y-axis ratio from Concordia
     * @param covariance double[2][2] covariance matrix
     * @param trialAge estimate of age in annum
     * @return double[1] Age in annum
     */
    public static double[] ageNLE(double xVal, double yVal, double[][] covariance, double trialAge) {

        double[] retVal = new double[]{0.0};

        int count = 0;
        double T;

        int maxCount = 1000;
        double tolerance = 0.000001;
        // default value
        double testTolerance = 1.0;

        double[] inverted = inv2x2(covariance[0][0], covariance[1][1], covariance[0][1]);

        double t2 = trialAge;
        do {
            count++;
            T = t2;
            double e5 = lambda235 * T;

            if ((count < maxCount) && (Math.abs(e5) <= MAXEXP)) {
                e5 = Math.exp(e5);
                double e8 = Math.exp(lambda238 * T);
                double Ee5 = e5 - 1.0;
                double Ee8 = e8 - 1.0;
                double Q5 = lambda235 * e5;
                double Q8 = lambda238 * e8;
                double Qq5 = lambda235 * Q5;
                double Qq8 = lambda238 * Q8;
                double Rx = xVal - Ee5;
                double Ry = yVal - Ee8;

                // First derivative of T w.r.t. S, times -0.5
                double d1 = Rx * Q5 * inverted[0] + Ry * Q8 * inverted[1] + (Ry * Q5 + Rx * Q8) * inverted[2];

                // Second derivative of T w.r.t. S, times +0.5
                double d2a = (Q5 * Q5 + Qq5 * Rx) * inverted[0] + (Q8 * Q8 + Qq8 * Ry) * inverted[1];
                double d2b = (2 * Q5 * Q8 + Ry * Qq5 + Rx * Qq8) * inverted[2];
                double d2 = d2a + d2b;
                if (d2 != 0.0) {
                    double Incr = d1 / d2;
                    testTolerance = Math.abs(Incr / T);
                    t2 = T + Incr;
                    // age in annum
                    retVal[0] = t2;
                } else {
                    // force termination when d2 == 0;
                    testTolerance = 0.0;
                }
            } else {
                // force termination when count or e5 are out of bounds
                testTolerance = 0.0;
            }
        } while (testTolerance >= tolerance);

        return retVal;
    }

    /**
     * Ludwig: Returns Concordia age for T-W concordia data See Concordia
     * function for usage.
     *
     * @note This implementation does not use inputs for rho or lambda
     * uncertainty inclusion
     *
     * @param r238U_206Pb
     * @param r238U_206Pb_1SigmaAbs
     * @param r207Pb_206Pb
     * @param r207Pb_206Pb_1SigmaAbs
     * @return double[4] {age, 1-sigma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordiaTW(double r238U_206Pb, double r238U_206Pb_1SigmaAbs, double r207Pb_206Pb, double r207Pb_206Pb_1SigmaAbs) {
        double[] retVal = new double[]{0, 0, 0};

        if ((r238U_206Pb > 0.0) && (r207Pb_206Pb > 0.0)) {
            double[] concConvert = concConvert(r238U_206Pb, r238U_206Pb_1SigmaAbs, r207Pb_206Pb, r207Pb_206Pb_1SigmaAbs, 0.0, true);

            retVal = concordia(concConvert[0], concConvert[1], concConvert[2], concConvert[3], concConvert[4]);
        }

        return retVal;
    }

    /**
     * Ludwig: Returns Concordia age for Conv.-concordia data; Input the
     * Concordia X,err,Y,err,RhoXY Output is 1 range of 4 values -- t, t-error
     * (1-sigma apriori),MSWD,Prob-of-fit If a second row is included in the
     * output range, include names of the 4 result-values. Output errors are
     * always 2-sigma.
     *
     * @note this implementation outputs 1-sigma abs uncertainty
     *
     * @note Assume only one data point for now, with 1-sigma absolute
     * uncertainty for each coordinate.
     *
     * @param r207Pb_235U
     * @param r207Pb_235U_1SigmaAbs
     * @param r206Pb_238U
     * @param r206Pb_238U_1SigmaAbs
     * @param rho
     * @return double[4] {age, 1-sigma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordia(double r207Pb_235U, double r207Pb_235U_1SigmaAbs, double r206Pb_238U, double r206Pb_238U_1SigmaAbs, double rho) {
        double[] retVal = new double[]{0, 0, 0};

        double inputData[];

        if ((r207Pb_235U > 0.0) && (r206Pb_238U > 0.0)) {
            inputData = new double[]{r207Pb_235U, r207Pb_235U_1SigmaAbs, r206Pb_238U, r206Pb_238U_1SigmaAbs, rho};

            retVal = concordiaAges(inputData);
        }

        return retVal;
    }

    /**
     * Ludwig: Calculate the weighted X-Y mean of the data pts (including error
     * correlation) & the "Concordia Age" & age-error of Xbar, Ybar. The
     * "Concordia Age" is the most probable age of a data point if one can
     * assume that the U/Pb ages of the true data point are precisely
     * concordant. Calculates the age & error both with & without uranium
     * decay-constant errors. See GCA 62, p. 665-676, 1998 for explanation.
     *
     * @note this implementation only handles the case of one data point with no
     * lambda errors
     *
     * @param inputData double[5] containing r207Pb_235U, r207Pb_235U_1SigmaAbs,
     * r206Pb_238U, r206Pb_238U_1SigmaAbs, rho
     * @return double[4] {age, 1-sgma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordiaAges(double[] inputData) {

        double[] retVal = new double[]{0.0, 0.0, 0.0, 0.0};

        double xBar = inputData[0];
        double errX = inputData[1];
        double yBar = inputData[2];
        double errY = inputData[3];
        double rhoXY = inputData[4];

        double[][] vcXY = new double[2][2];

        double xConc = xBar;
        double yConc = yBar;
        vcXY[0][0] = errX * errX;
        vcXY[1][1] = errY * errY;
        vcXY[0][1] = rhoXY * errX * errY;
        vcXY[1][0] = vcXY[0][1];
        double trialAge = 1.0 + yConc;
        double tNLE;

        if ((trialAge >= MINLOG) && (trialAge <= MAXLOG) && (yConc > 0.0)) {
            tNLE = Math.log(trialAge) / lambda238;

            tNLE = ageNLE(xConc, yConc, vcXY, tNLE)[0];
            if (tNLE > 0.0) {
                double SumsAgeOneNLE = concordSums(xConc, yConc, vcXY, tNLE)[0];
                double MswdAgeOneNLE = SumsAgeOneNLE;
                double SigmaAgeNLE = varTcalc(vcXY, tNLE)[0];

                FDistribution fdist = new FDistribution(1, 1E9);
                double probability = 1.0 - fdist.cumulativeProbability(MswdAgeOneNLE);

                retVal = new double[]{tNLE, SigmaAgeNLE, MswdAgeOneNLE, probability};
            }

        }
        return retVal;
    }

}
