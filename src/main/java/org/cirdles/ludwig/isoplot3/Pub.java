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
import static org.cirdles.ludwig.isoplot3.UPb.pbPbAge;
import static org.cirdles.ludwig.isoplot3.UPb.varTcalc;
import static org.cirdles.ludwig.isoplot3.U_2.inv2x2;
import static org.cirdles.ludwig.squid25.SquidConstants.MAXEXP;
import static org.cirdles.ludwig.squid25.SquidConstants.MAXLOG;
import static org.cirdles.ludwig.squid25.SquidConstants.MINLOG;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda232;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda235;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda238;
import static org.cirdles.ludwig.squid25.SquidConstants.sComm0_76;
import static org.cirdles.ludwig.squid25.SquidConstants.sComm0_86;
import static org.cirdles.ludwig.squid25.SquidConstants.uRatio;

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
            int upprInd = (int) conf95Calcs[1] - 1;

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
     * Ludwig: (Age No Lambda Errors) Using a 2-D Newton's method, find the age8Corr
 for a presumed-concordant point on the U-Pb Concordia diagram that
 minimizes Sums, assuming no decay-constant errors. See GCA 62, p.
     * 665-676, 1998 for explanation.
     *
     * @param xVal x-axis ratio from Concordia
     * @param yVal y-axis ratio from Concordia
     * @param covariance double[2][2] covariance matrix
     * @param trialAge estimate of age8Corr in annum
     * @return double[1] Age in annum
     */
    public static double[] ageNLE(double xVal, double yVal, double[][] covariance, double trialAge) {
        return Pub.ageNLE(xVal, yVal, covariance, trialAge, lambda235, lambda238);
    }

    /**
     * Ludwig: (Age No Lambda Errors) Using a 2-D Newton's method, find the age8Corr
 for a presumed-concordant point on the U-Pb Concordia diagram that
 minimizes Sums, assuming no decay-constant errors. See GCA 62, p.
     * 665-676, 1998 for explanation.
     *
     * @param xVal x-axis ratio from Concordia
     * @param yVal y-axis ratio from Concordia
     * @param covariance double[2][2] covariance matrix
     * @param trialAge estimate of age8Corr in annum
     * @param lambda235
     * @param lambda238
     * @return double[1] Age in annum
     */
    public static double[] ageNLE(
            double xVal,
            double yVal,
            double[][] covariance,
            double trialAge,
            double lambda235,
            double lambda238) {

        double[] retVal = new double[]{0.0};

        int count = 0;
        double T;

        int maxCount = 1000;
        double tolerance = 0.000001;
        double testTolerance;

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
                double d2b = (2.0 * Q5 * Q8 + Ry * Qq5 + Rx * Qq8) * inverted[2];
                double d2 = d2a + d2b;
                if (d2 != 0.0) {
                    double Incr = d1 / d2;
                    testTolerance = Math.abs(Incr / T);
                    t2 = T + Incr;
                    // age8Corr in annum
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
     * Ludwig: Returns Concordia age8Corr for T-W concordia data See Concordia
 function for usage.
     *
     * @note This implementation does not use inputs for rho or lambda
     * uncertainty inclusion
     *
     * @param r238U_206Pb
     * @param r238U_206Pb_1SigmaAbs
     * @param r207Pb_206Pb
     * @param r207Pb_206Pb_1SigmaAbs
     * @return double[4] {age8Corr, 1-sigma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordiaTW(double r238U_206Pb, double r238U_206Pb_1SigmaAbs, double r207Pb_206Pb, double r207Pb_206Pb_1SigmaAbs) {
        return Pub.concordiaTW(r238U_206Pb, r238U_206Pb_1SigmaAbs, r207Pb_206Pb, r207Pb_206Pb_1SigmaAbs, lambda235, lambda238, uRatio);
    }

    /**
     * Ludwig: Returns Concordia age8Corr for T-W concordia data See Concordia
 function for usage.
     *
     * @param lambda235
     * @param lambda238
     * @param uRatio
     * @note This implementation does not use inputs for rho or lambda
     * uncertainty inclusion
     *
     * @param r238U_206Pb
     * @param r238U_206Pb_1SigmaAbs
     * @param r207Pb_206Pb
     * @param r207Pb_206Pb_1SigmaAbs
     * @return double[4] {age8Corr, 1-sigma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordiaTW(
            double r238U_206Pb,
            double r238U_206Pb_1SigmaAbs,
            double r207Pb_206Pb,
            double r207Pb_206Pb_1SigmaAbs,
            double lambda235,
            double lambda238,
            double uRatio) {
        double[] retVal = new double[]{0, 0, 0};

        if ((r238U_206Pb > 0.0) && (r207Pb_206Pb > 0.0)) {
            double[] concConvert = concConvert(r238U_206Pb, r238U_206Pb_1SigmaAbs, r207Pb_206Pb, r207Pb_206Pb_1SigmaAbs, 0.0, true, uRatio);

            retVal = Pub.concordia(concConvert[0], concConvert[1], concConvert[2], concConvert[3], concConvert[4], lambda235, lambda238);
        }

        return retVal;
    }

    /**
     * Ludwig: Returns Concordia age8Corr for Conv.-concordia data; Input the
 Concordia X,err,Y,err,RhoXY Output is 1 range of 4 values -- t, t-error
 (1-sigma apriori),MSWD,Prob-of-fit If a second row is included in the
 output range, include names of the 4 result-values. Output errors are
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
     * @return double[4] {age8Corr, 1-sigma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordia(double r207Pb_235U, double r207Pb_235U_1SigmaAbs, double r206Pb_238U, double r206Pb_238U_1SigmaAbs, double rho) {
        return Pub.concordia(r207Pb_235U, r207Pb_235U_1SigmaAbs, r206Pb_238U, r206Pb_238U_1SigmaAbs, rho, lambda235, lambda238);
    }

    /**
     * Ludwig: Returns Concordia age8Corr for Conv.-concordia data; Input the
 Concordia X,err,Y,err,RhoXY Output is 1 range of 4 values -- t, t-error
 (1-sigma apriori),MSWD,Prob-of-fit If a second row is included in the
 output range, include names of the 4 result-values. Output errors are
     * always 2-sigma.
     *
     * @param lambda235
     * @param lambda238
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
     * @return double[4] {age8Corr, 1-sigma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordia(
            double r207Pb_235U,
            double r207Pb_235U_1SigmaAbs,
            double r206Pb_238U,
            double r206Pb_238U_1SigmaAbs,
            double rho,
            double lambda235,
            double lambda238) {
        double[] retVal = new double[]{0, 0, 0, 0};

        double inputData[];

        if ((r207Pb_235U > 0.0) && (r206Pb_238U > 0.0)) {
            inputData = new double[]{r207Pb_235U, r207Pb_235U_1SigmaAbs, r206Pb_238U, r206Pb_238U_1SigmaAbs, rho};

            retVal = Pub.concordiaAges(inputData, lambda235, lambda238);
        }

        return retVal;
    }

    /**
     * Ludwig: Calculate the weighted X-Y mean of the data pts (including error
     * correlation) & the "Concordia Age" & age8Corr-error of Xbar, Ybar. The
 "Concordia Age" is the most probable age8Corr of a data point if one can
 assume that the U/Pb ages of the true data point are precisely
 concordant. Calculates the age8Corr & error both with & without uranium
     * decay-constant errors. See GCA 62, p. 665-676, 1998 for explanation.
     *
     * @note this implementation only handles the case of one data point with no
     * lambda errors
     *
     * @param inputData double[5] containing r207Pb_235U, r207Pb_235U_1SigmaAbs,
     * r206Pb_238U, r206Pb_238U_1SigmaAbs, rho
     * @return double[4] {age8Corr, 1-sgma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordiaAges(double[] inputData) {
        return Pub.concordiaAges(inputData, lambda235, lambda238);
    }

    /**
     * Ludwig: Calculate the weighted X-Y mean of the data pts (including error
     * correlation) & the "Concordia Age" & age8Corr-error of Xbar, Ybar. The
 "Concordia Age" is the most probable age8Corr of a data point if one can
 assume that the U/Pb ages of the true data point are precisely
 concordant. Calculates the age8Corr & error both with & without uranium
     * decay-constant errors. See GCA 62, p. 665-676, 1998 for explanation.
     *
     * @param lambda235
     * @param lambda238
     * @note this implementation only handles the case of one data point with no
     * lambda errors
     *
     * @param inputData double[5] containing r207Pb_235U, r207Pb_235U_1SigmaAbs,
     * r206Pb_238U, r206Pb_238U_1SigmaAbs, rho
     * @return double[4] {age8Corr, 1-sgma abs uncert, MSWD, probabilityOfMSWD}
     */
    public static double[] concordiaAges(
            double[] inputData,
            double lambda235,
            double lambda238) {

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

            tNLE = Pub.ageNLE(xConc, yConc, vcXY, tNLE, lambda235, lambda238)[0];
            if (tNLE > 0.0) {
                double SumsAgeOneNLE = concordSums(xConc, yConc, vcXY, tNLE, lambda235, lambda238)[0];
                double MswdAgeOneNLE = SumsAgeOneNLE;
                double SigmaAgeNLE = varTcalc(vcXY, tNLE, lambda235, lambda238)[0];

                FDistribution fdist = new FDistribution(1, 1E9);
                double probability = 1.0 - fdist.cumulativeProbability(MswdAgeOneNLE);

                retVal = new double[]{tNLE, SigmaAgeNLE, MswdAgeOneNLE, probability};
            }

        }
        return retVal;
    }

    /**
     * Ludwig specifies Return radiogenic 207Pb/206Pb (secular equilibrium). All
     * calculations in annum.
     *
     * @param age
     * @return double [1] containing radiogenic 207Pb/206Pb.
     */
    public static double[] pb76(double age) {
        return Pub.pb76(age, lambda235, lambda238, uRatio);
    }

    /**
     * Ludwig specifies Return radiogenic 207Pb/206Pb (secular equilibrium). All
     * calculations in annum.
     *
     * @param age
     * @param lambda235
     * @param lambda238
     * @param uRatio
     * @return double [1] containing radiogenic 207Pb/206Pb.
     */
    public static double[] pb76(
            double age,
            double lambda235,
            double lambda238,
            double uRatio) {
        double[] retVal;

        if (age == 0.0) {
            retVal = new double[]{lambda235 / lambda238 / uRatio};
        } else {
            retVal = new double[]{Math.expm1(lambda235 * age) / Math.expm1(lambda238 * age) / uRatio};
        }

        return retVal;
    }

    /**
     * This method combines Ludwig's Age7Corr and AgeEr7Corr.
     *
     * Ludwig specifies Age7Corr: Age from uncorrected Tera-Wasserburg ratios,
 assuming the specified common-Pb 207/206.

 Ludwig specifies AgeEr7Corr: Calculation of 207-corrected age8Corr error.
     *
     * @param totPb6U8
     * @param totPb6U8err
     * @param totPb76
     * @param totPb76err
     * @return double [2] containing age7corrected, age7correctedErr
     */
    public static double[] age7corrWithErr(double totPb6U8, double totPb6U8err, double totPb76, double totPb76err)
            throws ArithmeticException {
        return Pub.age7corrWithErr(totPb6U8, totPb6U8err, totPb76, totPb76err, sComm0_76, lambda235, lambda238, uRatio);
    }

    /**
     * This method combines Ludwig's Age7Corr and AgeEr7Corr.
     *
     * Ludwig specifies Age7Corr: Age from uncorrected Tera-Wasserburg ratios,
 assuming the specified common-Pb 207/206.

 Ludwig specifies AgeEr7Corr: Calculation of 207-corrected age8Corr error.
     *
     * @param totPb6U8
     * @param totPb6U8err
     * @param totPb76
     * @param totPb76err
     * @param commPb76
     * @param lambda235
     * @param lambda238
     * @param uRatio
     * @return double [2] containing age7corrected, age7correctedErr
     */
    public static double[] age7corrWithErr(
            double totPb6U8,
            double totPb6U8err,
            double totPb76,
            double totPb76err,
            double commPb76,
            double lambda235,
            double lambda238,
            double uRatio)
            throws ArithmeticException {

        //commPb76 = sComm0_76;
        double commPb76err = 0.0;

        int iterationMax = 999;

        double totPb7U5 = totPb76 * uRatio * totPb6U8;
        double t = 0.0;
        double toler = 0.001;
        double delta = 0.0;
        double t1 = 1000.0e6;

        // Solve using Newton's method, using 1000 Ma as trial age8Corr.
        double e5;
        double e8;

        int iterations = 0;
        do {
            iterations++;

            t = t1;

            e5 = lambda235 * t;
            if (Math.abs(e5) > MAXEXP) {
                throw new ArithmeticException();
            }
            e8 = Math.exp(lambda238 * t);
            e5 = Math.exp(e5);
            double ee8 = e8 - 1.0;
            double ee5 = e5 - 1.0;

            double f = uRatio * commPb76 * (totPb6U8 - ee8) - totPb7U5 + ee5;
            double deriv = lambda235 * e5 - uRatio * commPb76 * lambda238 * e8;
            if (deriv == 0.0) {
                throw new ArithmeticException();
            }

            delta = -f / deriv;
            t1 = t + delta;

        } while ((Math.abs(delta) >= toler) && (iterations < iterationMax));

        // calculate error
        double totPb7U5var
                = uRatio * uRatio * (Math.pow(totPb6U8 * totPb76err, 2) + Math.pow(totPb76 * totPb6U8err, 2));
        e8 = Math.exp(lambda238 * t);
        e5 = Math.exp(lambda235 * t);
        double ee8 = e8 - 1.0;

        double denom = Math.pow(uRatio * commPb76 * lambda238 * e8 - lambda235 * e5, 2);
        double numer1 = Math.pow(uRatio * (totPb6U8 - ee8) * commPb76err, 2);
        double numer2 = uRatio * uRatio * commPb76 * (commPb76 - 2.0 * totPb76) * totPb6U8err * totPb6U8err;
        double numer3 = totPb7U5var;
        double numer = numer1 + numer2 + numer3;

        return new double[]{t, Math.sqrt(numer / denom)};
    }

    /**
     * This method combines Ludwig's AgePb76 and AgeErPb76.
     *
     * Ludwig specifies AgePb76: Age (Ma) from radiogenic 207Pb/206Pb (Note: we
 use annum here)

 Ludwig specifies AgeErPb76: Error in Pb7/6 age8Corr, input err is abs.
     *
     * @param pb76rad
     * @param pb76err
     * @return double [2] containing agePb76, agePb76Err
     * @throws ArithmeticException
     */
    public static double[] agePb76WithErr(double pb76rad, double pb76err)
            throws ArithmeticException {
        return Pub.agePb76WithErr(pb76rad, pb76err, lambda235, lambda238, uRatio);
    }

    /**
     * This method combines Ludwig's AgePb76 and AgeErPb76.
     *
     * Ludwig specifies AgePb76: Age (Ma) from radiogenic 207Pb/206Pb (Note: we
 use annum here)

 Ludwig specifies AgeErPb76: Error in Pb7/6 age8Corr, input err is abs.
     *
     * @param pb76rad
     * @param pb76err
     * @param lambda235
     * @param lambda238
     * @param uRatio
     * @return double [2] containing agePb76, agePb76Err
     * @throws ArithmeticException
     */
    public static double[] agePb76WithErr(
            double pb76rad,
            double pb76err,
            double lambda235,
            double lambda238,
            double uRatio)
            throws ArithmeticException {

        return pbPbAge(pb76rad, pb76err, lambda235, lambda238, uRatio);
    }

    public static void main(String[] args) {
        System.out.println(Arrays.toString(agePb76WithErr(0.05845338848554994, 4.84527392772108000)));
    }

    /**
     * This method combines Ludwig's Age8Corr and AgeEr8Corr.
     *
     * Ludwig specifies Age8Corr: Age from uncorrected Tera-Wasserburg ratios,
 assuming the specified common-Pb 207/206.

 Ludwig specifies AgeEr8Corr: Error in 208-corrected age8Corr (input-ratio
 errors are absolute).
     *
     * @param totPb6U8 double
     * @param totPb6U8err double
     * @param totPb8Th2 double
     * @param totPb8Th2err double
     * @param th2U8 double
     * @param th2U8err double
     * @return double [2] containing age8corrected, age8correctedErr
     */
    public static double[] age8corrWithErr(double totPb6U8, double totPb6U8err, double totPb8Th2, double totPb8Th2err,
            double th2U8, double th2U8err)
            throws ArithmeticException {
        return age8corrWithErr(totPb6U8, totPb6U8err, totPb8Th2, totPb8Th2err, th2U8, th2U8err, sComm0_86, lambda232, lambda238);
    }

    /**
     * This method combines Ludwig's Age8Corr and AgeEr8Corr.
     *
     * Ludwig specifies Age8Corr: Age from uncorrected Tera-Wasserburg ratios,
 assuming the specified common-Pb 207/206.

 Ludwig specifies AgeEr8Corr: Error in 208-corrected age8Corr (input-ratio
 errors are absolute).
     *
     * @param totPb6U8 double
     * @param totPb6U8err double
     * @param totPb8Th2 double
     * @param totPb8Th2err double
     * @param th2U8 double
     * @param th2U8err double
     * @param sComm0_86
     * @param lambda232
     * @param lambda238
     * @return double [2] containing age8corrected, age8correctedErr
     */
    public static double[] age8corrWithErr(
            double totPb6U8, 
            double totPb6U8err, 
            double totPb8Th2, 
            double totPb8Th2err,
            double th2U8, 
            double th2U8err,
            double sComm0_86,
            double lambda232,
            double lambda238)
            throws ArithmeticException {

        double commPb68 = 1.0 / sComm0_86;
        double commPb68err = 0.0;

        int iterationMax = 999;

        double t = 0.0;
        double toler = 0.001;
        double delta = 0.0;
        double t1 = 1000.0e6;

        // Solve using Newton's method, using 1000 Ma as trial age8Corr.
        double e2;
        double e8;

        int iterations = 0;
        do {
            iterations++;

            t = t1;

            e8 = lambda238 * t;
            if (Math.abs(e8) > MAXEXP) {
                throw new ArithmeticException();
            }
            e8 = Math.exp(e8);
            e2 = Math.exp(lambda232 * t);

            double f = totPb6U8 - e8 + 1.0 - th2U8 * commPb68 * (totPb8Th2 - e2 + 1);
            double deriv = th2U8 * commPb68 * lambda232 * e2 - lambda238 * e8;
            if (deriv == 0.0) {
                throw new ArithmeticException();
            }

            delta = -f / deriv;
            t1 = t + delta;

        } while ((Math.abs(delta) >= toler) && (iterations < iterationMax));

        double age8Corr = t1;
        
        // calculate error
        double g = totPb8Th2;
        double sigmaG = totPb8Th2err;
        double h = th2U8;
        double sigmaH = th2U8err;
        // July 2018 Simon dscovered that Ludwig zeroes this error - th2U8err;
        sigmaH = 0.0;
        double sigmaA = totPb6U8err;
        double psiI = commPb68;
        double sigmaPsiI = commPb68err;

        e2 = Math.exp(lambda232 * age8Corr);
        e8 = Math.exp(lambda238 * age8Corr);

        double p = g + 1.0 - e2;

        t1 = Math.pow(h * sigmaG, 2);
        double t2 = Math.pow(p * sigmaH, 2);
        double t3 = Math.pow(h * p * sigmaPsiI / psiI, 2);
        double k = lambda238 * e8 - h * psiI * lambda232 * e2;

        double numer = Math.pow(sigmaA, 2) + Math.pow(psiI, 2) * (t1 + t2 + t3);
        double denom = k * k;
        
         double ageEr8Corr = Math.sqrt(numer / denom);

        return new double[]{age8Corr, ageEr8Corr};
    }

    /**
     * Ludwig specifies Calculate the position, errs, & err corr of the weighted
     * mean of a suite of X-Y data pts.
     *
     * @param xValues double[]
     * @param xSigmaAbs double[]
     * @param yValues double[]
     * @param ySigmaAbs double[]
     * @param xyRho double[]
     * @return double[]{xBar, yBar, sumsXY, errX, errY, rhoXY}
     */
    public static double[] wtdXYmean(double[] xValues, double[] xSigmaAbs, double[] yValues, double[] ySigmaAbs, double[] xyRho) {
        double[] retVal = new double[]{0., 0., 0., 0., 0., 0.};

        int nPts = xValues.length;
        if ((xValues.length + xSigmaAbs.length + yValues.length + ySigmaAbs.length + xyRho.length) == (nPts * 5)) {

            double xVar;
            double yVar;
            double cov;
            double a = 0.0;
            double b = 0.0;
            double c = 0.0;
            double alpha = 0.0;
            double beta = 0.0;
            double[][] oh = new double[nPts][3];

            for (int i = 0; i < nPts; i++) {
                xVar = xSigmaAbs[i] * xSigmaAbs[i];
                yVar = ySigmaAbs[i] * ySigmaAbs[i];
                cov = xyRho[i] * Math.sqrt(xVar * yVar);
                double[] inverted = inv2x2(xVar, yVar, cov);
                a = a + inverted[0];
                b = b + inverted[1];
                c = c + inverted[2];
                alpha = alpha + xValues[i] * inverted[0] + yValues[i] * inverted[2];
                beta = beta + yValues[i] * inverted[1] + xValues[i] * inverted[2];
                oh[i][0] = inverted[0];
                oh[i][1] = inverted[1];
                oh[i][2] = inverted[2];
            }

            double denom = a * b - c * c;
            if (denom != 0.0) {
                double xBar = (b * alpha - beta * c) / denom;
                double yBar = (a * beta - alpha * c) / denom;

                double sumsXY = 0.0;

                for (int i = 0; i < nPts; i++) {
                    double rX = xValues[i] - xBar;
                    double rY = yValues[i] - yBar;
                    double s1 = rX * rX * oh[i][0] + rY * rY * oh[i][1];
                    double s2 = 2 * rX * rY * oh[i][2];
                    double wtdResidual = s1 + s2;
                    sumsXY = sumsXY + wtdResidual;
                }

                //  Now calculate the variance-covariance matrix of Xbar,Ybar
                double[] vcXY = inv2x2(a, b, c);
                double errX = Math.sqrt(vcXY[0]);
                double errY = Math.sqrt(vcXY[1]);
                double rhoXY = vcXY[2] / (errX * errY);

                retVal = new double[]{xBar, yBar, sumsXY, errX, errY, rhoXY};
            }
        }
        return retVal;
    }

    /**
     * Ludwig wrote this method as a wrapper for function wtdXYmean that returns
     * 2-sigma uncertainties. This method returns 1-sigma uncertainties in
     * keeping with our new standards.
     *
     * @param xValues double[]
     * @param xSigmaAbs double[]
     * @param yValues double[]
     * @param ySigmaAbs double[]
     * @param xyRho double[]
     * @return double [7] containing X mean, X uncertainty a priori, Y mean, Y
     * uncertainty a priori, uncertainty correlation, MSWD, probability of fit.
     */
    public static double[] xyWtdAv(double[] xValues, double[] xSigmaAbs, double[] yValues, double[] ySigmaAbs, double[] xyRho) {
        double[] retVal = new double[]{0., 0., 0., 0., 0., 0.};

        int nPts = xValues.length;
        if ((xValues.length + xSigmaAbs.length + yValues.length + ySigmaAbs.length + xyRho.length) == (nPts * 5)) {

            double[] wtdXYMean = wtdXYmean(xValues, xSigmaAbs, yValues, ySigmaAbs, xyRho);
            double MSWD;
            int df = 2 * nPts - 2;
            if (df <= 0) {
                MSWD = 0.0;
            } else {
                MSWD = wtdXYMean[2] / df;
            }

            FDistribution fdist = new FDistribution(df, 1E15);
            double probability = 1.0 - fdist.cumulativeProbability(MSWD);

            retVal = new double[]{wtdXYMean[0], wtdXYMean[3], wtdXYMean[1], wtdXYMean[4], wtdXYMean[5], MSWD, probability};
        }
        return retVal;
    }
}
