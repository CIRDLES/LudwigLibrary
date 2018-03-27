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
package org.cirdles.ludwig.squid25;

import static org.cirdles.ludwig.isoplot3.Pub.age7corrWithErr;
import static org.cirdles.ludwig.isoplot3.Pub.pb76;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda232;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda235;
import static org.cirdles.ludwig.squid25.SquidConstants.lambda238;
import static org.cirdles.ludwig.squid25.SquidConstants.sComm0_64;
import static org.cirdles.ludwig.squid25.SquidConstants.sComm0_74;
import static org.cirdles.ludwig.squid25.SquidConstants.sComm0_76;
import static org.cirdles.ludwig.squid25.SquidConstants.sComm0_84;
import static org.cirdles.ludwig.squid25.SquidConstants.sComm0_86;
import static org.cirdles.ludwig.squid25.SquidConstants.uRatio;

/**
 * double implementations of Ken Ludwig's Squid.PbUTh VBA code for use with
 * Shrimp prawn files data reduction. Each public function returns a one-
 * dimensional array of double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/squid2.5Basic/PbUTh_2.bas" target="_blank">Squid.PbUTh_2</a>
 *
 * @author James F. Bowring
 */
public class PbUTh_2 {

    private PbUTh_2() {
    }

    /**
     * Ludwig specifies: Returns 204Pb/206Pb required to force
     * 206Pb/238U-207Pb/206Pb ages to concordance.
     *
     * @param pb76tot
     * @param age7corPb6U8
     * @return double [1] as{204Pb/206Pb}
     */
    public static double[] pb46cor7(double pb76tot, double age7corPb6U8) {
        return pb46cor7(pb76tot, sComm0_64, sComm0_74, age7corPb6U8);
    }

    /**
     * Ludwig specifies: Returns 204Pb/206Pb required to force
     * 206Pb/238U-207Pb/206Pb ages to concordance.
     *
     * @param pb76tot
     * @param alpha0
     * @param beta0
     * @param age7corPb6U8
     * @return double [1] as{204Pb/206Pb}
     */
    public static double[] pb46cor7(double pb76tot, double alpha0, double beta0, double age7corPb6U8) {

        double[] retVal = new double[]{0.0};

        double pb76true = pb76(age7corPb6U8)[0];

        try {
            retVal = new double[]{(pb76tot - pb76true) / (beta0 - pb76true * alpha0)};
        } catch (Exception e) {
        }
        return retVal;
    }

    /**
     * Ludwig specifies: Returns 204Pb/206Pb required to force
     * 206Pb/238U-208Pb/232Th ages to concordance.
     *
     * @param pb86tot
     * @param th2U8
     * @param age8corPb6U8
     * @return double [1] as {204Pb/206Pb}
     */
    public static double[] pb46cor8(double pb86tot, double th2U8, double age8corPb6U8) {
        return pb46cor8(pb86tot, th2U8, sComm0_64, sComm0_84, age8corPb6U8, lambda232, lambda238);
    }

    /**
     * Ludwig specifies: Returns 204Pb/206Pb required to force
     * 206Pb/238U-208Pb/232Th ages to concordance.
     *
     * @param pb86tot
     * @param th2U8
     * @param alpha0
     * @param gamma0
     * @param age8corPb6U8
     * @param lambda232
     * @param lambda238
     * @return double [1] as {204Pb/206Pb}
     */
    public static double[] pb46cor8(double pb86tot, double th2U8, double alpha0, double gamma0, double age8corPb6U8, double lambda232, double lambda238) {

        double[] retVal = new double[]{0.0};

        try {
            double numer = Math.expm1(age8corPb6U8 * lambda232);
            double denom = Math.expm1(age8corPb6U8 * lambda238);
            double pb86rad = numer / denom * th2U8;
            retVal = new double[]{(pb86tot - pb86rad) / (gamma0 - pb86rad * alpha0)};
        } catch (Exception e) {
        }
        return retVal;
    }

    /**
     * Ludwig specifies: Returns radiogenic 208Pb/206Pb %err where the common
     * 204Pb/206Pb is that required to force the 206Pb/238U-208Pb/232Th ages to
     * concordance.
     *
     * @param pb86tot
     * @param pb86totPer
     * @param pb76tot
     * @param pb76totPer
     * @param pb6U8tot
     * @param pb6U8totPer
     * @param age7corPb6U8
     * @return double [1] as {radiogenic 208Pb/206Pb %err(999 is bad)}
     */
    public static double[] pb86radCor7per(double pb86tot, double pb86totPer, double pb76tot,
            double pb76totPer, double pb6U8tot, double pb6U8totPer, double age7corPb6U8) {
        return pb86radCor7per(pb86tot, pb86totPer, pb76tot, pb76totPer, pb6U8tot, pb6U8totPer,
                age7corPb6U8, sComm0_64, sComm0_74, sComm0_84, lambda235, lambda238, uRatio);
    }

    /**
     * Ludwig specifies: Returns radiogenic 208Pb/206Pb %err where the common
     * 204Pb/206Pb is that required to force the 206Pb/238U-208Pb/232Th ages to
     * concordance.
     *
     * @param pb86tot
     * @param pb86totPer
     * @param pb76tot
     * @param pb76totPer
     * @param pb6U8tot
     * @param pb6U8totPer
     * @param age7corPb6U8
     * @param alpha0
     * @param beta0
     * @param gamma0
     * @param lambda235
     * @param lambda238
     * @param uRatio
     * @return double [1] as {radiogenic 208Pb/206Pb %err(999 is bad)}
     */
    public static double[] pb86radCor7per(double pb86tot, double pb86totPer, double pb76tot,
            double pb76totPer, double pb6U8tot, double pb6U8totPer, double age7corPb6U8,
            double alpha0, double beta0, double gamma0, double lambda235, double lambda238, double uRatio) {

        double[] retVal = new double[]{999.0};

        double r = pb6U8tot;
        double sigmaR = pb6U8totPer / 100.0 * r;
        double phi = pb76tot;
        double sigmaPhi = pb76totPer / 100.0 * phi;
        double theta = pb86tot;
        double sigmaTheta = pb86totPer / 100.0 * theta;
        double u = 1.0 / uRatio;
        double phi0 = beta0 / alpha0;

        double exp5 = Math.exp(age7corPb6U8 * lambda235);
        double exp8 = Math.exp(age7corPb6U8 * lambda238);
        double rStar = exp8 - 1.0;
        double sStar = exp5 - 1.0;
        double phiStar = sStar / rStar * u;
        double p = rStar / r;

        double alphaPrime = (phi - phiStar) / (beta0 - phiStar * alpha0);
        double thetaStar7 = (theta / alphaPrime - gamma0) / (1.0 / alphaPrime - alpha0);
        double m1 = lambda238 * exp8;
        double m2 = lambda235 * exp5;

        double j1 = p / m1 - sStar / r / m2;
        double j2 = 1.0 / (r / m1 - u * r * phi0 / m2);
        double d1 = 1.0 - alpha0 * alphaPrime;
        double d2 = beta0 - phiStar * alpha0;

        double k1 = (p - r * j1 * j2) / m1;
        double k2 = u * j2 * r * r / m1 / m2;
        double k3 = u / rStar * (m2 - phiStar * m1);
        double k4 = (alpha0 * alphaPrime - 1.0) / d2;
        double k5 = alphaPrime / d1;
        double k7 = (alpha0 * thetaStar7 - gamma0) / d1;

        double varThetaStar7 = Math.pow(sigmaTheta / d1, 2)
                + Math.pow(k1 * k3 * k4 * k7 * sigmaR, 2)
                + Math.pow(k7 * (1.0 / d2 + k2 * k3 * k4) * sigmaPhi, 2);

        if (varThetaStar7 >= 0.0) {
            double sigmaThetaStar7 = Math.sqrt(varThetaStar7);
            retVal = new double[]{100.0 * sigmaThetaStar7 / Math.abs(thetaStar7)};
        }

        return retVal;
    }

    /**
     * This method implements Ludwig's Age7CorrPb8Th2.
     *
     * Ludwig specifies Age7CorrPb8Th2: Returns the 208Pb/232Th age, assuming
     * the true 206/204 is that required to force 206/238-207/235 concordance.
     *
     * @param totPb206U238
     * @param totPb208Th232
     * @param totPb86
     * @param totPb76
     * @return double [1] as {age7CorrPb8Th2}
     */
    public static double[] age7CorrPb8Th2(double totPb206U238, double totPb208Th232,
            double totPb86, double totPb76)
            throws ArithmeticException {
        return age7CorrPb8Th2(totPb206U238, totPb208Th232, totPb86, totPb76, sComm0_64, sComm0_86, lambda232, lambda238);
    }

    /**
     * This method implements Ludwig's Age7CorrPb8Th2.
     *
     * Ludwig specifies Age7CorrPb8Th2: Returns the 208Pb/232Th age, assuming
     * the true 206/204 is that required to force 206/238-207/235 concordance.
     *
     * @param totPb206U238
     * @param totPb208Th232
     * @param totPb86
     * @param totPb76
     * @param sComm0_64
     * @param sComm0_86
     * @param lambda232
     * @param lambda238
     * @return double [1] as {age7CorrPb8Th2}
     */
    public static double[] age7CorrPb8Th2(double totPb206U238, double totPb208Th232,
            double totPb86, double totPb76, double sComm0_64, double sComm0_86, double lambda232, double lambda238)
            throws ArithmeticException {

        double gamma0 = sComm0_64 * sComm0_86;

        double age7corPb6U8
                = age7corrWithErr(totPb206U238, 0.0, totPb76, 0.0)[0];
        double radPb6U8 = Math.expm1(lambda238 * age7corPb6U8);
        double term = totPb206U238 - radPb6U8;
        term = term == 0 ? SquidConstants.SQUID_VERY_SMALL_VALUE : term;
        double alpha = sComm0_64 * totPb206U238 / term;
        double gamma = totPb86 * alpha;
        double radfract8 = (gamma - gamma0) / gamma;
        double radPb8Th2 = totPb208Th232 * radfract8;
        double age7corrPb8Th2 = Math.log(1.0 + radPb8Th2) / lambda232;

        return new double[]{age7corrPb8Th2};
    }

    /**
     * This method combines Ludwig's Age7CorrPb8Th2 and AgeErr7CorrPb8Th2.
     *
     * Ludwig specifies Age7CorrPb8Th2: Returns the 208Pb/232Th age, assuming
     * the true 206/204 is that required to force 206/238-207/235 concordance.
     *
     * Ludwig specifies AgeErr7CorrPb8Th2: Returns the error of the 208Pb/232Th
     * age, where the 208Pb/232Th age is calculated assuming the true 206/204 is
     * that required to force 206/238-207/235 concordance. The error is
     * calculated numerically, by successive perturbation of the input errors.
     *
     * @param totPb206U238
     * @param totPb206U238percentErr
     * @param totPb208Th232
     * @param totPb208Th232percentErr
     * @param totPb86
     * @param totPb86percentErr
     * @param totPb76
     * @param totPb76percentErr
     * @return double [2] as {age7CorrPb8Th2, age7CorrPb8Th2Err}
     */
    public static double[] age7CorrPb8Th2WithErr(double totPb206U238, double totPb206U238percentErr, double totPb208Th232,
            double totPb208Th232percentErr, double totPb86, double totPb86percentErr, double totPb76, double totPb76percentErr)
            throws ArithmeticException {

        // Perturb each input variable by its assigned error
        double ptotPb6U8 = (1.0 + totPb206U238percentErr / 100.0) * totPb206U238;
        double ptotPb8Th2 = (1.0 + totPb208Th232percentErr / 100.0) * totPb208Th232;
        double ptheta = totPb86 * (1.0 + totPb86percentErr / 100.0);
        double pphi = totPb76 * (1.0 + totPb76percentErr / 100.0);

        double ageVariance = 0.0;

        double[] delta = new double[5];

        // delta[0] is age
        delta[0] = age7CorrPb8Th2(totPb206U238, totPb208Th232, totPb86, totPb76)[0];
        delta[1] = age7CorrPb8Th2(ptotPb6U8, totPb208Th232, totPb86, totPb76)[0];
        delta[2] = age7CorrPb8Th2(totPb206U238, ptotPb8Th2, totPb86, totPb76)[0];
        delta[3] = age7CorrPb8Th2(totPb206U238, totPb208Th232, ptheta, totPb76)[0];
        delta[4] = age7CorrPb8Th2(totPb206U238, totPb208Th232, totPb86, pphi)[0];

        for (int i = 1; i < 5; i++) {
            ageVariance += Math.pow(delta[i] - delta[0], 2);
        }

        return new double[]{delta[0], Math.sqrt(ageVariance)};

    }

    /**
     * Ludwig specifies: Returns the radiogenic 206Pb/238U ratio for the
     * specified age.
     *
     * @param age
     * @return double [1] as {radiogenic 206Pb/238U ratio}
     * @throws ArithmeticException
     */
    public static double[] pb206U238rad(double age)
            throws ArithmeticException {
        return pb206U238rad(age, lambda238);
    }

    /**
     * Ludwig specifies: Returns the radiogenic 206Pb/238U ratio for the
     * specified age.
     *
     * @param age
     * @param lambda238
     * @return double [1] as {radiogenic 206Pb/238U ratio}
     * @throws ArithmeticException
     */
    public static double[] pb206U238rad(double age, double lambda238)
            throws ArithmeticException {
        return new double[]{Math.expm1(lambda238 * age)};
    }

    /**
     * This method combines Ludwig's Rad8corPb7U5 and Rad8corPb7U5Perr.
     *
     * Ludwig specifies Rad8corPb7U5: Returns the radiogenic 208-corrected
     * 207PbSTAR/235U ratio.
     *
     * Ludwig specifies Rad8corPb7U5Perr: Returns the %error of a 208-corrected
     * 207PbSTAR/235U.
     *
     * @param totPb6U8
     * @param totPb6U8per
     * @param radPb6U8
     * @param totPb7U5
     * @param th2U8
     * @param th2U8per
     * @param totPb76
     * @param totPb76per
     * @param totPb86
     * @param totPb86per
     * @return double [2] as {ratio, percent error}
     * @throws ArithmeticException
     */
    public static double[] rad8corPb7U5WithErr(double totPb6U8, double totPb6U8per,
            double radPb6U8, double totPb7U5, double th2U8, double th2U8per, double totPb76,
            double totPb76per, double totPb86, double totPb86per)
            throws ArithmeticException {
        return rad8corPb7U5WithErr(totPb6U8, totPb6U8per, radPb6U8, totPb7U5, th2U8, th2U8per,
                totPb76, totPb76per, totPb86, totPb86per, sComm0_76, sComm0_86, uRatio, lambda232, lambda238);
    }

    /**
     * This method combines Ludwig's Rad8corPb7U5 and Rad8corPb7U5Perr.
     *
     * Ludwig specifies Rad8corPb7U5: Returns the radiogenic 208-corrected
     * 207PbSTAR/235U ratio.
     *
     * Ludwig specifies Rad8corPb7U5Perr: Returns the %error of a 208-corrected
     * 207PbSTAR/235U.
     *
     * @param totPb6U8
     * @param totPb6U8per
     * @param radPb6U8
     * @param totPb7U5
     * @param th2U8
     * @param th2U8per
     * @param totPb76
     * @param totPb76per
     * @param totPb86
     * @param totPb86per
     * @param sComm0_76
     * @param sComm0_86
     * @param uRatio
     * @param lambda232
     * @param lambda238
     * @return double [2] as {ratio, percent error}
     * @throws ArithmeticException
     */
    public static double[] rad8corPb7U5WithErr(double totPb6U8, double totPb6U8per,
            double radPb6U8, double totPb7U5, double th2U8, double th2U8per, double totPb76,
            double totPb76per, double totPb86, double totPb86per, double sComm0_76, double sComm0_86,
            double uRatio, double lambda232, double lambda238)
            throws ArithmeticException {

        // calculate ratio
        double radFract6 = radPb6U8 / totPb6U8;
        double commFract6 = 1.0 - radFract6;
        double ratio = (totPb76 - commFract6 * sComm0_76) * uRatio * totPb6U8;

        // calculate error
        double k = 1.0 / th2U8;
        double SigmaTotPb6U8 = totPb6U8per / 100.0 * totPb6U8;
        double SigmaTotPb76 = totPb76per / 100.0 * totPb76;
        double SigmaTotPb86 = totPb86per / 100.0 * totPb86;
        double SigmaK = th2U8per / 100.0 * k;
        double u = uRatio;

        double q = totPb86 - sComm0_86 * commFract6;

        double radPb8Th2 = (totPb86 - commFract6 * sComm0_86) * totPb6U8 / th2U8;
        double radPb7U5 = (totPb76 - commFract6 * sComm0_76) * u * totPb6U8;
        double m1 = lambda238 * (1.0 + radPb6U8);
        double m2 = lambda232 * (1.0 + radPb8Th2);
        double h1 = radFract6 / m1 - q * k / m2;
        double h2 = 1.0 / (totPb6U8 / m1 - k * totPb6U8 * sComm0_86 / m2);

        double Term1 = Math.pow(h1 * SigmaTotPb6U8, 2);
        double Term2 = Math.pow(totPb6U8 / m2, 2);
        double Term3 = Math.pow(q * SigmaK, 2);
        double term4 = Math.pow(th2U8 * SigmaTotPb86, 2);

        double sigmaCommfract6 = Math.sqrt(h2 * h2 * (Term1 + Term2 * (Term3 + term4)));
        double covTotPb86CommFract6 = h1 * h2 * SigmaTotPb6U8 * SigmaTotPb6U8;

        Term1 = Math.pow(radPb7U5 / totPb6U8 * SigmaTotPb6U8, 2);
        Term2 = Math.pow(u * totPb6U8, 2) * (SigmaTotPb76 * SigmaTotPb76 + Math.pow(sComm0_76 * sigmaCommfract6, 2));
        Term3 = -2.0 * radPb7U5 * u * sComm0_76 * covTotPb86CommFract6;

        double sigmaRadPb7U5 = Math.sqrt(Term1 + Term2 + Term3);

        return new double[]{ratio, sigmaRadPb7U5 / radPb7U5 * 100.0};
    }

    /**
     * Ludwig specifies: Returns the error correlation for 208-corrected
     * 206PbSTAR/238U-207PbSTAR/235U ratio-pairs.
     *
     * @param totPb6U8
     * @param totPb6U8per
     * @param radPb6U8
     * @param th2U8
     * @param th2U8per
     * @param totPb76
     * @param totPb76per
     * @param totPb86
     * @param totPb86per
     * @return double [1] = {error correlation}
     * @throws ArithmeticException
     */
    public static double[] rad8corConcRho(double totPb6U8, double totPb6U8per, double radPb6U8,
            double th2U8, double th2U8per, double totPb76, double totPb76per,
            double totPb86, double totPb86per)
            throws ArithmeticException {
        return rad8corConcRho(totPb6U8, totPb6U8per, radPb6U8, th2U8, th2U8per, totPb76, totPb76per,
                totPb86, totPb86per, sComm0_76, sComm0_86, uRatio, lambda232, lambda238);
    }

    /**
     * Ludwig specifies: Returns the error correlation for 208-corrected
     * 206PbSTAR/238U-207PbSTAR/235U ratio-pairs.
     *
     * @param totPb6U8
     * @param totPb6U8per
     * @param radPb6U8
     * @param th2U8
     * @param th2U8per
     * @param totPb76
     * @param totPb76per
     * @param totPb86
     * @param totPb86per
     * @param sComm0_76
     * @param sComm0_86
     * @param uRatio
     * @param lambda232
     * @param lambda238
     * @return double [1] = {error correlation}
     * @throws ArithmeticException
     */
    public static double[] rad8corConcRho(double totPb6U8, double totPb6U8per, double radPb6U8,
            double th2U8, double th2U8per, double totPb76, double totPb76per,
            double totPb86, double totPb86per, double sComm0_76, double sComm0_86,
            double uRatio, double lambda232, double lambda238)
            throws ArithmeticException {

        double u = uRatio;
        double k = 1.0 / th2U8;
        double SigmaK = th2U8per / 100.0 * k;
        double SigmaTotPb76 = totPb76per / 100.0 * totPb76;
        double SigmaTotPb86 = totPb86per / 100.0 * totPb86;
        double SigmaTotPb6U8 = totPb6U8per / 100.0 * totPb6U8;
        double radFract6 = radPb6U8 / totPb6U8;
        double commFract6 = 1.0 - radFract6;
        double q = totPb86 - sComm0_86 * commFract6;

        double radPb8Th2 = (totPb86 - commFract6 * sComm0_86) * totPb6U8 / th2U8;
        double radPb7U5 = (totPb76 - commFract6 * sComm0_76) * u * totPb6U8;
        double m1 = lambda238 * (1.0 + radPb6U8);
        double m2 = lambda232 * (1.0 + radPb8Th2);
        double h1 = radFract6 / m1 - q * k / m2;
        double h2 = 1.0 / (totPb6U8 / m1 - k * totPb6U8 * sComm0_86 / m2);

        double Term1 = Math.pow(h1 * SigmaTotPb6U8, 2);
        double Term2 = Math.pow(totPb6U8 / m2, 2);
        double Term3 = Math.pow(q * SigmaK, 2);
        double term4 = Math.pow(k * SigmaTotPb86, 2);

        double SigmaCommfract6 = Math.sqrt(h2 * h2 * (Term1 + Term2 * (Term3 + term4)));

        double CovTotPb86CommFract6 = h1 * h2 * SigmaTotPb6U8 * SigmaTotPb6U8;
        double SigmaRadPb6U8 = Math.sqrt(Math.pow(radFract6 * SigmaTotPb6U8, 2) + totPb6U8 * totPb6U8
                * SigmaCommfract6 * SigmaCommfract6);

        Term1 = Math.pow(radPb7U5 / totPb6U8 * SigmaTotPb6U8, 2);
        Term2 = Math.pow(u * totPb6U8, 2) * (SigmaTotPb76 * SigmaTotPb76 + Math.pow(sComm0_76 * SigmaCommfract6, 2));
        Term3 = -2.0 * radPb7U5 * u * sComm0_76 * CovTotPb86CommFract6;

        double SigmaRadPb7U5 = Math.sqrt(Term1 + Term2 + Term3);

        Term1 = radFract6 * radPb7U5 / totPb6U8 * SigmaTotPb6U8 * SigmaTotPb6U8;
        Term2 = u * totPb6U8 * totPb6U8 * sComm0_76 * SigmaCommfract6 * SigmaCommfract6;
        Term3 = -CovTotPb86CommFract6 * (u * totPb6U8 * radFract6 * sComm0_76 + radPb7U5);

        double CovRad68Rad75 = Term1 + Term2 + Term3;

        return new double[]{CovRad68Rad75 / (SigmaRadPb6U8 * SigmaRadPb7U5)};
    }

    /**
     * Ludwig specifies: Returns radiogenic 208Pb/206Pb where the common
     * 204Pb/206Pb is that required to force the 206Pb/238U-208Pb/232Th ages to
     * concordance. Actually calculates the 1-sigma percent uncertainty! and is
     * designed for reference materials (aka standards) - JFB
     *
     * @param pb86tot
     * @param pb86totPer
     * @param pb76tot
     * @param pb76totPer
     * @param radPb86cor7
     * @param pb46cor7
     * @param stdRadPb76
     * @param alpha0
     * @param beta0
     * @param gamma0
     * @return double [1] = {radiogenic 208Pb/206Pb %err(999 is bad)}
     * @throws ArithmeticException
     */
    public static double[] stdPb86radCor7per(double pb86tot, double pb86totPer, double pb76tot,
            double pb76totPer, double radPb86cor7, double pb46cor7, double stdRadPb76,
            double alpha0, double beta0, double gamma0)
            throws ArithmeticException {

        double[] retVal;

        double alphaPrime = pb46cor7;
        double phi = pb76tot;
        double sigmaPhi = pb76totPer / 100.0 * phi;
        double theta = pb86tot;
        double sigmaTheta = pb86totPer / 100.0 * theta;
        double thetaStar7 = radPb86cor7;
        double phiStar = stdRadPb76;

        double d1 = 1.0 - alphaPrime * alpha0;
        double d2 = beta0 - alpha0 * phiStar;
        double k7 = (alpha0 * thetaStar7 - gamma0) / d1;

        double varThetaStar7 = Math.pow((sigmaTheta / d1), 2) + Math.pow((k7 / d2 * sigmaPhi), 2);

        if (varThetaStar7 < 0) {
            retVal = new double[]{999};
        } else {
            double sigmaThetaStar7 = Math.sqrt(varThetaStar7);
            retVal = new double[]{100.0 * sigmaThetaStar7 / Math.abs(thetaStar7)};
        }

        return retVal;
    }
}
