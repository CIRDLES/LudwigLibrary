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

import java.util.ArrayList;
import java.util.List;
import org.apache.commons.math3.distribution.FDistribution;
import org.apache.commons.math3.distribution.TDistribution;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import static org.cirdles.ludwig.squid25.SquidConstants.SQUID_MINIMUM_PROBABILITY;
import org.cirdles.ludwig.squid25.SquidMathUtils;
import org.cirdles.ludwig.squid25.Utilities;
import static org.cirdles.ludwig.squid25.Utilities.median;

/**
 * double implementations of Ken Ludwig's Isoplot.Pub VBA code for use with
 * Shrimp prawn files data reduction. Each function returns an array of double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/Pub.bas" target="_blank">Isoplot.Pub</a>
 *
 * @author James F. Bowring
 */
public class Means {

    private Means() {
    }

    /**
     * Ludwig's WeightedAverage and assumes ConstExtErr = true since all
     * possible values are returned and caller can decide
     *
     * @param inValues as double[] with length nPts
     * @param inErrors as double[] with length nPts
     * @param canReject
     * @param canTukeys
     * @return double[2][7]{mean, 1sigmaMean, exterr68, exterr95, MSWD, probability, externalFlag}, {values
     * with rejected as 0.0}.  externalFlag = 1.0 for external uncertainty, 0.0 for internal
     */
    public static double[][] weightedAverage(double[] inValues, double[] inErrors, boolean canReject, boolean canTukeys) {

        double[] values = inValues.clone();
        double[] errors = inErrors.clone();

        double[][] retVal = new double[][]{{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0}, {}};

        // check precondition of same size values and errors and at least 3 points
        int nPts = values.length;
        int nN = nPts;
        int count = 0;

        // where does this come from??
        boolean hardRej = false;

        if ((nPts == errors.length) && nPts > 2) {
            // proceed
            double[] inverseVar = new double[nPts];
            double[] wtdResid = new double[nPts];
            double[] yy;
            double[] iVarY;
            double[] tbX = new double[nPts];

            double[][] wRejected = new double[nPts][2];

            for (int i = 0; i < nPts; i++) {
                inverseVar[i] = 1.0 / Math.pow(errors[i], 2);
            }

            double intMean = 0.0;
            double MSWD = 0.0;
            double intSigmaMean = 0.0;
            double probability = 0.0;
            double intMeanErr95 = 0.0;
            double intErr68 = 0.0;

            double extMean = 0.0;
            double extMeanErr95 = 0.0;
            double extMeanErr68 = 0.0;
            double extSigma = 0.0;

            double biWtMean = 0.0;
            double biWtSigma = 0.0;

            boolean reCalc;

            // entry point for RECALC goto - consider another private method?
            do {
                reCalc = false;

                extSigma = 0.0;
                double weight = 0.0;
                double sumWtdRatios = 0.0;
                double q = 0.0;
                count++;

                for (int i = 0; i < nPts; i++) {
                    if (values[i] * errors[i] != 0.0) {
                        weight += inverseVar[i];
                        sumWtdRatios += inverseVar[i] * values[i];
                        q += inverseVar[i] * Math.pow(values[i], 2);
                    }
                }

                int nU = nN - 1;// ' Deg. freedom
                TDistribution studentsT = new TDistribution(nU);
                // see https://stackoverflow.com/questions/21730285/calculating-t-inverse
                // for explanation of cutting the tail mass in two to get agreement with Excel two-tail
                double t68 = Math.abs(studentsT.inverseCumulativeProbability((1.0 - 0.6826) / 2.0));
                double t95 = Math.abs(studentsT.inverseCumulativeProbability((1.0 - 0.95) / 2));

                intMean = sumWtdRatios / weight;//  ' "Internal" error of wtd average

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

                MSWD = sums / nU;//  ' Mean square of weighted deviates
                intSigmaMean = Math.sqrt(1.0 / weight);

                // http://commons.apache.org/proper/commons-math/apidocs/org/apache/commons/math3/distribution/FDistribution.html
                FDistribution fdist = new FDistribution(nU, 1E9);
                probability = 1.0 - fdist.cumulativeProbability(MSWD);//     ChiSquare(.MSWD, (nU))
                intMeanErr95 = intSigmaMean * (double) (probability >= 0.3 ? 1.96
                        : t95 * Math.sqrt(MSWD));
                intErr68 = intSigmaMean * (double) (probability >= 0.3 ? 0.9998
                        : t68 * Math.sqrt(MSWD));

                extMean = 0.0;
                extMeanErr95 = 0.0;
                extMeanErr68 = 0.0;

                // need to find external uncertainty
                List<Double> yyList = new ArrayList<>();
                List<Double> iVarYList = new ArrayList<>();
                if ((probability < SQUID_MINIMUM_PROBABILITY) && (MSWD > 1.0)) {
                    // Find the MLE constant external variance
                    nN = 0;
                    for (int i = 0; i < nPts; i++) {
                        if (values[i] != 0.0) {
                            yyList.add(values[i]);
                            iVarYList.add(errors[i] * errors[i]);
                            nN++;
                        }
                    }
                    // resize arrays
                    yy = yyList.stream().mapToDouble(Double::doubleValue).toArray();
                    iVarY = iVarYList.stream().mapToDouble(Double::doubleValue).toArray();

                    // call secant method
                    double[] wtdExtRtsec = wtdExtRTSEC(0, 10.0 * intSigmaMean * intSigmaMean, yy, iVarY);

                    // check for failure
                    if (wtdExtRtsec[3] == 0.0) {
                        extMean = wtdExtRtsec[1];
                        extSigma = Math.sqrt(wtdExtRtsec[0]);

                        studentsT = new TDistribution(2 * nN - 2);
                        extMeanErr95 = Math.abs(studentsT.inverseCumulativeProbability((1.0 - 0.95) / 2.0)) * wtdExtRtsec[2];

                    } else if (MSWD > 4.0) {  //Failure of RTSEC algorithm because of extremely high MSWD
                        DescriptiveStatistics stats = new DescriptiveStatistics(yy);
                        extSigma = stats.getStandardDeviation();
                        extMean = stats.getMean();
                        extMeanErr95 = t95 * extSigma / Math.sqrt(nN);
                    } else {
                        extSigma = 0.0;
                        extMean = 0.0;
                        extMeanErr95 = 0.0;
                    }

                    extMeanErr68 = t68 / t95 * extMeanErr95;
                }

                if (canReject && (probability < SQUID_MINIMUM_PROBABILITY)) {
                    // GOSUB REJECT
                    double wtdAvg = 0.0;
                    if (extSigma != 0.0) {
                        wtdAvg = extMean;
                    } else {
                        wtdAvg = intMean;
                    }

                    //  reject outliers
                    int n0 = nN;

                    for (int i = 0; i < nPts; i++) {
                        if ((values[i] != 0.0) && (nN > 0.85 * nPts)) {   //  Reject no more than 30% of ratios
                            // Start rej. tolerance at 2-sigma, increase slightly each pass.
                            double pointError = 2.0 * Math.sqrt(errors[i] * errors[i] + extSigma * extSigma);
                            // 2-sigma error of point being tested
                            double totalError = Math.sqrt(pointError * pointError + (4.0 * extMeanErr68 * extMeanErr68));
                            // 1st-pass tolerance is 2-sigma; 2nd is 2.25-sigma; 3rd is 2.5-sigma.
                            double tolerance = (1.0 + (double)(count - 1.0) / 4.0) * totalError;
                            if (hardRej) {
                                tolerance = tolerance * 1.25;
                            }
                            // 1st-pass tolerance is 2-sigma; 2nd is 2.5-sigma; 3rd is 3-sigma...
                            q = values[i] - wtdAvg;

                            if ((Math.abs(q) > tolerance) && nN > 2) {
                                nN--;
                                wRejected[i][0] = values[i];
                                values[i] = 0.0;
                                wRejected[i][1] = errors[i];
                                errors[i] = 0.0;
                            } // check tolerance

                        } //  Reject no more than 30% of ratios
                    } // nPts loop               

                    reCalc = (nN < n0);
                } // canReject test
            } while (reCalc);

            if (canTukeys) { // March 2018 not finished as not sure where used
                System.arraycopy(values, 0, tbX, 0, nPts);
                
                double[] tukey = SquidMathUtils.tukeysBiweight(tbX, 6);
                biWtMean = tukey[0];
                biWtSigma = tukey[1];
                DescriptiveStatistics stats = new DescriptiveStatistics(tbX);
                double biWtErr95Median = stats.getPercentile(50);

                double median = median(tbX);
                double medianConf = medianConfLevel(nPts);
                double medianPlusErr = medianUpperLim(tbX) - median;
                double medianMinusErr = median - medianLowerLim(tbX);
            }

            // determine whether to return internal or external
            if (extMean != 0.0) {
                retVal = new double[][]{{
                    extMean,
                    extSigma,
                    extMeanErr68,
                    extMeanErr95,
                    MSWD,
                    probability,
                    1.0
                },
                // contains zero for each reject
                values
                };
            } else {
                retVal = new double[][]{{
                    intMean,
                    intSigmaMean,
                    intErr68,
                    intMeanErr95,
                    MSWD,
                    probability,
                    0.0
                },
                // contains zero for each reject
                values
                };
            }

        }

        return retVal;
    }

    /**
     * Using the secant method, find the root of a function WtdExtFunc thought
     * to lie between ExtVar1 and ExtVar2. The root, returned as WtdExtRTSEC, is
     * refined until its accuracy is +-xacc. Press et al, 1987, p. 250-251.
     *
     * @param extVar1
     * @param extVar2
     * @param x
     * @param intVar
     * @return double[4] where 0 = extVar, 1 = xBar, 2 = xBarSigma, 3 =
     * failedFlag where 0 = no failed, 1 = failed
     */
    public static double[] wtdExtRTSEC(double extVar1, double extVar2, double[] x, double[] intVar) {
        double[] retVal = new double[]{0.0, 0.0, 0.0, 1.0};

        int maxIt = 99;
        int maxD = 100;
        double xacc = 0.000000001;
        double facc = 0.0000001;
        double rts;
        double xl;

        double[] fL = wtdExtFunc(extVar1, x, intVar);
        double[] f = wtdExtFunc(extVar2, x, intVar);
        double lastf2 = 0.0;
        double lastf1 = 0.0;
        int j;
        
        double failedFlag = 0.0; // false as in did not fail

        if (Math.abs(f[0] - fL[0]) > 1e-10) {
            if (Math.abs(fL[0]) < Math.abs(f[0])) {
                rts = extVar1;
                xl = extVar2;
                double[] swap = f;
                f = fL;
                fL = swap;
            } else {
                rts = extVar2;
                xl = extVar1;
            }
            double dx;
            for (j = 0; j < maxIt; j++) {

                if (f[0] != fL[0]) {
                    dx = (xl - rts) * f[0] / (f[0] - fL[0]);
                    xl = rts;
                    int rct = 0;
                    fL = f;
                    double tmp = 0.0;

                    do {
                        tmp = rts + dx;
                        if (tmp < 0.0) {
                            dx = dx / 2.0;
                            rct += 1;
                        }
                    } while ((tmp < 0) && (rct <= maxD));

                    if (rct > maxD){
                        failedFlag = 1.0;
                        break;
                    }
                    
                    rts = tmp;
                    f = wtdExtFunc(rts, x, intVar);
                    if ((Math.abs(f[0]) >= facc) && (Math.abs(f[0]) != Math.abs(lastf2))) {
                        lastf2 = lastf1;
                        lastf1 = f[0];
                    } else {
                        break;
                    }
                }

            }

            retVal = new double[]{rts, f[1], f[2], failedFlag};
        }

        return retVal;
    }

    public static double[] wtdExtFunc(double extVar, double[] x, double[] intVar) {

        int n = x.length;
        double[] w = new double[n];
        double sumW = 0.0;
        double sumXW = 0.0;

        for (int i = 0; i < n; i++) {
            w[i] = 1.0 / (intVar[i] + extVar);
            sumW += w[i];
            sumXW += x[i] * w[i];
        }

        double xBar = sumXW / sumW;

        double sumW2resid2 = 0.0;
        for (int i = 0; i < n; i++) {
            double residual = x[i] - xBar;
            sumW2resid2 += w[i] * w[i] * residual * residual;
        }

        double ff = sumW2resid2 - sumW;

        double xBarSigma = Math.sqrt(Math.abs(1.0 / sumW));
        return new double[]{ff, xBar, xBarSigma};
    }

    /**
     * Calculates Confidence limit (%) of error on median
     *
     * @param n
     * @return double median confidence level
     */
    public static double medianConfLevel(int n) {
        double retVal;

        // Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
        double[] rockTable = new double[]{75.0, 87.8, 93.8, 96.9, 98.4, 93.0, 96.1, 97.9, 93.5, 96.1,
            97.8, 94.3, 96.5, 97.9, 95.1, 96.9, 93.6, 95.9, 97.3, 94.8, 96.5, 97.7, 95.7};

        double conf = 0.0;
        if (n > 25) {
            conf = 95.0;
        } else if (n > 2) {
            conf = rockTable[n - 3]; // because 0-based array
        }

        retVal = Utilities.roundedToSize(conf, 5);

        return retVal;
    }

    /**
     * Upper error on median of double[] values
     *
     * @param values
     * @return double median upper error
     */
    public static double medianUpperLim(double[] values) {
        double retVal;

        // Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
        double[] uR = new double[]{1, 1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 5, 6, 6, 6, 7, 7, 7, 8};

        int n = values.length;
        double u = 0.0;
        if (n > 25) {
            u = 0.5 * (n + 1.0 - 1.96 * Math.sqrt(n));
        } else if (n > 2) {
            u = uR[n - 3]; // because 0-based array
        }

        DescriptiveStatistics stats = new DescriptiveStatistics(values);
        retVal = stats.getPercentile((n - u + 1) / n); // vba = App.Large(v, u);

        return retVal;
    }

    /**
     * Lower error on median of double[] values
     *
     * @param values
     * @return
     */
    public static double medianLowerLim(double[] values) {
        double retVal;

        // Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
        double[] lR = new double[]{3, 4, 5, 6, 7, 07, 8, 9, 9, 10, 11, 11, 12, 13, 13, 14, 14, 15, 16, 16, 17, 18, 18};

        int n = values.length;
        double u = 0.0;
        double l = 0.0;
        if (n > 25) {
            u = 0.5 * (n + 1.0 - 1.96 * Math.sqrt(n));
            l = n + 1 - u;
        } else if (n > 2) {
            l = lR[n - 3]; // because 0-based array
        }

        DescriptiveStatistics stats = new DescriptiveStatistics(values);
        retVal = stats.getPercentile((n - l + 1) / n); // vba = App.Large(v, l);

        return retVal;

    }
}
