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

import org.apache.commons.math3.distribution.FDistribution;
import org.apache.commons.math3.distribution.TDistribution;

/**
 * double implementations of Ken Ludwig's Isoplot.Pub VBA code for use with
 * Shrimp prawn files data reduction. Each function returns an array of double.
 *
 * @see
 * https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/Pub.bas
 *
 * @author James F. Bowring
 */
public class Means {

    private Means() {
    }

    /**
     * Simplified version of Ludwig's WeightedAverage - does not provide for
     * external uncertainty, rejection or Tukeys math.
     *
     * @param values as double[] with length nPts
     * @param errors as double[] with length nPts
     * @return double[]{MSWD, intSigmaMean, intErr68, intMeanErr95, probability}
     */
    public static double[] weightedAverage(double[] values, double[] errors) {

        double[] retVal = new double[]{0, 0, 0};

        // check precondition of same size xValues and yValues and at least 3 points
        int nPts = values.length;
        if ((nPts == errors.length) && nPts > 2) {
            // proceed
            double[] inverseVar = new double[nPts];
            double[] wtdResid = new double[nPts];

            for (int i = 0; i < nPts; i++) {
                inverseVar[i] = Math.pow(errors[i], 2);
            }

            double weight = 0.0;
            double sumWtdRatios = 0.0;
            double q = 0.0;

            for (int i = 0; i < nPts; i++) {
                if (values[i] * errors[i] != 0.0) {
                    weight += inverseVar[i];
                    sumWtdRatios += inverseVar[i] * values[i];
                    q += inverseVar[i] * Math.pow(values[i], 2);
                }
            }

            int nU = nPts - 1;// ' Deg. freedom
            TDistribution studentsT = new TDistribution(nU);
            double t95 = studentsT.inverseCumulativeProbability(95.0);

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

            double MSWD = sums / nU;//  ' Mean square of weighted deviates
            double intSigmaMean = Math.sqrt(1.0 / weight);

            // http://commons.apache.org/proper/commons-math/apidocs/org/apache/commons/math3/distribution/FDistribution.html
            FDistribution fdist = new FDistribution(nU, 1E9);
            double probability = 1.0 - fdist.cumulativeProbability(MSWD);//     ChiSquare(.MSWD, (nU))
            double intMeanErr95 = intSigmaMean * (double) (probability >= 0.3 ? 1.96
                    : t95 * Math.sqrt(MSWD));
            double intErr68 = intSigmaMean * (double) (probability >= 0.3 ? 0.9998
                    : studentsT.inverseCumulativeProbability(68.26) * Math.sqrt(MSWD));

            //TODO: (VBA line 508) Resolve how to specify minProb for next two sections of code
            // at this point we have the basic weighted mean info
            // referenced using: ww As wWtdAver
            // MSWD            
            // IntMeanErr2sigma >> we supply intSigmaMean as 1 sigma abs
            // Probability
            // IntMeanErr95
            // -------
            // modifying to return this array:
            retVal = new double[]{MSWD, intSigmaMean, intErr68, intMeanErr95, probability};

            // Ext2Sigma
            // IntMean
            // ExtMean     
            // ChosenMean
            // ChosenErr
            // ChosenErrPercent
            // BiwtMean       
            // ExtMeanErr95
            // ExtMeanErr68
        }

        return retVal;
    }
}
