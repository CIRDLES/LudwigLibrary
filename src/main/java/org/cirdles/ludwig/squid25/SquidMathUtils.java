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

import static org.cirdles.ludwig.squid25.SquidConstants.SQUID_EPSILON;
import static org.cirdles.ludwig.squid25.SquidConstants.SQUID_TINY_VALUE;

/**
 * double implementations of Ken Ludwig's Squid.MathUtils VBA code for use with
 * Shrimp prawn files data reduction. Each public function returns a one-
 * dimensional array of double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/squid2.5Basic/MathUtils.bas" target="_blank">Squid.MathUtils</a>
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/Pub.bas" target="_blank">Isoplot.Pub</a>
 *
 * @author James F. Bowring
 */
public final class SquidMathUtils {

    private SquidMathUtils() {
    }

    /**
     * Ludwig specifies: Calculates Tukey's biweight estimator of location &
     * scale. Mean is a very robust estimator of "mean", Sigma is the robust
     * estimator of "sigma". These estimators converge to the true mean & true
     * sigma for Gaussian distributions, but are very resistant to outliers. The
     * lower the "Tuning" constant is, the more the tails of the distribution '
     * are effectively "trimmed" (& the more robust the estimators are against
     * outliers), with the price that more "good" data is disregarded. Data that
     * deviate from the "mean" greater that "Tuning" times the "standard
     * deviation" are assigned a weight of zero ('rejected'). Err95 is the 95%
     * confidence limit on Mean. Adapted & inferred from Hoaglin, Mosteller, &
     * Tukey, 1983, Understanding Robust & Exploratory Data Analysis: John Wiley
     * & Sons, pp. 341, 367, 376-378, 385-387, 423,& 425-427.
     *
     * @param values double[] array of values
     * @param tuningConstant integer 0 to 9
     * @return double[1][3] containing mean, 1-sigma absolute, 95% confidence
     * @throws ArithmeticException
     */
    public static double[] tukeysBiweight(double[] values, double tuningConstant)
            throws ArithmeticException {

        int iterationMax = 100;
        int iterationCounter = 0;

        int n = values.length;
        // initial mean is median
        double mean = Utilities.median(values);

        // initial sigma is median absolute deviation from mean = median (MAD)
        double deviations[] = new double[n];
        for (int i = 0; i < values.length; i++) {
            deviations[i] = Math.abs(values[i] - mean);
        }

        double sigma = Math.max(Utilities.median(deviations), SQUID_TINY_VALUE);

        double previousMean;
        double previousSigma;

        do {
            iterationCounter++;
            previousMean = mean;
            previousSigma = sigma;

            // init to zeroes
            double[] deltas = new double[n];
            double[] u = new double[n];
            double sa = 0.0;
            double sb = 0.0;
            double sc = 0.0;

            double tee = tuningConstant * sigma;

            for (int i = 0; i < n; i++) {
                deltas[i] = values[i] - mean;
                if (tee > Math.abs(deltas[i])) {
                    deltas[i] = values[i] - mean;
                    u[i] = deltas[i] / tee;
                    double uSquared = u[i] * u[i];

                    sa += Math.pow(deltas[i] * (1.0 - uSquared) * (1.0 - uSquared), 2);
                    sb += (1.0 - uSquared) * (1.0 - 5.0 * uSquared);
                    sc += u[i] * (1.0 - uSquared) * (1.0 - uSquared);
                }
            }

            sigma = Math.sqrt(sa * n) / Math.abs(sb);
            sigma = Math.max(sigma, SQUID_TINY_VALUE);
            mean = previousMean + (tee * sc / sb);

        } // both tests against epsilon must pass OR iterations top out
        // april 2016 Simon B discovered we need 101 iterations possible, hence the "<=" below
        while (((Math.abs(sigma - previousSigma) / sigma > SQUID_EPSILON)//
                || (Math.abs(mean - previousMean) / mean > SQUID_EPSILON))//
                && (iterationCounter <= iterationMax));

        if (sigma <= SQUID_TINY_VALUE) {
            sigma = 0.0;
        }

        double t;
        double w;
        switch (n) {
            case 1:
                t = 0.0;
                break;
            case 2:
            case 3:
                t = 47.2;
                break;
            case 4:
                t = 4.736;
                break;
            default:
                w = n - 4.358;
                t = 1.96 + (0.401 / Math.sqrt(w)) + (1.17 / w) + (0.0185 / (w * w));
        }

        double err95 = t * sigma / Math.sqrt(n);

        return new double[]{mean, sigma, err95};
    }
}
