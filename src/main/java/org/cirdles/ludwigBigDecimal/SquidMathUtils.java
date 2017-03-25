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
package org.cirdles.ludwigBigDecimal;

import java.math.BigDecimal;
import java.math.MathContext;
import static org.cirdles.ludwigBigDecimal.BigDecimalCustomAlgorithms.bigDecimalSqrtBabylonian;
import static org.cirdles.squid.SquidConstants.SQUID_TINY_VALUE;
import org.cirdles.utilities.Utilities;

/**
 * BigDecimal implementations of Ken Ludwig's Squid VBA code for use with Shrimp
 * prawn files data reduction. Each function returns a two dimensional array of
 * BigDecimal.
 *
 * @see
 * https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/squid2.5Basic/MathUtils.bas
 * @see
 * https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/Pub.bas
 *
 * @author James F. Bowring
 */
public final class SquidMathUtils {

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
     * @return BigDecimal[1][3] containing mean, 1-sigma absolute, 95%
     * confidence
     * @throws ArithmeticException
     */
    public static BigDecimal[][] tukeysBiweight(double[] values, double tuningConstant)
            throws ArithmeticException {
        // guarantee termination
        BigDecimal epsilon = BigDecimal.ONE.movePointLeft(10);
        int iterationMax = 100;
        int iterationCounter = 0;

        int n = values.length;
        // initial mean is median
        BigDecimal mean = new BigDecimal(Utilities.median(values));

        // initial sigma is median absolute deviation from mean = median (MAD)
        double deviations[] = new double[n];
        for (int i = 0; i < values.length; i++) {
            deviations[i] = StrictMath.abs(values[i] - mean.doubleValue());
        }
        BigDecimal sigma = new BigDecimal(Utilities.median(deviations)).max(BigDecimal.valueOf(SQUID_TINY_VALUE));

        BigDecimal previousMean;
        BigDecimal previousSigma;

        do {
            iterationCounter++;
            previousMean = mean;
            previousSigma = sigma;

            // init to zeroes
            BigDecimal[] deltas = new BigDecimal[n];
            BigDecimal[] u = new BigDecimal[n];
            BigDecimal sa = BigDecimal.ZERO;
            BigDecimal sb = BigDecimal.ZERO;
            BigDecimal sc = BigDecimal.ZERO;

            BigDecimal tee = new BigDecimal(tuningConstant).multiply(sigma);

            for (int i = 0; i < n; i++) {
                deltas[i] = new BigDecimal(values[i]).subtract(mean);
                if (tee.compareTo(deltas[i].abs()) > 0) {
                    deltas[i] = new BigDecimal(values[i]).subtract(mean);
                    u[i] = deltas[i].divide(tee, MathContext.DECIMAL128);
                    BigDecimal uSquared = u[i].multiply(u[i]);
                    sa = sa.add(deltas[i].multiply(BigDecimal.ONE.subtract(uSquared).pow(2)).pow(2));
                    sb = sb.add(BigDecimal.ONE.subtract(uSquared).multiply(BigDecimal.ONE.subtract(new BigDecimal(5.0).multiply(uSquared))));
                    sc = sc.add(u[i].multiply(BigDecimal.ONE.subtract(uSquared).pow(2)));
                }
            }

            sigma = bigDecimalSqrtBabylonian(sa.multiply(new BigDecimal(n))).divide(sb.abs(), MathContext.DECIMAL128);
            sigma = sigma.max(BigDecimal.valueOf(SQUID_TINY_VALUE));
            mean = previousMean.add(tee.multiply(sc).divide(sb, MathContext.DECIMAL128));

        } // both tests against epsilon must pass OR iterations top out
        // april 2016 Simon B discovered we need 101 iterations possible, hence the "<=" below
        while (((sigma.subtract(previousSigma).abs().divide(sigma, MathContext.DECIMAL128).compareTo(epsilon) > 0)//
                || mean.subtract(previousMean).abs().divide(mean, MathContext.DECIMAL128).compareTo(epsilon) > 0)//
                && (iterationCounter <= iterationMax));

        if (sigma.compareTo(BigDecimal.valueOf(SQUID_TINY_VALUE)) <= 0) {
            sigma = BigDecimal.ZERO;
        }

        BigDecimal t;
        BigDecimal w;
        switch (n) {
            case 1:
                t = BigDecimal.ZERO;
                break;
            case 2:
            case 3:
                t = BigDecimal.valueOf(47.2);
                break;
            case 4:
                t = BigDecimal.valueOf(4.736);
                break;
            default:
                w = BigDecimal.valueOf(n).subtract(BigDecimal.valueOf(4.358));
                t = BigDecimal.valueOf(1.96)//
                        .add(BigDecimal.valueOf(0.401)//
                                .divide(bigDecimalSqrtBabylonian(w), MathContext.DECIMAL128))//
                        .add(BigDecimal.valueOf(1.17).divide(w, MathContext.DECIMAL128))//
                        .add(BigDecimal.valueOf(0.0185).divide(w.multiply(w), MathContext.DECIMAL128));
        }

        BigDecimal err95 = t.multiply(sigma)
                .divide(bigDecimalSqrtBabylonian(BigDecimal.valueOf(n)), MathContext.DECIMAL128);

        return new BigDecimal[][]{{mean, sigma, err95}};
    }

}
