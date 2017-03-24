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

import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import static org.cirdles.squid.SquidConstants.SQUID_TINY_VALUE;

/**
 * From Ken Ludwig's Squid VBA code for use with Shrimp prawn files data
 * reduction. Note code extracted by Simon Bodorkos in emails to bowring
 * Feb.2016
 *
 * @author James F. Bowring
 */
public final class TukeysBiweight {

    public static double[][] tukeysBiweight(double[] values, double tuningConstant) {
        // guarantee termination
        double epsilon = 1e-10;
        int iterationMax = 100;
        int iterationCounter = 0;

        int n = values.length;
        // initial mean is median
        double mean = calculateMedian(values);

        // initial sigma is median absolute deviation from mean = median (MAD)
        double deviations[] = new double[n];
        for (int i = 0; i < values.length; i++) {
            deviations[i] = Math.abs(values[i] - mean);
        }

        double sigma = Math.max(calculateMedian(deviations), SQUID_TINY_VALUE);

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

                    sa = sa + Math.pow(deltas[i] * (1.0 - uSquared) * (1.0 - uSquared), 2);
                    sb = sb + (1.0 - uSquared) * (1.0 - 5.0 * uSquared);
                    sc = sc + u[i] * (1.0 - uSquared) * (1.0 - uSquared);
                }
            }

            sigma = Math.sqrt(sa * n) / Math.abs(sb);
            sigma = Math.max(sigma, SQUID_TINY_VALUE);
            mean = previousMean + (tee * sc / sb);

        } // both tests against epsilon must pass OR iterations top out
        // april 2016 Simon B discovered we need 101 iterations possible, hence the "<=" below
        while (((Math.abs(sigma - previousSigma) / sigma > epsilon)//
                || (Math.abs(mean - previousMean) / mean > epsilon))//
                && (iterationCounter <= iterationMax));

        return new double[][]{{mean, sigma}};
    }

    /**
     * Calculates arithmetic median of array of doubles.
     *
     * @pre values has one element
     * @param values
     * @return
     */
    public static double calculateMedian(double[] values) {
        double median;

        // enforce precondition
        if (values.length == 0) {
            median = 0.0;
        } else {
            DescriptiveStatistics stats = new DescriptiveStatistics();

            // Add the data from the array
            for (int i = 0; i < values.length; i++) {
                stats.addValue(values[i]);
            }
            median = stats.getPercentile(50);
        }

        return median;
    }
}
