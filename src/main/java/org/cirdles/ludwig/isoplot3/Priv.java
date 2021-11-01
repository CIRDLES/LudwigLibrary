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

import static org.cirdles.ludwig.squid25.Utilities.median;

/**
 * double implementations of Ken Ludwig's Isoplot.Priv VBA code for use with
 * Shrimp prawn files data reduction. Each public function returns an array of
 * double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/Priv.bas" target="_blank">Isoplot.UPb</a>
 * @author James F. Bowring
 */
public class Priv {

    private Priv() {
    }

    /**
     * Determine the Median Absolute Deviation (MAD) from the median for the
     * first N values in vector xValues with median MedianVal.
     *
     * @param xValues
     * @param medianVal
     * @return double[2] where [0] = Median Absolute Deviation and [1] = 95%
     * uncertainty
     */
    public static double[] getMAD(double[] xValues, double medianVal) {
        double[] retVal = new double[]{0.0, 0.0};
        int n = xValues.length;

        if (n > 0) {
            double[] absDev = new double[n];
            for (int i = 0; i < n; i++) {
                absDev[i] = StrictMath.abs(xValues[i] - medianVal);

            }

            double madd = median(absDev);

            // KRL-derived numerical approx., valid for normal distr. w. Tuning=9
            double[] tStar = new double[]{0.0, 0.0, 12.7, 15.3, 3.54 / StrictMath.sqrt(n) - 3.92 / n + 70.9 / (n * n) - 60.6 / (n * n * n)};

            double err95 = ((n < 4) ? tStar[n] : tStar[4]) * madd;

            retVal = new double[]{madd, err95};
        }

        return retVal;
    }
}