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

import static org.cirdles.ludwig.squid25.Utilities.median;

/**
 * double implementations of Ken Ludwig's Squid2.5.Resistant VBA code for use
 * with Shrimp prawn files data reduction. Each public function returns a one-
 * dimensional array of double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/squid2.5Basic/Resistant.bas" target="_blank">Squid.Resistant</a>
 *
 * @author James F. Bowring
 */
public class Resistant {

    private Resistant() {
    }

    /**
     * Determine the Median Absolute Deviation (MAD) from the median for the
     * first N values in vector xValues.
     *
     * March 2018 Bodorkos and Bowring note that this operation is flawed when used on
     * uncertainties because it does not consider the values related to the
     * uncertainties. The only way to replicate Squid2.5 results is to input
     * 1sigma percent uncertainties rather than absolute uncertainties.
     *
     * @param xValues
     * @return double[1] where [0] = Median Absolute Deviation
     */
    public static double[] fdNmad(double[] xValues) {
        double[] retVal = new double[]{0.0};
        int n = xValues.length;

        if (n > 0) {
            int nN = Math.max(3, n);
            double[] yR2 = new double[n];

            double median = median(xValues);

            for (int i = 0; i < n; i++) {
                yR2[i] = Math.pow(xValues[i] - median, 2.0);
            }

            double medianyR2 = median(yR2);

            retVal = new double[]{1.4826 * (1.0 + 5.0 / (nN - 2)) * Math.sqrt(medianyR2)};
        }

        return retVal;
    }

}
