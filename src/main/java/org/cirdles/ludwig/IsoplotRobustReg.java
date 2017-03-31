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

import java.util.Random;
import static org.cirdles.squid.SquidConstants.SQUID_EPSILON;
import org.cirdles.utilities.Utilities;

/**
 * double implementations of Ken Ludwig's Isoplot.RobustReg VBA code for use
 * with Shrimp prawn files data reduction. Each public function returns a two
 * dimensional array of double.
 *
 * @see
 * https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/RobustReg.bas
 * @author James F. Bowring
 */
public class IsoplotRobustReg {

    /**
     * Calculates slope and intercepts for a set of points.  Does not implement
     * Ludwig's outlier rejection.
     *
     * @param xValues
     * @param yValues
     * @return double[4][] with row 0 containing slope, y-intercept,
     * x-intercept, row 1 containing slope array, row 2 containing y-intercept
     * array, and row 3 containing x-intercept array.
     *
     * @throws ArithmeticException
     */
    protected static double[][] getRobSlope(double[] xValues, double[] yValues)
            throws ArithmeticException {

        double[][] retVal = new double[][]{{0, 0, 0}};
        // check precondition of same size xValues and yValues and at least 3 points
        int n = xValues.length;
        if ((n == yValues.length) && n > 2) {
            // proceed
            Random random = new Random();
            int m = n * (n - 1) / 2;
            double[] slp = new double[m];
            double[] xInter = new double[m];
            double[] yInter = new double[m];
            int k = -1;

            for (int i = 0; i < (n - 1); i++) {
                for (int j = i + 1; j < n; j++) {
                    double vs = 0.0;
                    double vy = 0.0;
                    k++;
                    if (xValues[i] != xValues[j]) {
                        vs = (yValues[j] - yValues[i]) / (xValues[j] - xValues[i]);
                    }

                    vs += (0.5 - random.nextDouble()) * SQUID_EPSILON;
                    slp[k] = vs;

                    vy = yValues[i] - vs * xValues[i] + (0.5 - random.nextDouble()) * SQUID_EPSILON;
                    yInter[k] = vy;
                    xInter[k] = -vy / vs;
                } // end inner j loop
            } // end outer i loop

            double slope = Utilities.median(slp);
            double yInt = Utilities.median(yInter);
            double xInt = Utilities.median(xInter);

            retVal = new double[][]{{slope, yInt, xInt}, slp, yInter, xInter};
        }

        return retVal;
    }

    /**
     * Finds sorting-indexes to get 95%-conf. limits for repeated pairwise
     * slope/inter medians using algorithm coded in Rock & Duffy, 1986 (Comp.
     * Geosci. 12, 807-818), derived from Vugorinovich (1981, J. Math. Geol. 13,
     * 443-454).
     *
     * @param nPts
     * @param nMedians
     * @return
     */
    protected static double[][] conf95(int nPts, int nMedians) {

        int lowInd = 1;
        int upprInd = nPts;

        if (nPts > 4) {
            int star95;
            if (nPts < 14) {
                String c$ = "081012141719222528";
                star95 = Integer.parseInt(c$.substring(2 * nPts - 9, 2 * nPts - 7));
            } else {
                double x = Math.sqrt(nPts * (nPts - 1.0) * (2.0 * nPts + 5.0) / 18.0);
                star95 = (int) (1.96 * x);
            }
            lowInd = (nMedians - star95) / 2;
            upprInd = (nMedians + star95) / 2;
        }

        return new double[][]{{lowInd, upprInd}};
    }
}
