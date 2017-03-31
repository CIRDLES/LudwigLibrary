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

import java.util.Arrays;

/**
 * double implementations of Ken Ludwig's Isoplot.Pub VBA code for use with
 * Shrimp prawn files data reduction. Each function returns a two dimensional
 * array of double.
 *
 * @see
 * https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/Pub.bas
 *
 * @author James F. Bowring
 */
public class IsoplotPub {

    /**
     * Robust linear regression using median of all pairwise slopes/intercepts,
     * after Hoaglin, Mosteller & Tukey, Understanding Robust & Exploratory Data
     * Analysis, John Wiley & Sons, 1983, p. 160, with errors from code in Rock
     * & Duffy, 1986 (Comp. Geosci. 12, 807-818), derived from Vugrinovich
     * (1981), J. Math. Geol. 13, 443-454). Has simple, rapid solution for
     * errors. Ludwig used flags and our approach is to do all the math and
     * return all possible values available as if those flags were true.
     *
     * @param xValues double [] array length n
     * @param yValues double [] array length n
     * @return double[1][9] containing slope, lSlope, uSlope, yInt, xInt, lYint,
     * uYint, lXint, uXint
     */
    public static double[][] robustReg2(double[] xValues, double[] yValues) {

        double[][] retVal = new double[][]{{0, 0, 0}};

        // check precondition of same size xValues and yValues and at least 3 points
        int n = xValues.length;
        if ((n == yValues.length) && n > 2) {
            // proceed
            double[][] slopeCalcs = IsoplotRobustReg.getRobSlope(xValues, yValues);
            double slope = slopeCalcs[0][0];
            double yInt = slopeCalcs[0][1];
            double xInt = slopeCalcs[0][2];

            double[] slp = slopeCalcs[1];
            Arrays.sort(slp);
            double[] yInter = slopeCalcs[2];
            Arrays.sort(yInter);
            double[] xInter = slopeCalcs[3];
            Arrays.sort(xInter);

            double[][] conf95Calcs = IsoplotRobustReg.conf95(n, slp.length);
            // reduce indices by 1 to zero-based - this did not work but keeping them did
            // TODO: understand why - probably integer division related
            int lwrInd = (int) conf95Calcs[0][0] - 0;
            int upprInd = (int) conf95Calcs[0][1] - 0;

            double lSlope = slp[lwrInd];
            double uSlope = slp[upprInd];

            double lYint = yInter[lwrInd];
            double uYint = yInter[upprInd];

            double lXint = xInter[lwrInd];
            double uXint = xInter[upprInd];

            retVal = new double[][]{{slope, lSlope, uSlope, yInt, xInt, lYint, uYint, lXint, uXint}};
        }

        return retVal;
    }

}
