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

/**
 * Double implementations of Ken Ludwig's Isoplot.U_2 VBA code for use with
 * Shrimp prawn files data reduction. Each public function returns an array of
 * double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/U_2.bas" target="_blank">Isoplot.U_2</a>
 * @author James F. Bowring
 */
public class U_2 {

    private U_2() {
    }

    /**
     * Ludwig: Invert a symmetric 2x2 matrix.
     *
     * @param xx matrix element 0,0
     * @param yy matrix element 1,1
     * @param xy matrix element 0,1
     * @return double[3] containing inverted iXX, iYY, iXY
     */
    public static double[] inv2x2(double xx, double yy, double xy) {//isoplot3.U_2.bas
        double[] retVal = new double[]{0.0, 00, 0.0};

        double determinant = xx * yy - xy * xy;
        if (determinant != 0.0) {
            retVal = new double[]{yy / determinant, xx / determinant, -xy / determinant};
        }

        return retVal;
    }

}
