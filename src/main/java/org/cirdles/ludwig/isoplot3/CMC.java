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

import org.cirdles.ludwig.squid25.SquidConstants;

/**
 * Double implementations of Ken Ludwig's Isoplot.cMC VBA code for use with
 * Shrimp prawn files data reduction. Each public function returns an array of
 * double.
 *
 * @see
 * <a href="https://raw.githubusercontent.com/CIRDLES/LudwigLibrary/master/vbaCode/isoplot3Basic/cMC.bas" target="_blank">Isoplot.cMC</a>
 * @author James F. Bowring
 */
public class CMC {

    private CMC() {
    }

    /**
     * Ludwig's comments: Convert T-W concordia data to Conv., or vice-versa eg
     * 238/206-207/206[-204/206] to/from 207/235-206/238[-204/238]. This
     * implementation is for 2D only and excludes the optional 3D conversion.
     *
     * @param ratioX TW 238/206 or WC 207/235
     * @param ratioX_1SigmaAbs 1-sigma uncertainty for ratioX
     * @param ratioY TW 207/206 or WC 206/238
     * @param ratioY_1SigmaAbs 1-sigma uncertainty for ratioY
     * @param rhoXY correlation coefficient between uncertainties in ratioX and
     * ratioY
     * @param inTW true if data is Terra Waserburg (TW), false if Wetherill
     * Conconcordia
     * @return double[5] of conversions: ratioX, ratioX_1SigmaAbs, ratioY,
     * ratioY_1SigmaAbs, rhoXY
     */
    public static double[] concConvert(
            double ratioX, double ratioX_1SigmaAbs, double ratioY, double ratioY_1SigmaAbs, double rhoXY, boolean inTW) {
            return concConvert(ratioX, ratioX_1SigmaAbs, ratioY, ratioY_1SigmaAbs, rhoXY, inTW, SquidConstants.uRatio);
    }

    /**
     * Ludwig's comments: Convert T-W concordia data to Conv., or vice-versa eg
     * 238/206-207/206[-204/206] to/from 207/235-206/238[-204/238]. This
     * implementation is for 2D only and excludes the optional 3D conversion.
     *
     * @param ratioX TW 238/206 or WC 207/235
     * @param ratioX_1SigmaAbs 1-sigma uncertainty for ratioX
     * @param ratioY TW 207/206 or WC 206/238
     * @param ratioY_1SigmaAbs 1-sigma uncertainty for ratioY
     * @param rhoXY correlation coefficient between uncertainties in ratioX and
     * ratioY
     * @param inTW true if data is Terra Waserburg (TW), false if Wetherill
     * Conconcordia
     * @param uRatio
     * @return double[5] of conversions: ratioX, ratioX_1SigmaAbs, ratioY,
     * ratioY_1SigmaAbs, rhoXY
     */
    public static double[] concConvert(
            double ratioX, 
            double ratioX_1SigmaAbs, 
            double ratioY, 
            double ratioY_1SigmaAbs, 
            double rhoXY, 
            boolean inTW,
            double uRatio) {
        double[] retVal;

        double xP = Math.abs(ratioX_1SigmaAbs / ratioX);
        double yP = Math.abs(ratioY_1SigmaAbs / ratioY);

        double xP2 = xP * xP;
//        double yP2 = yP * yP;

        double abP = 0.0;
        double a = 0.0;
        double b = 0.0;
        double aP = 0.0;
        double bP = 0.0;
        // rho outside [-1...1] as default
        double rAB = 2.0;

        try {
            abP = Math.sqrt(xP2 + yP * yP - 2 * xP * yP * rhoXY);
        } catch (Exception e) {
            abP = 0.0;
        }
        if (abP >= 0.0) {

            if (inTW) {
                a = ratioY / ratioX * uRatio; // 207/235
                b = 1.0 / ratioX; // 206/238
                if (abP != 0.0) {
                    aP = abP;
                    bP = xP;
                    rAB = (xP - yP * rhoXY) / abP;
                }
            } else {
                a = 1.0 / ratioY;  // 238/206
                b = ratioX / ratioY / uRatio; // 207/206
                aP = yP;
                bP = abP;
                if (abP != 0.0) {
                    rAB = (yP - xP * rhoXY) / abP;
                }
            }
        }

        if (Math.abs(rAB) > 1.0) {
            // bad uncertainties
            retVal = new double[]{a, 0.0, b, 0.0, rAB};
        } else {
            retVal = new double[]{a, aP * a, b, bP * b, rAB};
        }

        return retVal;
    }

}
