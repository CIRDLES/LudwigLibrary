/*
 * Copyright 2016 CIRDLES
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.cirdles.ludwig.squid25;

/**
 * Defines constants used throughout Ludwig's Squid VBA code.
 *
 * @author James F. Bowring
 */
public final class SquidConstants {

    /**
     *
     */
    public static final double SQUID_TINY_VALUE = 1e-30;

    public static final double SQUID_VERY_SMALL_VALUE = 1e-16;

    /**
     *
     */
    public static final double SQUID_EPSILON = 1e-10;

    /**
     *
     */
    public static final double SQUID_ERROR_VALUE = -9.87654321012346;

    /**
     *
     */
    public static final double MAXEXP = 709;

    /**
     *
     */
    public static final double MAXLOG = 1E+308;

    /**
     *
     */
    public static final double MINLOG = 1E-307;

    /**
     *
     */
    public static final double SQUID_UPPER_LIMIT_1_SIGMA_PERCENT = 9999.0;

    /**
     *
     */
    public static final double SQUID_MINIMUM_PROBABILITY = 0.05;

    // March 2017 place holders until constants models are implemented
    /**
     *
     */
    public static final double lambda238 = 1.55125e-10; // Ludwig uses 1e6 for MA

    /**
     *
     */
    public static final double lambda235 = 9.8485e-10;
    public static final double lambda232 = 4.9475e-11;

    /**
     *
     */
    public static final double uRatio = 137.88;

    /**
     *
     */
    public static final double badAge = -1.23456789; // Ludwig calls it BadT

    public static final double sComm0_64 = 18.053;
    public static final double sComm0_76 = 0.8637;
    public static final double sComm0_86 = 2.0971;
    public static final double sComm0_74 = 15.5923761;// (i.e. sComm0_64 * sComm0_76)
    public static final double sComm0_84 = 37.8589463;// (i.e. sComm0_64 * sComm0_86)

    public static final double PRESENT_238U235U = 137.88;

}
