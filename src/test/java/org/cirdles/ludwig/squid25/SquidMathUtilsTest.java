/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.cirdles.ludwig.squid25;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author James F. Bowring
 */
public class SquidMathUtilsTest {

    public SquidMathUtilsTest() {
    }

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() {
    }

    @After
    public void tearDown() {
    }

    @Test
    public void testTukeysBiweight() {
        System.out.println("calculateTukeysBiweightMean");
        String name = "";
        double tuningConstant = 9.0;
        double[] values = {2494, 2524, 2455, 2427, 2396, 2545, 2483, 2436, 2548, 2619};
        // oracle by Squid
        double expValue = 2492.51139904333;
        double expSigma = 206.312497307535;
        double[] result = SquidMathUtils.tukeysBiweight(values, tuningConstant);
        double value = result[0];
        double sigma = result[1];
//        assertEquals(expValue, value, SquidConstants.SQUID_EPSILON);
//        assertEquals(expSigma, sigma, SquidConstants.SQUID_EPSILON);

        tuningConstant = 9.0;
        values = new double[]{0.302198828429556,
            0.300788957475996,
            0.297713166278977,
            0.297778760994429,
            0.297483827242158};
        // oracle by Simon Bodokos by hand
        expValue = 0.297659637730707;
        expSigma = 0.000166784902889577;
        result = SquidMathUtils.tukeysBiweight(values, tuningConstant);
        value = result[0];
        sigma = result[1];
        assertEquals(expValue, value, SquidConstants.SQUID_EPSILON);
        assertEquals(expSigma, sigma, SquidConstants.SQUID_EPSILON);

        System.out.println("calculateTukeysBiweightMean test 2");
        values = new double[]{0.0701269311,
            0.07015861952,
            0.07003238026,
            0.07104611923,
            0.06736331324,
            0.07342762035,
            0.07308375966,
            0.06964566169,
            0.07073497843,
            0.06920469862,
            0.07288619162,
            0.06970642527,
            0.07029836374,
            0.06661243075,
            0.06820837433,
            0.07088312165,
            0.06715154317,
            0.07189521506,
            0.06629794724,
            0.06658774174,
            0.07016820432};

        expValue = 0.0697911642765795;
        expSigma = 0.00212084209983744;
        result = SquidMathUtils.tukeysBiweight(values, tuningConstant);
        value = result[0];
        sigma = result[1];
        assertEquals(expValue, value, SquidConstants.SQUID_EPSILON);
        assertEquals(expSigma, sigma, SquidConstants.SQUID_EPSILON);

    }

}
