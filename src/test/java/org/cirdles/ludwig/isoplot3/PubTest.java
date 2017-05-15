/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.cirdles.ludwig.isoplot3;

import org.cirdles.ludwig.squid25.SquidConstants;
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
public class PubTest {

    /**
     *
     */
    public PubTest() {
    }

    /**
     *
     */
    @BeforeClass
    public static void setUpClass() {
    }

    /**
     *
     */
    @AfterClass
    public static void tearDownClass() {
    }

    /**
     *
     */
    @Before
    public void setUp() {
    }

    /**
     *
     */
    @After
    public void tearDown() {
    }

    /**
     * Test of robustReg2 method, of class Pub.
     */
    @Test
    public void testRobustReg2() {
        System.out.println("robustReg2");
        double[] xValues = null;
        double[] yValues = null;
        double[] expResult = null;
//        double[] result = Pub.robustReg2(xValues, yValues);
//        assertArrayEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
    }

    /**
     * Test of ageNLE method, of class Pub.
     */
    @Test
    public void testAgeNLE() {
        System.out.println("ageNLE");
        double xVal = 0.0;
        double yVal = 0.0;
        double[][] covariance = null;
        double trialAge = 0.0;
        double[] expResult = null;
//        double[] result = Pub.ageNLE(xVal, yVal, covariance, trialAge);
//        assertArrayEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
    }

    /**
     * Test of concordiaTW method, of class Pub.
     *
     * Data and outcome taken from Squid2.5 workbook
     */
    @Test
    public void testConcordiaTW() {
        System.out.println("concordiaTW");
        double r238U_206Pb = 6.65756816656;
        double r238U_206Pb_1SigmaAbs = 6.65756816656 * 1.87624507122 / 100;
        double r207Pb_206Pb = 0.0552518706529;
        double r207Pb_206Pb_1SigmaAbs = 0.0552518706529 * 1.96293438298 / 100;
        double[] expResult = new double[]{8.140922087390351E8, 1.3673235207954608E7, 113.12698774576353, 0.0};
        double[] result = Pub.concordiaTW(r238U_206Pb, r238U_206Pb_1SigmaAbs, r207Pb_206Pb, r207Pb_206Pb_1SigmaAbs);
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);

        System.out.println("concordiaTW");
        r238U_206Pb = 6.91259509041;
        r238U_206Pb_1SigmaAbs = 6.91259509041 * 1.18363396151 / 100;
        r207Pb_206Pb = 0.0610677354475;
        r207Pb_206Pb_1SigmaAbs = 0.0610677354475 * 2.93532493394 / 100;
        expResult = new double[]{8.630954888631245E8, 9431402.438459378, 14.58887204116458, 1.3370175743965262E-4};
        result = Pub.concordiaTW(r238U_206Pb, r238U_206Pb_1SigmaAbs, r207Pb_206Pb, r207Pb_206Pb_1SigmaAbs);
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);

    }

    /**
     * Test of concordia method, of class Pub.
     */
    @Test
    public void testConcordia() {
        System.out.println("concordia");
        double r207Pb_235U = 0.0;
        double r207Pb_235U_1SigmaAbs = 0.0;
        double r206Pb_238U = 0.0;
        double r206Pb_238U_1SigmaAbs = 0.0;
        double rho = 0.0;
        double[] expResult = null;
//        double[] result = Pub.concordia(r207Pb_235U, r207Pb_235U_1SigmaAbs, r206Pb_238U, r206Pb_238U_1SigmaAbs, rho);
//        assertArrayEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
    }

    /**
     * Test of concordiaAges method, of class Pub.
     */
    @Test
    public void testConcordiaAges() {
        System.out.println("concordiaAges");
        double[] inputData = null;
        double[] expResult = null;
//        double[] result = Pub.concordiaAges(inputData);
//        assertArrayEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
    }

}
