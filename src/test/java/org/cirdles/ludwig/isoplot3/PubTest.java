/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.cirdles.ludwig.isoplot3;

import org.cirdles.ludwig.squid25.SquidConstants;
import org.cirdles.ludwig.squid25.Utilities;
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

        System.out.println("concordiaTW - age calcs impossible");
        r238U_206Pb = 0.13527488;
        r238U_206Pb_1SigmaAbs = 0.0055755;
        r207Pb_206Pb = 0.1068302;
        r207Pb_206Pb_1SigmaAbs = 0.0018455;
        expResult = new double[]{0.0, 0.0, 0.0, 0.0};
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

    /**
     * Test of pb76 method, of class Pub.
     */
    @Test
    public void testPb76() {
        System.out.println("pb76");
        double age = 564738561.835384;
        double[] expResult = new double[]{0.05893615898302093};
        double[] result = Pub.pb76(age);
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of age7corrWithErr method, of class Pub.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_original_frozen.xls:
     *
     * this method combines Squid2.5's Age7Corr and AgeEr7Corr
     *
     * Column header: 207corr206Pb/238UAge Column header: 1σ err (to the right
     * of 207corr206Pb/238UAge)
     *
     * Usage: "=Age7Corr([Total206Pb/238U],[207/206],sComm0_76)"
     *
     * Usage:
     * "=AgeEr7Corr([207corr206Pb/238UAge],[Total206Pb/238U],[Total206Pb/238U
     * %err]/100*[Total206Pb/238U],[207/206],[207/206
     * %err]/100*[207/206],sComm0_76,0)"
     */
    @Test
    public void testAge7corrWithErr() {
        System.out.println("age7corrWithErr");
        double totPb6U8 = 0.0915928758248389;
        double totPb6U8err = 0.945106721698445 / 100.0 * 0.0915928758248389;
        double totPb76 = 0.0592518351787661;
        double totPb76err = 0.688011049802507 / 100.0 * 0.0592518351787661;
        double[] expResult = new double[]{564738561.835384, 5212379.83636884};
        double[] result = Pub.age7corrWithErr(totPb6U8, totPb6U8err, totPb76, totPb76err);
        assertEquals(Utilities.roundedToSize(expResult[0], 14), Utilities.roundedToSize(result[0], 14), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 14), Utilities.roundedToSize(result[1], 14), SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of agePb76WithErr method, of class Pub.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_original_frozen.xls:
     *
     * this method combines Squid2.5's AgePb76 and AgeErPb76
     *
     * Column header: 4-corr207Pb/206Pbage
     *
     * Column header: ±1σ (to the right of 4-corr207Pb/206Pbage)
     *
     * Usage: "=AgePb76([4-corr207Pb/206Pb])"
     *
     * Usage: "=AgeErPb76([4-corr207Pb/206Pb],[4-corr207Pb/206Pb
     * %err]/100*[4-corr207Pb/206Pb])"
     */
    @Test
    public void testAgePb76WithErr() {
        System.out.println("agePb76WithErr");
        double pb76rad = 0.0589848267158073;
        double pb76err = 0.720404321194942 / 100.0 * 0.0589848267158073;
        double[] expResult = new double[]{566536050.577858, 15685433.1235601};
        double[] result = Pub.agePb76WithErr(pb76rad, pb76err);
        assertEquals(Utilities.roundedToSize(expResult[0], 14), Utilities.roundedToSize(result[0], 14), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 14), Utilities.roundedToSize(result[1], 14), SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of age8corrWithErr method, of class Pub.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_original_frozen.xls:
     *
     * this method combines Squid2.5's Age8Corr and AgeEr8Corr
     *
     * Column header: 208corr206Pb/238UAge
     *
     * Column header: 1σ err (to the right of 208corr206Pb/238UAge)
     *
     * Usage:
     * "=Age8Corr([Total206Pb/238U],[Total208Pb/232Th],[232Th/238U],1/sComm0_86)"
     *
     * Usage:
     * "=AgeEr8Corr([208corr206Pb/238UAge],[Total206Pb/238U],[Total206Pb/238U
     * %err]/100*[Total206Pb/238U],[Total208Pb/232Th],[Total208Pb/232Th
     * %err]/100*[Total208Pb/232Th],[232Th/238U],0,1/sComm0_86,0)"
     */
    @Test
    public void testAge8corrWithErr() {
        System.out.println("age8corrWithErr");
        double totPb6U8 = 0.0915928758248389;
        double totPb6U8err = 0.945106721698445 / 100.0 * 0.0915928758248389;
        double totPb8Th2 = 0.0280531210114337;
        double totPb8Th2err = 1.69397823766887 / 100.0 * 0.0280531210114337;
        double th2U8 = 0.27150941987617;
        double th2U8err = 0.0;
        double[] expResult = new double[]{565183115.935407, 5332237.68794901};
        double[] result = Pub.age8corrWithErr(totPb6U8, totPb6U8err, totPb8Th2, totPb8Th2err, th2U8, th2U8err);
        assertEquals(Utilities.roundedToSize(expResult[0], 14), Utilities.roundedToSize(result[0], 14), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 14), Utilities.roundedToSize(result[1], 14), SquidConstants.SQUID_EPSILON);
    }

}
