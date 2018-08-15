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
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_orig_2017-07-17_frozen.xls:
     */
    @Test
    public void testRobustReg2() {
        System.out.println("robustReg2");
        double[] xValues = new double[]{
            1.817367664570000,
            1.797514664390000,
            1.812274675390000,
            1.817512881760000,
            1.826225765800000,
            1.813529053720000,
            1.821474271250000,
            1.805010623890000,
            1.813475607270000,
            1.829050148160000,
            1.819181881460000,
            1.821059779300000,
            1.816694626730000,
            1.809579671480000,
            1.809251581380000,
            1.801146788710000,
            1.814803586650000,
            1.832461520260000,
            1.817673161160000,
            1.814281811960000,
            1.822463451570000,
            1.811327095690000,
            1.814736103060000};
        double[] yValues = new double[]{
            -1.895899338360000,
            -1.926677362850000,
            -1.884482279160000,
            -1.892839687350000,
            -1.875440239210000,
            -1.901411644860000,
            -1.881311669340000,
            -1.935191031100000,
            -1.898809350600000,
            -1.874550919730000,
            -1.907383108250000,
            -1.890047164730000,
            -1.906381561480000,
            -1.878485000840000,
            -1.904840468160000,
            -1.909001318840000,
            -1.896419438340000,
            -1.866847792700000,
            -1.879134636380000,
            -1.894351689880000,
            -1.889768860260000,
            -1.943694824540000,
            -1.898724650100000};
        double[] expResult = new double[]{1.55574514119097};
        double[] result = Pub.robustReg2(xValues, yValues);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        // TODO: test 8 other results

        System.out.println("robustReg2 test 2");
        xValues = new double[]{
            1.76417316,
            1.714489829,
            1.734569224,
            1.771351941,
            1.768402061,
            1.766342834,
            1.766277735};
        yValues = new double[]{
            -1.529271388,
            -1.650928654,
            -1.573850599,
            -1.574514219,
            -1.577817993,
            -1.535529651,
            -1.579515567};
        expResult = new double[]{0.9856414973973155};
        result = Pub.robustReg2(xValues, yValues);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        // TODO: test 8 other results
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
     *
     * per Bodorkos from Squid2.5 file
     * 100142_ShowcaseTaskSwitches_2017-05-15_frozen.xls:
     */
    @Test
    public void testConcordia() {
        System.out.println("concordia");
        double r207Pb_235U = 0.256640683538696;
        double r207Pb_235U_1SigmaAbs = 0.00188371155299724;
        double r206Pb_238U = 0.0367475296433289;
        double r206Pb_238U_1SigmaAbs = 0.0000998234261966304;
        double rho = 0.370096880869828;
        double[] expResult = new double[]{232628743.083742, 0620177.514250217, 0.937383339484404};
        double[] result = Pub.concordia(r207Pb_235U, r207Pb_235U_1SigmaAbs, r206Pb_238U, r206Pb_238U_1SigmaAbs, rho);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 12), Utilities.roundedToSize(result[1], 12), SquidConstants.SQUID_EPSILON);
//        assertEquals(Utilities.roundedToSize(expResult[2], 12), Utilities.roundedToSize(result[2], 12), SquidConstants.SQUID_EPSILON);
        // todo: test additional outputs
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
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_original_frozen.xls:
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
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 12), Utilities.roundedToSize(result[1], 12), SquidConstants.SQUID_EPSILON);

        System.out.println("age7corrWithErr  # 2 (from perm4_7corr)");
        totPb6U8 = 0.515826107781068;
        totPb6U8err = 0.922853256399605 / 100.0 * 0.515826107781068;
        totPb76 = 0.182991055066;
        totPb76err = 0.414818785308 / 100.0 * 0.182991055066;
        expResult = new double[]{2682151784.44307, 31459190.5096457};
        result = Pub.age7corrWithErr(totPb6U8, totPb6U8err, totPb76, totPb76err, 0.8741, 9.8485E-10, 1.55125E-10, 137.88);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 12), Utilities.roundedToSize(result[1], 12), SquidConstants.SQUID_EPSILON);

    }

    /**
     * Test of agePb76WithErr method, of class Pub.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_orig_2017-07-17_frozen.xls:
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
        double pb76rad = 0.0552518706519236;
        double pb76err = 1.96293438301707 / 100.0 * 0.0552518706519236;
        double[] expResult = new double[]{422429481.64047, 43809508.1776918};
        double[] result = Pub.agePb76WithErr(pb76rad, pb76err);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 12), Utilities.roundedToSize(result[1], 12), SquidConstants.SQUID_EPSILON);
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
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 12), Utilities.roundedToSize(result[1], 12), SquidConstants.SQUID_EPSILON);

        System.out.println("age8corrWithErr #2");
        totPb6U8 = 0.516689203818454;
        totPb6U8err = 0.922853256399605 / 100.0 * 0.516689203818454;
        totPb8Th2 = 0.147397260057089;
        totPb8Th2err = 2.32602099147682 / 100.0 * 0.147397260057089;
        th2U8 = 0.581453384857;
        th2U8err = 0.513472145943 / 100.0 * 0.581453384857;
        expResult = new double[]{2.67844233729399E9, 2.2149610752408966E7};
        result = Pub.age8corrWithErr(totPb6U8, totPb6U8err, totPb8Th2, totPb8Th2err, th2U8, th2U8err, 2.1095, 4.9475E-11, 1.55125E-10);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 12), Utilities.roundedToSize(result[1], 12), SquidConstants.SQUID_EPSILON);

    }

    /**
     * Test of wtdXYmean method, of class Pub.
     */
    @Test
    public void testWtdXYmean() {
        System.out.println("wtdXYmean");
        double[] xValues = null;
        double[] xSigmaAbs = null;
        double[] yValues = null;
        double[] ySigmaAbs = null;
        double[] xyRho = null;
        double[] expResult = null;
//        double[] result = Pub.wtdXYmean(xValues, xSigmaAbs, yValues, ySigmaAbs, xyRho);
//        assertArrayEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
    }

    /**
     * Test of xyWtdAv method, of class Pub.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_orig_2017-07-17_frozen.xls:
     */
    @Test
    public void testXyWtdAv() {
        System.out.println("xyWtdAv");
        double[] xValues = new double[]{29.038276620114600,
            28.962731803448900,
            29.097060635479500,
            28.842384807160300,
            28.861563548018500,
            29.107374849558500,
            29.125213776627200,
            29.528318842930300,
            29.169244944767400,
            29.122758104293000,
            29.747385484389400,
            29.750080913075500,
            29.082376643611700,
            29.513185951318100,
            29.007162817001100,
            28.628556887449900};
        double[] xSigmaAbs = new double[]{0.339732464331131,
            0.316884522282400,
            0.269048291437233,
            0.285435980605201,
            0.280748328855637,
            0.269249215849140,
            0.345307114137523,
            0.277884441873062,
            0.294446996395807,
            0.270715610118071,
            0.343658642338273,
            0.324832736000108,
            0.305641622358978,
            0.295063499013700,
            0.293424006015355,
            0.285342147238329};
        double[] yValues = new double[]{0.703168242014160,
            0.702705092581175,
            0.705316527140717,
            0.699262093759952,
            0.703970115413977,
            0.704549396293564,
            0.704294472944355,
            0.712497620659772,
            0.705606392638837,
            0.705658519110398,
            0.719502681294069,
            0.723962155403346,
            0.706905406205872,
            0.712136584650369,
            0.701585244982455,
            0.697793590133956};
        double[] ySigmaAbs = new double[]{0.007911113090672,
            0.007439111803168,
            0.006201310416375,
            0.006488097247013,
            0.006442467323845,
            0.006197705221348,
            0.008117310853411,
            0.006355998210951,
            0.006660507180141,
            0.006231053858570,
            0.007989790917658,
            0.006744266777283,
            0.006892808937877,
            0.006729114150157,
            0.006632324976409,
            0.006504175799737};
        double[] xyRho = new double[]{0.961639144784375,
            0.967579103091148,
            0.950863750785816,
            0.937561605494920,
            0.940807295482065,
            0.950972724513276,
            0.972123615036394,
            0.947926877224024,
            0.935110326478509,
            0.949918006472440,
            0.961223830311981,
            0.853192909851003,
            0.927795658119983,
            0.945137955370786,
            0.934533688754913,
            0.935187787918873};
        double[] expResult = new double[]{29.160301234180100, 0.147378519040586 / 2.0, 0.706560044064936, 0.003383795098116 / 2.0, 0.947573268184246, 0.960019497573836, 0.5280897145990542};//todo why?  0.528090890062103};
        double[] result = Pub.xyWtdAv(xValues, xSigmaAbs, yValues, ySigmaAbs, xyRho);
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);

    }

}
