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
public class PbUTh_2Test {

    public PbUTh_2Test() {
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

    /**
     * Test of pb46cor7 method, of class PbUTh_2.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_original_frozen.xls:
     *
     * Column header: 7-corr204Pb/206Pb
     *
     * Usage: "=Pb46cor7([207/206],sComm0_64,sComm0_74,[207corr206Pb/238UAge])"
     */
    @Test
    public void testPb46cor7() {
        System.out.println("pb46cor7");
        double pb76tot = 0.0592518351787661;
        double age7corPb6U8 = 564738561.835384;
        double[] expResult = new double[]{0.0000217282123636872};
        double[] result = PbUTh_2.pb46cor7(pb76tot, age7corPb6U8);
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of pb46cor8 method, of class PbUTh_2.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_original_frozen.xls:
     *
     * Column header: 8-corr204Pb/206Pb
     *
     * Usage:
     * "=Pb46cor8([208/206],[232Th/238U],sComm0_64,sComm0_84,[208corr206Pb/238UAge])"
     */
    @Test
    public void testPb46cor8() {
        System.out.println("pb46cor8");
        double pb86tot = 0.0831580681678389;
        double th2U8 = 0.27150941987617;
        double age8corPb6U8 = 565183115.935407;
        double[] expResult = new double[]{-0.0000237974958258082};
        double[] result = PbUTh_2.pb46cor8(pb86tot, th2U8, age8corPb6U8);
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of pb86radCor7per method, of class PbUTh_2.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_original_frozen.xls:
     *
     * Column header: %err (to the right of 7-corr208PbSTAR/206PbSTAR)
     *
     * Usage: "=Pb86radCor7per([208/206],[208/206 %err],[207/206],[207/206
     * %err],[Total206Pb/238U],[Total206Pb/238U
     * %err],[207corr206Pb/238UAge],sComm0_64,sComm0_74,sComm0_84)"
     *
     */
    @Test
    public void testPb86radCor7per() {
        System.out.println("pb86radCor7per");
        double pb86tot = 0.0831580681678389;
        double pb86totPer = 1.3868239637189;
        double pb76tot = 0.0592518351787661;
        double pb76totPer = 0.688011049802507;
        double pb6U8tot = 0.0915928758248389;
        double pb6U8totPer = 0.945106721698445;
        double age7corPb6U8 = 564738561.835384;
        double[] expResult = new double[]{2.51601400669704};
        double[] result = PbUTh_2.pb86radCor7per(pb86tot, pb86totPer, pb76tot, pb76totPer, pb6U8tot, pb6U8totPer, age7corPb6U8);
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of age7CorrPb8Th2 method, of class PbUTh_2.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_original_frozen.xls:
     *
     * Column header: 207corr208Pb/232ThAge
     *
     * Usage:
     * "=Age7CorrPb8Th2([Total206Pb/238U],[Total208Pb/232Th],[208/206],[207/206],sComm0_64,sComm0_76,sComm0_86)"
     *
     */
    @Test
    public void testAge7CorrPb8Th2() {
        System.out.println("age7CorrPb8Th2");
        double totPb206U238 = 0.0915928758248389;
        double totPb208Th232 = 0.0280531210114337;
        double totPb86 = 0.0831580681678389;
        double totPb76 = 0.0592518351787661;
        double[] expResult = new double[]{2.51601400669704};
        double[] result = PbUTh_2.age7CorrPb8Th2(totPb206U238, totPb208Th232, totPb86, totPb76);
//        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of age7CorrPb8Th2WithErr method, of class PbUTh_2.
     */
    @Test
    public void testAge7CorrPb8Th2WithErr() {
        System.out.println("age7CorrPb8Th2WithErr");
        double totPb206U238 = 0.0;
        double totPb206U238percentErr = 0.0;
        double totPb208Th232 = 0.0;
        double totPb208Th232percentErr = 0.0;
        double totPb86 = 0.0;
        double totPb86percentErr = 0.0;
        double totPb76 = 0.0;
        double totPb76percentErr = 0.0;
        double[] expResult = null;
        double[] result = PbUTh_2.age7CorrPb8Th2WithErr(totPb206U238, totPb206U238percentErr, totPb208Th232, totPb208Th232percentErr, totPb86, totPb86percentErr, totPb76, totPb76percentErr);
//        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }

}
