/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.cirdles.ludwig.squid25;

import static org.cirdles.ludwig.squid25.SquidConstants.PRESENT_238U235U;
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
     * per Bodorkos from Squid2.5 file 100142_G6147_orig_2017-07-17_frozen.xls:
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
        double totPb206U238 = 0.0915570339460798;
        double totPb208Th232 = 0.0280421792577493;
        double totPb86 = 0.0831580681678;
        double totPb76 = 0.0592518351787661;
        double[] expResult = new double[]{553438043.974625};
        double[] result = PbUTh_2.xage7CorrPb8Th2(totPb206U238, totPb208Th232, totPb86, totPb76);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of age7CorrPb8Th2WithErr method, of class PbUTh_2.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_orig_2017-07-17_frozen.xls:
     *
     * this method combines Squid2.5's Age7CorrPb8Th2 and AgeErr7CorrPb8Th2
     *
     * Column header: 207corr208Pb/232ThAge
     *
     * Column header: 1σ err (to the right of 207corr208Pb/232ThAge)
     *
     * Usage:
     * "=Age7CorrPb8Th2([Total206Pb/238U],[Total208Pb/232Th],[208/206],[207/206],sComm0_64,sComm0_76,sComm0_86)"
     *
     * Usage: "=AgeErr7corrPb8Th2([Total206Pb/238U],[Total206Pb/238U
     * %err],[Total208Pb/232Th],[Total208Pb/232Th %err],[207/206],[207/206
     * %err],[208/206],… [208/206 %err],sComm0_64,sComm0_76,sComm0_86)"
     */
    @Test
    public void testAge7CorrPb8Th2WithErr() {
        System.out.println("age7CorrPb8Th2WithErr");
        double totPb206U238 = 0.0915570339460798;
        double totPb206U238percentErr = 0.839844375205747;// / 100.0 * 0.0915570339460798;
        double totPb208Th232 = 0.0280421792577493;
        double totPb208Th232percentErr = 1.63774616675883;// / 100.0 * 0.0280421792577493;
        double totPb86 = 0.0831580681678;
        double totPb86percentErr = 1.38682396372;// / 100. * 0.0831580681678;
        double totPb76 = 0.0592518351787661;
        double totPb76percentErr = 0.688011049803;// / 100.0 * 0.0592518351787661;
        double[] expResult = new double[]{553438043.974625, 11665992.0078184};
        double[] result = PbUTh_2.xage7CorrPb8Th2WithErr(totPb206U238, totPb206U238percentErr, totPb208Th232, totPb208Th232percentErr, totPb86, totPb86percentErr, totPb76, totPb76percentErr);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 12), Utilities.roundedToSize(result[1], 12), SquidConstants.SQUID_EPSILON);

        System.out.println("age7CorrPb8Th2WithErr #2");
        totPb206U238 = 0.515826107781068;
        totPb206U238percentErr = 0.922853256399605;
        totPb208Th232 = 0.147151042427349;
        totPb208Th232percentErr = 2.32602099147682;
        totPb86 = 0.165872704801;
        totPb86percentErr = 2.07245310569;
        totPb76 = 0.182991055066;
        totPb76percentErr = 0.414818785308;
        expResult = new double[]{2785249252.63747, 187820763.531841};
        result = PbUTh_2.age7CorrPb8Th2WithErr(totPb206U238, totPb206U238percentErr, totPb208Th232, totPb208Th232percentErr, 
                totPb86, totPb86percentErr, totPb76, totPb76percentErr,
                17.821,0.8741,2.1095, 4.9475E-11, 9.8485E-10, 1.55125E-10, 137.88);
        assertEquals(Utilities.roundedToSize(expResult[0], 11), Utilities.roundedToSize(result[0], 11), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 10), Utilities.roundedToSize(result[1], 10), SquidConstants.SQUID_EPSILON);

    }

    /**
     * Test of pb206U238rad method, of class PbUTh_2.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_orig_2017-07-17_frozen.xls:
     *
     * Column header: 8corr206STAR/238
     *
     * Usage: "=Pb206U238rad([208corr206Pb/238UAge])"
     */
    @Test
    public void testPb206U238rad() {
        System.out.println("pb206U238rad");
        double age = 564971592.74734;
        double[] expResult = new double[]{0.0915964069749142};
        double[] result = PbUTh_2.xpb206U238rad(age);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of rad8corPb7U5WithErr method, of class PbUTh_2.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_orig_2017-07-17_frozen.xls:
     *
     * this method combines Squid2.5's Rad8corPb7U5 and Rad8corPb7U5PErr
     *
     * Column header: 8corr207STAR/235
     *
     * Column header: %err (to the right of 8corr207STAR/235)
     *
     * Usage:
     * "=Rad8corPb7U5([208corr206Pb/238UAge],[Total206Pb/238U],[207/206],sComm0_76)"
     *
     * Usage: "=Rad8corPb7U5perr([Total206Pb/238U],[Total206Pb/238U
     * %err],[8corr206STAR/238],[Total206Pb/238U]TIMES[207/206]/Present238U235U,[232Th/238U],[232Th/238U
     * %err],[207/206], [207/206 %err],[208/206],[208/206
     * %err],sComm0_76,sComm0_86)"
     */
    @Test
    public void testRad8corPb7U5WithErr() {
        System.out.println("rad8corPb7U5WithErr");
        double totPb6U8 = 0.0915570339460798;
        double totPb6U8per = 0.839844375205747;
        double radPb6U8 = 0.0915964069749142;
        double totPb7U5 = 0.0915570339460798 * 0.0592518351788 / PRESENT_238U235U;
        double th2U8 = 0.271509072107;
        double th2U8per = 0.231502107594;
        double totPb76 = 0.0592518351788;
        double totPb76per = 0.688011049803;
        double totPb86 = 0.0831580681678;
        double totPb86per = 1.38682396372;
        double[] expResult = new double[]{0.752677098785157, 1.0823631953636};
        double[] result = PbUTh_2.xrad8corPb7U5WithErr(totPb6U8, totPb6U8per, radPb6U8,
                totPb7U5, th2U8, th2U8per, totPb76, totPb76per, totPb86, totPb86per);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResult[1], 12), Utilities.roundedToSize(result[1], 12), SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of rad8corConcRho method, of class PbUTh_2.
     *
     * per Bodorkos from Squid2.5 file 100142_G6147_orig_2017-07-17_frozen.xls:
     *
     * Column header: err.corr. (to the right of %err, to the right of
     * 8corr206STAR/238)
     *
     * Usage: "=Rad8corConcRho([Total206Pb/238U],[Total206Pb/238U
     * %err],[8corr206STAR/238],[232Th/238U],[232Th/238U
     * %err],[207/206],[207/206 %err],[208/206],[208/206
     * %err],sComm0_76,sComm0_86)"
     */
    @Test
    public void testRad8corConcRho() {
        System.out.println("rad8corConcRho");
        double totPb6U8 = 0.0915570339460798;
        double totPb6U8per = 0.839844375205747;
        double radPb6U8 = 0.0915964069749142;
        double th2U8 = 0.271509072107;
        double th2U8per = 0.231502107594;
        double totPb76 = 0.0592518351788;
        double totPb76per = 0.688011049803;
        double totPb86 = 0.0831580681678;
        double totPb86per = 1.38682396372;
        double[] expResult = new double[]{0.646167502023579};
        double[] result = PbUTh_2.xrad8corConcRho(totPb6U8, totPb6U8per, radPb6U8, th2U8, th2U8per, totPb76, totPb76per, totPb86, totPb86per);
        assertEquals(Utilities.roundedToSize(expResult[0], 12), Utilities.roundedToSize(result[0], 12), SquidConstants.SQUID_EPSILON);
    }

    /**
     * Test of stdPb86radCor7per method, of class PbUTh_2.
     */
    @Test
    public void testStdPb86radCor7per() {
        System.out.println("stdPb86radCor7per");
        double pb86tot = 0.0701269310998;
        double pb86totPer = 6.04834961324;
        double pb76tot = 0.0596696323508;
        double pb76totPer = 0.600518775911;
        double radPb86cor7 = 0.0679088858715221;
        double pb46cor7 = 0.0000609634534116361;
        double stdRadPb76 = 0.0587838486664528;
        double alpha0 = 17.821;
        double beta0 = 15.5773361;
        double gamma0 = 37.5933995;
        double expResult = 6.3910701101908876;
        double result = PbUTh_2.stdPb86radCor7per(pb86tot, pb86totPer, pb76tot, pb76totPer, radPb86cor7, pb46cor7, stdRadPb76, alpha0, beta0, gamma0)[0];
        assertEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }

}
