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
 * @author James F. Bowring, CIRDLES.org, and Earth-Time.org
 */
public class UtilitiesTest {

    public UtilitiesTest() {
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
     * Test of median method, of class Utilities.
     */
    @Test
    public void testMedian() {
        System.out.println("median");
        double[] values = new double[]{
            8.927153515620000,
            10.574811745900000,
            8.894700141240000,
            8.507551563150000,
            9.014881630990000,
            10.683879450300000,
            10.142231294300000,
            10.617826606000000,
            10.387153749700000,
            10.421390821500000,
            10.415758910000000,
            10.351636018900000,
            8.875345257050000,
            9.148679714420000,
            10.349959416100000,
            10.455234223700000,
            10.295134872800000,
            11.305933414200000,
            10.511166224700000,
            10.711717353900000,
            10.212016243500000,
            10.013806994600000,
            10.474133712300000,
            9.135694512270000,
            9.755354878260000};

        double expResult = 10.3499594161;
        double result = Utilities.median(values);
        assertEquals(expResult, result, 0.0);

    }

}
