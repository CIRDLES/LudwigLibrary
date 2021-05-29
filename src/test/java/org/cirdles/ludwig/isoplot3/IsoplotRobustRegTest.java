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
import org.junit.*;

import static org.junit.Assert.assertArrayEquals;

/**
 * @author James F. Bowring
 */
public class IsoplotRobustRegTest {

    /**
     *
     */
    public IsoplotRobustRegTest() {
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
     * Test of getRobSlope method, of class RobustReg.
     */
    @Test
    public void testGetRobSlope() {
        System.out.println("getRobSlope");
        // line of slope 1 ;y-intercept 1; x-intercept -1
        double[] xValues = new double[]{0., 1., 2.};
        double[] yValues = new double[]{1., 2., 3.};
        double[] expResult = new double[]{1., 1., -1.};
        // just test the slope and intercepts
        double[] result = RobustReg.getRobSlope(xValues, yValues)[0];
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);

        System.out.println("getRobSlope");
        // robreg

        xValues = new double[]{
                1.82684879994743,
                1.83905643445086,
                1.820273240146,
                1.8223358268425,
                1.8232888110035,
                1.81743251694429,
                1.85319576798179,
                1.82935167703584,
                1.83868674461262,
                1.86572618266616,
                1.85870452296212,
                1.8546790828951,
                1.82476896750346,
                1.81840536440913,
                1.83265923536829,
                1.8276018554676,
                1.83789183216491,
                1.82638811829226,
                1.82931562128174,
                1.81580244447232

        };

        yValues = new double[]{
                -1.75658646460949,
                -1.75074717160227,
                -1.79079998499096,
                -1.78013307476607,
                -1.78666472357592,
                -1.7760918059022,
                -1.72757172718571,
                -1.74311348589945,
                -1.74330318664838,
                -1.71033331509239,
                -1.71373755156126,
                -1.72046861678066,
                -1.74955910362793,
                -1.80236708191103,
                -1.75507728407841,
                -1.78180935314371,
                -1.74063008616374,
                -1.76758846208862,
                -1.75891552633745,
                -1.79520670836701
        };

        expResult = new double[]{1.7845064314976458, -5.039733687443322, 2.658424309002852};
        // just test the slope and intercepts
        result = RobustReg.getRobSlope(xValues, yValues)[0];
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }

}