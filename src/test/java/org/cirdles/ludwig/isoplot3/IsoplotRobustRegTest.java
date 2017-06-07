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
import org.junit.After;
import org.junit.AfterClass;
import static org.junit.Assert.assertArrayEquals;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
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
        double[] xValues = new double[]{0.,1.,2.};
        double[] yValues = new double[]{1.,2.,3.};
        double[] expResult = new double[]{1.,1.,-1.};
        // just test the slope and intercepts
        double[] result = RobustReg.getRobSlope(xValues, yValues)[0];
        assertArrayEquals(expResult, result, SquidConstants.SQUID_EPSILON);
    }
    
}
