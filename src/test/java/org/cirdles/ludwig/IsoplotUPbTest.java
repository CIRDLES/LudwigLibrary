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
package org.cirdles.ludwig;

import java.math.BigDecimal;
import java.math.RoundingMode;
import static org.cirdles.squid.SquidConstants.SQUID_EPSILON;
import org.junit.After;
import org.junit.AfterClass;
import static org.junit.Assert.assertEquals;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
 * @author James F. Bowring
 */
public class IsoplotUPbTest {
    
    public IsoplotUPbTest() {
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
     * Test of pbPbAge method, of class IsoplotUPb.
     */
    @Test
    public void testPbPbAge() {
        System.out.println("pbPbAge");
        double pb76Rad = 0.055251870652859;
        double expResult = 422429481.678253;
        double result = IsoplotUPb.pbPbAge(pb76Rad)[0][0];
        
        // force to 15 digits to match excel and vba
        BigDecimal resultBD = new BigDecimal(result);
        int newScale = 15 - (resultBD.precision() - resultBD.scale());
        result = Double.parseDouble(resultBD.setScale(newScale, RoundingMode.HALF_UP).toPlainString());
        assertEquals(expResult, result, SQUID_EPSILON);
    }
    
}
