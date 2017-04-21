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

import static org.cirdles.squid.SquidConstants.SQUID_EPSILON;
import org.cirdles.utilities.Utilities;
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
     * Test of pbPbAge method, of class UPb.
     */
    @Test
    public void testPbPbAge() {
        System.out.println("pbPbAge");
        double pb76Rad = 0.055251870652859;
        double expResultAge = 422429481.678253;

        double pb76RadErr = 1.96293438298184 * pb76Rad / 100.0;// convert from % err
        double expResultAgeErr = 43809508.176617; // 1 sigma abs

        double[] result = UPb.pbPbAge(pb76Rad, pb76RadErr);

        assertEquals(Utilities.roundedToSize(expResultAge, 10), Utilities.roundedToSize(result[0], 10), SQUID_EPSILON);
        assertEquals(Utilities.roundedToSize(expResultAgeErr, 10), Utilities.roundedToSize(result[1], 10), SQUID_EPSILON);
    }

}
