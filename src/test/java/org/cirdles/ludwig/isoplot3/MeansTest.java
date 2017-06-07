/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.cirdles.ludwig.isoplot3;

import static org.cirdles.ludwig.squid25.SquidConstants.SQUID_EPSILON;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author bowring
 */
public class MeansTest {

    /**
     *
     */
    public MeansTest() {
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
     * Test of weightedAverage method, of class Means.
     */
    @Test
    public void testWeightedAverage() {
        System.out.println("weightedAverage using Ma data and uncertainties - oracle is Ludwig's Isoplot3.Means.WeighetedAverage");
        double[] values = new double[]{
            422.429481678253000,
            445.673004549890000,
            395.105564398245000,
            462.735195934239000,
            424.348437438747000,
            438.562012822161000,
            321.342581478880000,
            641.620699188942000,
            338.497170392008000,
            418.963816709918000,
            279.948361234572000,
            478.621649698142000,
            438.873773839037000,
            448.524445440179000,
            457.124787696760000,
            427.661504983727000,
            447.746216854234000,
            452.439908000677000,
            345.150478715612000,
            392.919246783614000,
            409.785559971735000,
            486.714775252251000,
            441.117853817133000

        };
        double[] errors = new double[]{
            43.809508176617000,
            30.272846598902800,
            57.018611692881700,
            38.073695378441000,
            39.359880328441900,
            45.398691532122000,
            50.895625766862800,
            63.112407173255800,
            70.996046586758100,
            56.368690915476400,
            95.711499355788400,
            47.052509854362600,
            41.811434926229700,
            69.081419146139200,
            33.840546325141300,
            44.523690999237800,
            36.484120802534300,
            47.762864733551300,
            54.239452697957400,
            30.475598210695400,
            54.371062520740100,
            56.614505311394500,
            83.429587216997200
        };

        double[] expResult = new double[]{431.72278878305707, 9.52995899409635, 1.3076310129910607, 0.15160797601388276, 5.249036183859748, 18.712872106652267};
        double[] result = Means.weightedAverage(values, errors);
        assertArrayEquals(expResult, result, SQUID_EPSILON);

    }

}
