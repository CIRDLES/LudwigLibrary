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

import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;

/**
 *
 * @author James F. Bowring
 */
public class Median {

    /**
     * Calculates arithmetic median of array of doubles.
     *
     * @param values
     * @return
     * @pre values has one element
     */
    public static double median(double[] values) {
        double median;
        // enforce precondition
        if (values.length == 0) {
            median = 0.0;
        } else {
            DescriptiveStatistics stats = new DescriptiveStatistics();
            // Add the data from the array
            for (int i = 0; i < values.length; i++) {
                stats.addValue(values[i]);
            }
            median = stats.getPercentile(50);
        }
        return median;
    }

}
