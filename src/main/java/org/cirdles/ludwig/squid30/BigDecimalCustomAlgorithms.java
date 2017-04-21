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
package org.cirdles.ludwig.squid30;

import java.math.BigDecimal;
import java.math.MathContext;

/**
 *
 * @author James F. Bowring
 */
public class BigDecimalCustomAlgorithms {

    /**
     *
     * @param value BigDecimal to find square root
     * @return BigDecimal square root of value
     */
    protected static BigDecimal bigDecimalSqrtBabylonian(BigDecimal value) {

        BigDecimal guess = new BigDecimal(StrictMath.sqrt(value.doubleValue()));

        if (guess.compareTo(BigDecimal.ZERO) > 0) {

            BigDecimal precision = BigDecimal.ONE.movePointLeft(34);
            BigDecimal theError = BigDecimal.ONE;
            while (theError.compareTo(precision) > 0) {
                BigDecimal nextGuess;
                try {
                    nextGuess = guess.add(value.divide(guess, MathContext.DECIMAL128)).divide(new BigDecimal(2.0), MathContext.DECIMAL128);
                    theError = guess.subtract(nextGuess, MathContext.DECIMAL128).abs();
                    guess = nextGuess;
                } catch (java.lang.ArithmeticException e) {
                    break;
                }
            }
        }
        return guess;
    }

}
