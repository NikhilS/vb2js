/*
 * Copyright 2010 Google Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.google.vb2js;

import com.google.common.collect.ImmutableMap;

/**
 * Helper/utility class for the converter. Contains some constants, data structures and
 * static functions.
 *
 * @author Brian Kernighan
 * @author Nikhil Singhal
 */
final class ConverterUtil {

  // Note: The order is relevant here.
  /**
   * Map to convert between VB and JS operators.
   */
  private static final ImmutableMap<String, String> POSSIBLE_FIXES =
        new ImmutableMap.Builder<String, String>()
      .put("=", " == ")
      .put("<>", " != ")
      .put("<=", " <= ")
      .put(">=", " >= ")
      .put("<", " < ")
      .put(">", " > ")
      .put("&", " + ")
      .put("\\+", " + ")
      .put("-", " - ")
      .put("\\*", " * ")
      .put("/", " / ")
      .put("\\\\", " / ")
      .put("\\^", " BUG exp() ")
      .put("\\bXor\\b", " ^ ")
      .put("\\bAnd\\b", " && ")
      .put("\\bOr\\b", " || ")
      .put("\\bIs\\b", " == ")
      .put("\\bIsNot\\b", " != ")
      .put("\\bMod\\b", " % ")
      .put("\\bNew\\b", "new ")
      .put("\\bNot\\b", "!")
      .build();

  private static final String ONE_LINE_IF_THEN_ELSE = "(?i).*Then .+ Else .*";
  private static final String ONE_LINE_IF_THEN = "(?i).*Then .+";

  /** Marker for after last line */
  static final String EOF = "(EOF)";

  /** Cross-platform line separator */
  static final String LINE_SEPARATOR = System.getProperty("line.separator");

  /** Value generated for non-existent arguments */
  static final String EMPTY = "undefined";

  private ConverterUtil() {
  }

  /**
   * Replace VB operators by JS. don't apply to 'strings.'
   * ^ * / \ mod + - & = <> <= >= := >
   * < ! is not and or xor eqv imp like
   *
   * Precedence is mostly implemented
   */
  static String fixOperators(String token) {
    for (String fix : POSSIBLE_FIXES.keySet()) {
      if (token.matches(fix)) {
        return POSSIBLE_FIXES.get(fix);
      }
    }
    return token;
  }

  /**
   * Test whether line is one-line: If ... Then ... [Else ...].
   */
  static boolean isOneLineIf(String line) {
    return (line.matches(ONE_LINE_IF_THEN_ELSE) ||
        new Line().parseLine(line).getLine().matches(ONE_LINE_IF_THEN));
  }
}