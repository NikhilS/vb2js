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

import com.google.common.collect.Lists;
import com.google.common.collect.Sets;

import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * This class contains basic data about a VB file.
 *
 * @author Brian Kernighan
 * @author Nikhil Singhal
 */
final class TranslationUnit {

  private static final Pattern CONTINUATION_PATTERN = Pattern.compile("(.*)_$");

  /** Current line, in a Line object */
  private final Line currentLine;

  private final GlobalState globalState;

  /** All input lines as is, no \n. Fixed up in convert() */
  private final List<String> lines;

  /** Current line number. Advance happens first, so start at -1 */
  private int currentLineNumber;

  /** Depth of nested constructs */
  private int depth;

  /** Name of function currently being translated */
  private String functionName;

  private int subNestingValue;

  /**
   * User-defined type names. These are used when variables to these types are defined.
   * For normal variables, the type is just erased (Dim x as String => var x; // String), but for
   * types/classes this is not enough, since they will have some variables already bound to them.
   * Those are therefore translated as: Dim x as MyType => var x = new MyType();.
   */
  private final Set<String> typeNames;

  TranslationUnit() {
    this.globalState = new GlobalState();
    this.currentLine = new Line(globalState);
    this.currentLineNumber = -1;
    this.lines = Lists.newArrayList();

    this.depth = 0;
    this.functionName = "";
    this.subNestingValue = 0;

    this.typeNames = Sets.newHashSet();
  }

  void cleanup(Iterable<String> vba) {
    for (String line : vba) {
      if (line != null) {
        lines.add(line.trim());
      }
    }

    // Merge continuation lines (ending with _) into one long one
    for (int i = lines.size() - 1; i >= 0; --i) {
      String line = lines.get(i);
      Matcher continuationMatcher = CONTINUATION_PATTERN.matcher(line);
      if (continuationMatcher.matches()) {
        line = continuationMatcher.group(1) + lines.get(i + 1);
        lines.set(i, line);
        lines.remove(i + 1);
      }
    }

    // Convert 1-line If's into multi-line
    for (int i = lines.size() - 1; i >= 0; --i) {
      if (ConverterUtil.isOneLineIf(lines.get(i))) {
        rewriteOneLineIf(i);
      }
    }

    lines.add(ConverterUtil.EOF);
  }

  void addGlobalName(String name) {
    globalState.addGlobalName(name);
  }

  void addLocalName(String name) {
    globalState.addLocalName(name);
  }

  void addWithName(String name) {
    globalState.addWithName(name);
  }

  boolean isTypeName(String name) {
    return typeNames.contains(name);
  }

  void addTypeName(String name) {
    typeNames.add(name);
  }

  /**
   * Advance to the next line in lines[]
   */
  void advance() {
    ++currentLineNumber;
    if (currentLineNumber < lines.size()) {
      currentLine.parseLine(lines.get(currentLineNumber));
    }
  }

  /**
   * Entered a Sub/Function.
   */
  void enterSub() {
    ++subNestingValue;
  }

  /**
   * Left a Sub/Function
   */
  void leaveSub() {
    --subNestingValue;
    if (subNestingValue == 0) {
      globalState.clearLocalNames();
    }
  }

  Line getCurrentLine() {
    return currentLine;
  }

  int getDepth() {
    return depth;
  }

  String getFunctionName() {
    return functionName;
  }

  String getLine(int lineNumber) {
    return lines.get(lineNumber);
  }

  String getCurrentLineAsString() {
    return currentLine.getLine();
  }

  int getCurrentLineNumber() {
    return currentLineNumber;
  }

  int getSubNestingValue() {
    return subNestingValue;
  }

  String getWithName() {
    return globalState.getWithName();
  }

  void indent() {
    ++depth;
  }

  void undent() {
    --depth;
  }

  boolean isArrayName(String name) {
    return globalState.isArrayName(name);
  }

  void popWithName() {
    globalState.popWithName();
  }

  void setFunctionName(String functionName) {
    this.functionName = functionName;
  }

  /**
   * Convert If ... Then ... [Else ...] on one line into multiple lines so translateIf() can
   * handle it. Note: this needs to be case-independent if one is going down that path. Probably
   * should be done with re.I as an argument so it can be adjusted at run time rather than being
   * wired in. (but there's no "do nothing" 3rd arg to re.sub). Not coordinated with
   * Line.toUpperCase(), which tests whether to do conversion.
  */
  private void rewriteOneLineIf(int lineNumber) {
    String original = lines.get(lineNumber);
    String thenPart;
    String elsePart;
    int where;

    lines.set(lineNumber, original.replaceFirst("(?i)Then .*", "Then")); // if part
    thenPart = original.replaceFirst("(?i).*Then ", "");
    thenPart = thenPart.replaceFirst("(?i)Else .*", "").trim();
    where = lineNumber + 1;
    lines.add(where, thenPart);
    if (original.matches("(?i).*Else .+")) {
      elsePart = original.replaceFirst("(?i).*Else ", "").trim();
      where += 1;
      lines.add(where, "Else");
      where += 1;
      lines.add(where, elsePart);
    }
    lines.add(where + 1, "End If");
  }
}