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

import com.google.common.base.Joiner;
import com.google.common.base.Strings;

import java.util.Arrays;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * A translator to convert VBA to JavaScript. It is based on a recursive descent parser.
 * It does a syntactical conversion, while leaving application-specific constructs untouched.
 * Some constructs are translated to the greatest degree possible, and the responsibility for
 * further processing is left upto the compatibility layer or the user. Some examples:
 * <ol>
 *  <li /> Named parameter lists in VBA allow out-of-order parameters and have no counterpart
 *  in JS. They are simply broken up into 2 parameters, the name part and the value part.
 *  The responsibility for fixing this is left up to the user.
 *  <li /> Optional parameters
 * </ol>
 *
 * @author Brian Kernighan
 * @author Nikhil Singhal
 */
final class VbaJsConverter {

  /** Indent spacing (2 at the moment) */
  private static final String INDENT_SPACES = "  ";

  private final TranslationUnit unit;

  /** Stores the generated output. Using a StringBuilder for efficient concatenation of lines */
  private final StringBuilder generatedOutput;

  private VbaJsConverter() {
    this.unit = new TranslationUnit();
    this.generatedOutput = new StringBuilder();
  }

  static String convert(List<String> vbaCode) {
    return new VbaJsConverter().conversionHelper(vbaCode);
  }

  static String convert(String vbaCode) {
    return new VbaJsConverter().conversionHelper(vbaCode);
  }

  String conversionHelper(List<String> vbaCode)  {
    if (vbaCode == null || vbaCode.isEmpty()) {
      return "";
    }

    unit.cleanup(vbaCode);

    unit.advance();
    while (!unit.getCurrentLine().peek().equals(ConverterUtil.EOF)) {
      translate();
    }

    // Consistency check on indent level
    if (unit.getDepth() != 0) {
      throw new ParseException("Statement nesting error: depth = " + unit.getDepth(),
          unit.getCurrentLineNumber(), unit.getLine(unit.getCurrentLineNumber()));
    }

    return generatedOutput.toString();
  }

  /**
   * This function converts the input VB script into syntactically identical JavaScript (for the
   * most part). It is NOT thread-safe.
   * @param vbaCode The VB script that needs to be converted
   * @return The generated JavaScript
   */
  private String conversionHelper(String vbaCode)  {
    if (Strings.isNullOrEmpty(vbaCode)) {
      return "";
    }

    return convert(Arrays.asList(vbaCode.split(ConverterUtil.LINE_SEPARATOR)));
  }

  /**
   * Generates a single line of output, with comment (if any) at proper indentation level.
   */
  private void generateOneLine(String... linePieces) {
    String jsLine = Joiner.on("").join(linePieces);

    String comment = "";
    if (unit.getCurrentLine().hasComment()) {
      comment = " // " + unit.getCurrentLine().getComment().trim();
    }

    if (jsLine.isEmpty()) {
      comment = comment.trim();
    }

    String blanks = Strings.repeat(INDENT_SPACES, unit.getDepth());

    generatedOutput.append(blanks);
    generatedOutput.append(jsLine);
    generatedOutput.append(comment);
    generatedOutput.append(ConverterUtil.LINE_SEPARATOR);
  }

  /**
   * Collect argument list for subroutine and function definitions.
   * Deletes ByVal and ByRef, preserves type as comment
   */
  private String collectArgList()  {
    String argList = "";
    if (!unit.getCurrentLine().peek().equals("(")) {
      return argList;
    }

    unit.getCurrentLine().eat("(");

    while (!unit.getCurrentLine().peek().equals(")")) {
      String ref = "";
      if (unit.getCurrentLine().peek().equals("ByRef")) {
        ref = "/*ByRef*/";
        unit.getCurrentLine().eat("ByRef");
      } else if (unit.getCurrentLine().peek().equals("Optional")) {
        ref = "/*Optional*/";
        unit.getCurrentLine().eat("Optional");
      } else if (unit.getCurrentLine().peek().equals("ByVal")) {
        unit.getCurrentLine().eat("ByVal");
      }
      String currentLine = unit.getCurrentLine().peek();
      String name = unit.getCurrentLine().getName();
      if (!currentLine.equals(name)) {
        setArrayName(currentLine);
      }

      name = ref + currentLine;
      argList += name;

      if (unit.getCurrentLine().peek().equals("As")) {
        unit.getCurrentLine().getToken(true); // As
        unit.getCurrentLine().getName(); // type
      }

      if (unit.getCurrentLine().peek().equals("=")) {  // presumably only if Optional
        unit.getCurrentLine().eat("=");
        String expr = unit.getCurrentLine().getExpression();
        argList += " /*= " + expr + "*/";
      }

      if (unit.getCurrentLine().peek().equals(",")) {
        argList += unit.getCurrentLine().getToken(true) + " ";
      }
    }
    unit.getCurrentLine().eat(")");
    return argList;
  }

  /**
   * Generate and properly initialize a multi-dimensional JavaScript array. Multi-dimensional
   * arrays in JS are arrays of arrays, each new dimension needs to be initialized separately.
   * We use the fact the variable names in VB cannot start with an underscore, but they can in
   * JS. This is used to make sure that our loop variables have no scoping clashes with the
   * user's VB variables.
   *
   * @param var The name of the array
   * @param vtype The array type
   * @param indices The indices of the multi-dimensional array. eg: Dim x(1, 2, 3) => [1, 2, 3]
   */
  private void generateMultiDimArray(String var, String vtype, String indices[]) {
    // Array declaration for first dimension
    generateOneLine("var ", var, " = new Array(", indices[0], "); ", vtype);
    // TODO(nikhil): Handle the wrap-around case (> 26 dimensions)
    char idx = 'a';
    String subscript = "";
    // Generate the nested for-loops to initialize the remaining n-1 dimensions.
    for (int i = 1, len = indices.length; i < len; ++i) {
      // Use the fact that VB variables cannot start with an _ to prevent scoping clashes.
      String internalIdx = "_" + idx;
      generateOneLine("for (var ", internalIdx, " = 0; ", internalIdx, " < ", indices[i - 1],
          "; ++", internalIdx + ") {");
      unit.indent();
      subscript += "[" + internalIdx + "]";
      ++idx;
      generateOneLine(var, subscript, " = new Array(", indices[i], ");");
    }

    // Back out of the nested for loops
    for (int i = 1, len = indices.length; i < len; ++i) {
      unit.undent();
      generateOneLine("}");
    }
  }

  /**
   * This is for lines that the translator doesn't know how to handle. They are commented out
   * for now; this behavior might change in the future.
   */
  private String markLineAsUntouched(String... linePieces) {
    return "// " + Joiner.on("").join(linePieces) + "; // UNTOUCHED";
  }

  /**
   * Put parens around a string if it has any non-alphanumerics.
   */
  private String parenthesize(String str) {
    if (str.matches("^\\w+$") || str.matches("^\"[^\"]*\"$")) {
      return str;
    } else {
      return "(" + str + ")";
    }
  }

  private void setArrayName(String str) {
    if (unit.getSubNestingValue() > 0) {
      unit.addLocalName(str);
    } else {
      unit.addGlobalName(str);
    }
  }

  /**
   * Empty lines might include comments; either way, print them
   */
  private void skipEmptyLines()  {
    while (unit.getCurrentLine().peek().isEmpty()) {
      translateEmpty();
    }
  }

  /**
   * Starts with current line. Decide what kind of statement we have and call the right
   * translation function.
   */
  private void translate()  {
    String peek = unit.getCurrentLine().peek();
    String peektype = unit.getCurrentLine().peekTokenType();

    if (peek.equals(ConverterUtil.EOF)) {
      throw new ParseException("Unexpected end of file, line " +
          unit.getCurrentLine().getOriginal(), unit.getCurrentLineNumber(), null);
    }

    if (peek.isEmpty()) {
      translateEmpty();
    } else if (peek.equals("Dim") || peek.equals("ReDim")
        || peek.equals("Global") || peek.equals("Const")) {
      translateDim();
    } else if (peek.equals("If")) {
      translateIf();
    } else if (peek.equals("For")) {
      translateFor();
    } else if (peek.equals("Do")) {
      translateDo();
    } else if (peek.equals("While")) {
      translateWhile();
    } else if (peek.equals("Sub")) {
      translateSub();
    } else if (peek.equals("Function")) {
      translateFunction();
    } else if (peek.equals("Call")) {
      translateCall();
    } else if (peek.equals("Select")) {
      translateSelect();
    } else if (peek.equals("Exit")) {
      translateExit();
    } else if (peek.equals("With")) {
      translateWith();
    } else if (peek.equals("Type")) {
      translateType();
    } else if (peektype.equals("PUNT")) {
      translatePunt();
    } else if (peek.equals("On Error")) {
      translateOnError();
    } else if (peektype.equals("ID")) {
      translateAssignmentOrCall();
    } else if (peek.equals(".")) {
      translateAssignmentOrCall();
    } else {
      translateOther();
    }
  }

  /**
   * Translate foo, foo(bar) and foo bar. gets it wrong if the first argument starts with a
   * paren -- too ambiguous. This is balanced on a pinhead.
   */
  private void translateAssignmentOrCall()  {
    String name = unit.getCurrentLine().getName();
    // The assignment or call expression
    String expr;

    if (unit.getCurrentLine().peek().equals(":")) { // a label?
      String rest = unit.getCurrentLine().getRest().trim();
      generateOneLine(markLineAsUntouched(name, " ", rest));
      unit.advance();
      return;
    }

    // For cases like: foo (p1), (p2). These are transformed into foo ((p1), (p2)) and
    // put back into the list of lines.
    if (unit.getCurrentLine().peek().equals(",")) {
      String original = unit.getCurrentLine().getOriginal();
      int separatorIndex = original.indexOf(" ");
      if (separatorIndex >= 0) {
        unit.getCurrentLine().parseLine(original);  // start over with original line
        original = unit.getCurrentLineAsString();
        original = original.substring(0, separatorIndex) + "(" +
            original.substring(separatorIndex + 1, original.length()).trim() + ")";
        if (unit.getCurrentLine().hasComment()) { // restore comment if there was one
          original += "' " + unit.getCurrentLine().getComment();
        }
        unit.getCurrentLine().parseLine(original);   // parse the modified line
        translateAssignmentOrCall();
        return;
      }
    }

    if (unit.getCurrentLine().peek().equals("=")) { // assignment
      unit.getCurrentLine().eat("=");
      if (name.equals(unit.getFunctionName())) {
        name = "_" + name;
      }
      String newstr = "";

      if (unit.getCurrentLine().peek().equals("New")) {
        unit.getCurrentLine().eat("New");
        newstr = "new ";
      } else if (unit.getCurrentLine().peek().startsWith("Array")) {
        newstr = "new ";
        setArrayName(name);
      }

      expr = name + " = " + newstr + unit.getCurrentLine().getExpression();
    } else if (currentTokenTypeEquals("ID") ||
        currentTokenTypeEquals("NUM") ||
        currentTokenTypeEquals("STR") ||
        unit.getCurrentLine().peek().equals("-")) {
      // probably foo bar,glop
      StringBuilder callParamsList = new StringBuilder();
      while (!unit.getCurrentLine().peek().isEmpty() &&
          !currentTokenTypeEquals("KEY") &&
          !unit.getCurrentLine().peek().equals(":")) {
        callParamsList.append(unit.getCurrentLine().getExpression());
        if (unit.getCurrentLine().peek().equals(",")) {
          callParamsList.append(unit.getCurrentLine().getToken(true)).append(" ");
        }
      }
      expr = name + "(" + callParamsList.toString() + ")";
    } else { // who knows
      String rest = unit.getCurrentLine().getRest().trim();
      if (rest.isEmpty() && !name.matches(".*\\(.*\\)$")) {
        expr = name + "()"; // guess it"s a function call
      } else {
        expr = name + " " + rest;
      }
    }
    generateOneLine(expr.trim(), ";");

    // Handles multiple statements on one line separated :
    if (unit.getCurrentLine().peek().equals(":")) {
      unit.getCurrentLine().eat(":");
    } else {
      unit.advance();
    }
  }

  /**
   * Translate an explicit Call statement, either Call this, that, theother or
   * Call(this, that, theother).
   */
  private void translateCall()  {
    unit.getCurrentLine().eat("Call");
    String name = unit.getCurrentLine().getName();
    StringBuilder callParamsList = new StringBuilder();
    if (unit.getCurrentLine().peek().isEmpty()) { // Call foo(...) or Call foo
      if (name.matches(".*\\(.*\\)$")) {
        generateOneLine(name, ";");
      } else {
        generateOneLine(name, "();");
      }
    } else if (unit.getCurrentLine().peek().equals("(")) {
      while (!unit.getCurrentLine().peek().isEmpty()) {
        callParamsList.append(unit.getCurrentLine().getExpression());
        if (unit.getCurrentLine().peek().equals(",")) {
          callParamsList.append(unit.getCurrentLine().getToken(true)).append(" ");
        }
      }
      generateOneLine(name, callParamsList.toString(), ";");
      // should eat the closing paren
    } else {
      while (!unit.getCurrentLine().peek().isEmpty()) {
        callParamsList.append(unit.getCurrentLine().getExpression());
        if (unit.getCurrentLine().peek().equals(",")) {
          callParamsList.append(unit.getCurrentLine().getToken(true)).append(" ");
        }
      }
      generateOneLine(name, "(", callParamsList.toString(), ");");
    }
    unit.advance();
  }

  /**
   * Innards of a single Case
   */
  private void translateCase(String expr, int n)  {
    unit.getCurrentLine().eat("Case");
    String elsePart;
    if (n == 1) {
      elsePart = "";
    } else {
      elsePart = "} else ";
    }

    if (unit.getCurrentLine().peek().equals("Else")) {
      unit.getCurrentLine().eat("Else");
      generateOneLine("} else {");
    } else {
      // expression1 To expression2
      // [ Is ] comparisonoperator expression
      // expression
      String ifExpr = "";
      String toExpr;
      while (!unit.getCurrentLine().peek().isEmpty() && !unit.getCurrentLine().peek().equals(":")) {
        if (unit.getCurrentLine().peek().equals("Is")) {
          unit.getCurrentLine().eat("Is");
        }
        if (currentTokenTypeEquals("OP") &&
            !(unit.getCurrentLine().peek().equals("-") ||
                unit.getCurrentLine().peek().equals("+"))) {
          String relOp = ConverterUtil.fixOperators(unit.getCurrentLine().getToken(true));
          toExpr = unit.getCurrentLine().getExpression();
          ifExpr += expr + " " + relOp + " " + parenthesize(toExpr);
        } else {
          toExpr = unit.getCurrentLine().getExpression();
          if (unit.getCurrentLine().peek().equals("To")) {
            unit.getCurrentLine().eat("To");
            String e3 = unit.getCurrentLine().getExpression();
            ifExpr += expr + " >= " + toExpr + " && " + expr + " <= " + e3;
          } else {
            ifExpr += expr + " == " + parenthesize(toExpr);
          }
        }
        if (unit.getCurrentLine().peek().equals(",")) {
          unit.getCurrentLine().eat(",");
          ifExpr += " || ";
        }
      }
      generateOneLine(elsePart, "if (", ifExpr, ") {");
    }
    unit.indent();
    if (unit.getCurrentLine().peek().equals(":")) { // meant to handle 1-liners
      unit.getCurrentLine().eat(":");
      translate();
    } else {
      unit.advance();
      while (!unit.getCurrentLine().peek().equals("Case") && !unit.getCurrentLine().peek().equals(
          "End Select")) {
        translate();
      }
    }
    unit.undent();
  }

  /**
   * Dim x As type, y(10) As type, z As type = expr.
   * Generates new Array() for arrays and remembers names so can convert () to [] when used
   * in expression.
   */
  private void translateDim()  {
    String kind = unit.getCurrentLine().getToken(true); // Dim, ReDim, Global or Const
    String indices[] = null;
    boolean isUserDefinedType = false;

    while (true) {
      String var = unit.getCurrentLine().getToken(true);
      if (var.equals("Preserve")) {
        var = unit.getCurrentLine().getToken(true);
      }

      String dim = ""; // not an array
      if (unit.getCurrentLine().peek().equals("(")) {
        dim = unit.getCurrentLine().getBalancedParentheses();

        Pattern rangePattern = Pattern.compile("(.*)To(.*)");

        indices = dim.replaceAll("\\(", "").replaceAll("\\)", "").split(",");
        for (int i = 0, len = indices.length; i < len; ++i) {
          Matcher rangeMatcher = rangePattern.matcher(indices[i]);
          if (rangeMatcher.matches()) {
            // TODO(nikhil): We aren't storing the lower limit now. Might want
            // to do that later.
            indices[i] = rangeMatcher.group(2);
          }
        }

        if (dim.matches(".*To.*")) {
          dim = dim.replaceAll("To", " To ");
          dim = "(/* " + dim + " */)";
        }
      }

      String vtype = "";
      if (unit.getCurrentLine().peek().equals("As")) { // As [New] type
        unit.getCurrentLine().eat("As");
        if (unit.getCurrentLine().peek().equals("New")) {
          vtype = "New ";
          unit.getCurrentLine().eat("New");
        }
        vtype += unit.getCurrentLine().getName();

        // Dim foo as String * 100 (String with length 100)
        if (unit.getCurrentLine().peek().equals("*")) {
          vtype += unit.getCurrentLine().getToken(true);
          vtype += unit.getCurrentLine().getExpression();
        }
      }

      StringBuilder expr = new StringBuilder();
      if (unit.getCurrentLine().peek().equals("=")) { // some kind of initializer
        unit.getCurrentLine().eat("=");
        if (unit.getCurrentLine().peek().equals("{")) {
          unit.getCurrentLine().eat("{");
          while (!unit.getCurrentLine().peek().equals("}") && !unit.getCurrentLine().peek().equals(
              ConverterUtil.EOF)) {
            expr.append(unit.getCurrentLine().getToken(true));
          }
          unit.getCurrentLine().eat("}");
        } else {
          // scalar
          expr.append(unit.getCurrentLine().getExpression());
        }
      }

      if (!vtype.isEmpty()) {
        if (unit.isTypeName(vtype)) {
          isUserDefinedType = true;
        } else {
          vtype = "// " + vtype;
        }
      }

      if (dim.isEmpty()) { // it's not an array
        if (!(expr.length() == 0)) {
          expr.insert(0, " = ");
        }
        if (isUserDefinedType) {
          generateOneLine("var ", var, expr.toString(), " = new ", vtype, "();");
        } else {
          generateOneLine("var ", var, expr.toString(), "; ", vtype);
        }
      } else if (kind.equals("ReDim")) {
        if (!unit.isArrayName(var)) { // uses ReDim to declare array
          generateOneLine("var ", var, " = new Array", dim, "; ", vtype, " // ReDim decl");
          setArrayName(var);
        } else if (dim.indexOf(',') != -1) {  // flag multi-dim ReDim
          generateMultiDimArray(var, vtype, indices);
        }
      } else { // it is an array
        if (expr.length() == 0) {
          if (indices.length > 1) {
            vtype += " // multi-dim";
            generateMultiDimArray(var, vtype, indices);
          } else {
            generateOneLine("var ", var, " = new Array(", indices[0], ");");
          }
        } else {
          generateOneLine("var ", var, " = new Array(", expr.toString(), "); ", vtype);
        }
        setArrayName(var);
      }

      if (!unit.getCurrentLine().peek().equals(",")) {
        break;
      }
      unit.getCurrentLine().eat(",");
    }
    unit.advance();
  }

  /**
   * Translate Do [while/until e] ... Loop [while/until e]
   */
  private void translateDo()  {
    String doExpr;
    unit.getCurrentLine().eat("Do");
    if (unit.getCurrentLine().peek().equals("While")) {
      unit.getCurrentLine().eat("While");
      doExpr = unit.getCurrentLine().getExpression();
      generateOneLine("while (", doExpr, ") {");
    } else if (unit.getCurrentLine().peek().equals("Until")) {
      unit.getCurrentLine().eat("Until");
      doExpr = unit.getCurrentLine().getExpression();
      generateOneLine("while (!(", doExpr, ")) {");
    } else {
      generateOneLine("while (1) {");
    }

    unit.advance();
    unit.indent();

    while (!unit.getCurrentLine().peek().equals("Loop")) {
      translate();
    }

    unit.getCurrentLine().eat("Loop");
    if (unit.getCurrentLine().peek().equals("While")) {
      unit.getCurrentLine().eat("While");
      doExpr = unit.getCurrentLine().getExpression();
      generateOneLine("if (!(", doExpr, "))");
      unit.indent();
      generateOneLine("break;");
      unit.undent();
    } else if (unit.getCurrentLine().peek().equals("Until")) {
      unit.getCurrentLine().eat("Until");
      doExpr = unit.getCurrentLine().getExpression();
      generateOneLine("if (", doExpr, ")");
      unit.indent();
      generateOneLine("break;");
      unit.undent();
    }

    unit.undent();
    generateOneLine("}");
    unit.advance();
  }

  /**
   * Translate empty line (perhaps with comment)
   */
  private void translateEmpty() {
    generateOneLine("");
    unit.advance();
  }

  /**
   * Translate various kinds of Exits
   */
  private void translateExit()  {
    unit.getCurrentLine().eat("Exit");
    String token = unit.getCurrentLine().getToken(true);
    if (token.equals("For") || token.equals("While") || token.equals("Do")) {
      generateOneLine("break;");
    } else if (token.equals("Sub")) {
      generateOneLine("return;");
    } else if (token.equals("Function")) {
      generateOneLine("return _", unit.getFunctionName(), ";");
    } else {
      generateOneLine(unit.getCurrentLine().getRest(), "; // BUG");
    }

    unit.advance();
  }

  /**
   * For i = startExpr To stopExpr [Step stepExpr] =>
   * for (var i = start; i <= stop; i += step)
   */
  private void translateFor()  {
    String startExpr = "";
    String stopExpr = "";
    String rel;
    String incr;
    String stepExpr;

    unit.getCurrentLine().eat("For");

    if (unit.getCurrentLine().peek().equals("Each")) {
      translateForEach();
      return;
    }

    String var = unit.getCurrentLine().getToken(true);
    unit.getCurrentLine().eat("=");
    startExpr = unit.getCurrentLine().getExpression();
    String updown = unit.getCurrentLine().getToken(true);

    if (updown.equals("To")) {
      rel = "<=";
      incr = "+=";
    } else { // Downto
      rel = ">=";
      incr = "-=";
    }

    stopExpr = unit.getCurrentLine().getExpression();

    if (unit.getCurrentLine().peek().equals("Step")) {
      unit.getCurrentLine().eat("Step");
      stepExpr = unit.getCurrentLine().getExpression();
      if (stepExpr.substring(0, 1).equals("-")) {
        rel = ">=";
        incr = "+=";
      }
    } else {
      stepExpr = "1";
    }

    // Convert increments/decrements of 1 to ++/--
    String reincr;
    if (stepExpr.equals("1") && incr.equals("+=")) {
      reincr = "++" + var;
    } else if (stepExpr.equals("-1") && incr.equals("-=")) {
      reincr = "++" + var;
    } else if (stepExpr.equals("1") && incr.equals("-=")) {
      reincr = "--" + var;
    } else if (stepExpr.equals("-1") && incr.equals("+=")) {
      reincr = "--" + var;
    } else {
      reincr = var + " " + incr + " " + stepExpr;
    }

    // JS hoists all variables to function scope
    generateOneLine("for (var ", var, " = ", startExpr, "; ", var, " ", rel, " ",  stopExpr, "; ",
        reincr, ") {");
    unit.indent();
    unit.advance();

    while (!unit.getCurrentLine().peek().equals("Next") &&
        !unit.getCurrentLine().peek().equals(ConverterUtil.EOF)) {
      translate();
    }

    unit.undent();
    generateOneLine("}");
    unit.advance();
  }

  /**
   * For Each var In whatever ... Next
   */
  private void translateForEach()  {
    unit.getCurrentLine().eat("Each");
    String var = unit.getCurrentLine().getToken(true);
    if (unit.getCurrentLine().peek().equals("As")) { // skip optional As type
      unit.getCurrentLine().eat("As");
      unit.getCurrentLine().getName();
    }
    unit.getCurrentLine().eat("In");
    String expr = unit.getCurrentLine().getExpression();
    generateOneLine("for (var ", var, " in ", expr, ") {");
    unit.indent();
    unit.advance();
    while (!unit.getCurrentLine().peek().equals("Next") && !unit.getCurrentLine().peek().equals(
        ConverterUtil.EOF)) {
      translate();
    }
    unit.undent();
    generateOneLine("}");
    unit.advance();
  }

  /**
   * Function whatever(arglist) As whatever ... End Function.
   * This should do something with the function name so return expr works properly.
   */
  private void translateFunction()  {
    unit.enterSub();
    unit.getCurrentLine().eat("Function");
    unit.setFunctionName(unit.getCurrentLine().getToken(true));
    String argList = collectArgList();
    StringBuilder ret = new StringBuilder();
    String returnVariable = "_" + unit.getFunctionName();

    while (unit.getCurrentLine().hasToken()) {
      unit.getCurrentLine().getToken(true);
      if (!unit.getCurrentLine().getCurrentToken().equalsIgnoreCase("As")) {// skip 'As Double'
        ret.append(unit.getCurrentLine().getCurrentToken());
      } else {
        unit.getCurrentLine().eat("As");
        ret.append(unit.getCurrentLine().getCurrentToken());
      }
    }

    if (ret.length() != 0) {
      ret.insert(0, " // ");
    }

    generateOneLine("function ", unit.getFunctionName(), "(", argList, ") {", ret.toString());
    unit.indent();
    generateOneLine("var ", returnVariable, " = \"\"; // Stores return value");
    unit.advance();

    while (!unit.getCurrentLine().peek().equals("End Function")) {
      translate();
    }

    unit.getCurrentLine().eat("End Function");
    generateOneLine("return ", returnVariable, ";");
    unit.undent();
    unit.setFunctionName("");
    generateOneLine("}");
    unit.leaveSub();
    unit.advance();
  }

  /**
   * If ... Then \n stat \n [ElseIf ... \n stat ] [Else \n stat ] End If
   */
  private void translateIf()  {
    unit.getCurrentLine().eat("If");
    String expression = unit.getCurrentLine().getExpression();
    unit.getCurrentLine().eat("Then");
    generateOneLine("if (", expression, ") {");
    unit.indent();
    unit.advance();

    while (!unit.getCurrentLine().peek().equals("End If") &&
        !unit.getCurrentLine().peek().equals("Else") &&
        !unit.getCurrentLine().peek().equals("ElseIf")) {
      translate();
    }

    while (unit.getCurrentLine().peek().equals("ElseIf")) {
      unit.getCurrentLine().eat("ElseIf");
      unit.undent();
      expression = unit.getCurrentLine().getExpression();
      unit.getCurrentLine().eat("Then");
      generateOneLine("} else if (", expression, ") {");
      unit.indent();
      unit.advance();

      while (!unit.getCurrentLine().peek().equals("End If") &&
          !unit.getCurrentLine().peek().equals("Else") &&
          !unit.getCurrentLine().peek().equals("ElseIf")) {
        translate();
      }
    }

    if (unit.getCurrentLine().peek().equals("Else")) {
      unit.getCurrentLine().eat("Else");
      unit.undent();
      generateOneLine("} else {");
      unit.advance();
      unit.indent();
      while (!unit.getCurrentLine().peek().equals("End If")) {
        translate();
      }
    }

    unit.getCurrentLine().eat("End If");
    unit.undent();
    generateOneLine("}");
    unit.advance();
  }

  /**
   * On Error [Resume Next / Resume lab / GoTo lab. No idea what the scope of these things is.
   * "Scope" is probably the wrong idea, more like setting a state.
   */
  private void translateOnError()  {
    unit.getCurrentLine().eat("On Error");
    if (unit.getCurrentLine().peek().equals("Resume")) {
      unit.getCurrentLine().eat("Resume");
      generateOneLine("// On Error Resume ", unit.getCurrentLine().getRest(), "; // UNTOUCHED");
      unit.advance();

    } else if (unit.getCurrentLine().peek().equals("GoTo")) {
      unit.getCurrentLine().eat("GoTo");
      String token = unit.getCurrentLine().getToken(true);
      if (token.equals("0")) {
        generateOneLine("// On Error GoTo 0; // UNTOUCHED");
        unit.advance();
        return; // special case in VB: restore normal handling
      }

      generateOneLine("try {");
      unit.indent();
      unit.advance();

      while (!unit.getCurrentLine().peek().equals(token)) {
        translate();
      }

      unit.advance();
      unit.undent();
      generateOneLine("} catch(e) { // ", token);
      unit.indent();

      while (!unit.getCurrentLine().peek().equals("End Sub") &&
          !unit.getCurrentLine().peek().equals("End Function")) {
        translate();
      }

      unit.undent();
      generateOneLine("}");
    }
  }

  /**
   * Not sure so just put it out.
   */
  private void translateOther()  {
    generateOneLine(markLineAsUntouched(unit.getCurrentLine().getRest()));
    unit.advance();
  }

  /**
   * Something sufficiently bad that we know to ignore it.
   * e.g., Attribute|Option|Open|Close|Declare|Line
   */
  private void translatePunt() {
    generateOneLine(markLineAsUntouched(unit.getCurrentLine().getLine()));
    unit.advance();
  }

  /**
   * Select ... Case ... [Case Else] End Select.
   * This is a nightmare statement because Case exprs are a mess.
   */
  private void translateSelect()  {
    unit.getCurrentLine().eat("Select");
    unit.getCurrentLine().eat("Case");
    String e = unit.getCurrentLine().getExpression();
    skipEmptyLines(); // vbFile.advance();
    int n = 1;

    while (!unit.getCurrentLine().peek().equals("End Select")) {
      if (unit.getCurrentLine().peek().equals("Case")) {
        translateCase(e, n);
        n += 1;
      }
    }

    unit.getCurrentLine().eat("End Select");
    generateOneLine("}");
    unit.advance();
  }

  /**
   * Sub name(arglist) ... End Sub. Should be skipping Private, etc.
   */
  private void translateSub()  {
    unit.enterSub();
    unit.getCurrentLine().eat("Sub");
    String subname = unit.getCurrentLine().getToken(true);
    String argList = collectArgList();
    generateOneLine("function ", subname, "(", argList, ") {");
    unit.indent();
    unit.advance();

    while (!unit.getCurrentLine().peek().equals("End Sub") &&
        !unit.getCurrentLine().peek().equals(ConverterUtil.EOF)) {
      translate();
    }

    unit.getCurrentLine().eat("End Sub");
    unit.undent();
    generateOneLine("}");
    unit.leaveSub();
    unit.advance();
  }

  /**
   * Translates user-defined VB types. <br />
   * eg: <br />
   * <code>
   * Type foo
   *   x as Integer
   *   y
   * End Type
   * </code>
   */
  private void translateType()  {
    boolean isUserDefinedType = false;
    unit.getCurrentLine().eat("Type");
    // Type <name>
    String typeName = unit.getCurrentLine().getToken(true);

    // Add the name to the set of Type names. We use this later in case the user declares variables
    // of that type.
    unit.addTypeName(typeName);

    unit.advance();

    // JS class constructor
    generateOneLine(typeName, " = function() {};  // Creates an empty class");

    // We are in the middle of a Type declaration
    while (!unit.getCurrentLine().peek().equals("End Type")) {
      // Parse the variable declaration
      String name = unit.getCurrentLine().getToken(true);
      String vtype = "";
      if (unit.getCurrentLine().peek().equals("As")) {
        unit.getCurrentLine().eat("As");
        vtype = unit.getCurrentLine().peek();
      }

      // Is the variable type a user-defined Type?
      if (unit.isTypeName(vtype)) {
        isUserDefinedType = true;
      } else {
        vtype = "// " + vtype;
      }

      if (name.isEmpty()) {
        // Only a comment
        generateOneLine(unit.getCurrentLine().getRest());
      } else {
        // Attach the variable prototype
        if (isUserDefinedType) {
          generateOneLine(typeName, ".prototype.", name, " = new ", vtype, "();");
        } else {
          generateOneLine(typeName, ".prototype.", name, "; ", vtype);
        }
      }
      unit.advance();
    }

    unit.getCurrentLine().eat("End Type");
    unit.advance();
  }

  /**
   * Translate While e ... End While.
   */
  private void translateWhile()  {
    unit.getCurrentLine().eat("While");
    String expr = unit.getCurrentLine().getExpression();
    unit.advance();
    generateOneLine("while (", expr, ") {");
    unit.indent();

    while (!unit.getCurrentLine().peek().equals("End While") &&
        !unit.getCurrentLine().peek().equals("Wend")) {
      translate();
    }

    unit.getCurrentLine().getToken(true); // End While or Wend
    unit.undent();
    generateOneLine("}");
    unit.advance();
  }

  /**
   * With name ... End With.
   */
  private void translateWith()  {
    unit.getCurrentLine().eat("With");
    unit.addWithName(unit.getCurrentLine().getName());
    generateOneLine("// With ", unit.getWithName());
    unit.advance();

    while (!unit.getCurrentLine().peek().equals("End With")) {
      translate();
    }

    unit.getCurrentLine().eat("End With");
    try {
      unit.popWithName();
    } catch (NoSuchElementException e) {
      throw new ParseException(
          "Failed while translating With... End With. Out of statements to parse.");
    }
    unit.advance();
  }

  private boolean currentTokenTypeEquals(String other) {
    return unit.getCurrentLine().peekTokenType().equals(other);
  }

  // Main function for converting one macro at a time. Useful for
  // testing/debugging.
  // Takes a string as input.
  static void main(String[] args) {
    if (args.length == 1) {
      System.out.println(VbaJsConverter.convert(args[0]));
    }
  }
}