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
import com.google.common.collect.ImmutableSet;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * This localizes most of the processing for tokenizing a single line of input. Constructor
 * does strings and comments.
 * Lots of ad hocery here: ! is a component separator in VB, just blindly included in IDs here.
 * $ is valid at end of VB name.
 *
 * @author Brian Kernighan
 * @author Nikhil Singhal
 */
final class Line {

  // precedence order, high to low, from VB manual:
  // Arithmetic and Concatenation Operators
  // Exponentiation (^)
  // Unary identity and negation (+, -)
  // Multiplication and floating-point division (*, /)
  // Integer division (\)
  // Modulus arithmetic (Mod)
  // Addition and subtraction (+, -), string concatenation (+)
  // String concatenation (&)
  // Arithmetic bit shift (<<, >>)
  // Comparison Operators
  // =, <>, <, <=, >, >=, Is, IsNot, Like, TypeOf...Is
  // Logical and Bitwise Operators
  // Negation (Not)
  // Conjunction (And, AndAlso)
  // Inclusive disjunction (Or, OrElse)
  // Exclusive disjunction (Xor)
  // all operators are evaluated left to right
  // other complications:
  // And, Or, Xor are both bitwise if operands are integer
  // and logical if they are booleans (e.g., relational tests).
  // AndAlso, OrElse are short-circuit (really && and ||)

  // We're using an ImmutableMap here since the order in which the elements are
  // accessed matters.
  private static final ImmutableMap<Pattern, String> TYPES =
        new ImmutableMap.Builder<Pattern, String>()
      .put(Pattern.compile("(?i)\\b(Mod|Is|Not|AndAlso|And|OrElse|Or|Xor|Eqv|Like|New)\\b"),
          "OP")
      .put(Pattern.compile("(?i)\\b(End +(If|Sub|Function|While|With|Select))\\b"), "ENDXX")
      .put(Pattern.compile("(?i)\\b(Exit)\\b"), "EXIT")
      .put(Pattern.compile("(?i)\\b(Private|Public|Static|Let|Set)\\b"), "TOSS")
      .put(Pattern.compile("(?i)\\b(Attribute|Option|Declare)\\b"), "PUNT")
      .put(Pattern.compile("(?i)\\b(Open .* For |Close #\\w+)\\b"), "PUNT")
      .put(Pattern.compile("(?i)\\b(Print #|Line Input #)\\b"), "PUNT")
      .put(Pattern.compile("(?i)\\b(On Error (Resume Next|GoTo 0)|Resume|GoTo)\\b"), "PUNT")
      .put(Pattern.compile("(?i)\\b(On Error)\\b"), "ONERROR")
      .put(Pattern.compile("(?i)\\b(Then|Else|To|Downto|Step|As|ByVal|ByRef)\\b"), "KEY")
      .put(Pattern.compile("(?i)\\b(Type|End Type)\\b"), "TYPE")
      .put(Pattern.compile("[a-zA-Z](\\w)*\\$?"), "ID")
      .put(Pattern.compile("#\\d+/\\d+/\\d+#"), "DATE")
      .put(Pattern.compile("((\\d+\\.?\\d*)|(\\.\\d+))([eE][-+]?\\d+)?[&#]?"), "NUM")
      .put(Pattern.compile("&H[a-fA-F0-9]+"), "HEX")
      .put(Pattern.compile("<>|<=|>=|:="), "OP")
      .put(Pattern.compile("[*^/\\\\+\\-&=><]"), "OP")
      .put(Pattern.compile("\"[^\"]*\""), "STR")
      .put(Pattern.compile("\".*"), "COMMENT")
      .put(Pattern.compile("."), "CHR")
      .put(Pattern.compile("^$"), "END")
      .build();

  private static final ImmutableMap<String, String> KEYWORDS =
        new ImmutableMap.Builder<String, String>()
      .put("and", "And")
      .put("as", "As")
      .put("byref", "ByRef")
      .put("byval", "ByVal")
      .put("case", "Case")
      .put("const", "Const")
      .put("dim", "Dim")
      .put("do", "Do")
      .put("double", "Double")
      .put("downto", "Downto")
      .put("each", "Each")
      .put("else", "Else")
      .put("elseif", "ElseIf")
      .put("end", "End")
      .put("end function", "End Function")
      .put("end if", "End If")
      .put("end sub", "End Sub")
      .put("end select", "End Select")
      .put("end while", "End While")
      .put("end with", "End With")
      .put("error", "Error")
      .put("exit", "Exit")
      .put("false", "False")
      .put("for", "For")
      .put("function", "Function")
      .put("global", "Global")
      .put("goto", "GoTo")
      .put("if", "If")
      .put("integer", "Integer")
      .put("is", "Is")
      .put("like", "Like")
      .put("loop", "Loop")
      .put("mod", "Mod")
      .put("new", "New")
      .put("next", "Next")
      .put("not", "Not")
      .put("nothing", "Nothing")
      .put("null", "Null")
      .put("on", "On")
      .put("or", "Or")
      .put("private", "Private")
      .put("public", "Public")
      .put("resume", "Resume")
      .put("select", "Select")
      .put("single", "Single")
      .put("static", "Static")
      .put("step", "Step")
      .put("sub", "Sub")
      .put("then", "Then")
      .put("to", "To")
      .put("true", "True")
      .put("type", "Type")
      .put("until", "Until")
      .put("while", "While")
      .put("with", "With")
      .put("xor", "Xor")
      .build();

  private static final ImmutableSet<String> LOGICAL_OPS = ImmutableSet.of(
      "And",
      "Or",
      "Xor");

  private static final ImmutableSet<String> RELATIONAL_OPS = ImmutableSet.of(
      "<",
      ">",
      "=",
      "<=",
      ">=",
      "<>",
      "Is",
      "IsNot",
      "Like");

  private static final ImmutableSet<String> ARITHMETIC_OPS = ImmutableSet.of(
      "+",
      "-",
      "*",
      "/",
      "\\",
      "Mod",
      "&",
      ">>",
      "<<");

  private static final int MAX_PEEK_LIMIT = 1000;

  private final GlobalState globalState;

  private String original;
  private String converted;
  private int peekCount;
  private String comment;

  private String tokenType;
  private String token;

  Line(GlobalState globalState) {
    this.globalState = globalState;
  }

  Line() {
    this(null);
  }

  /**
   * Step over expected token.
   */
  @SuppressWarnings("unused")
  void eat(String expected) {
    String token = getToken(true);
    // TODO(nikhil): It _seems_ like this should work. Figure out why it causes
    // tests to fail.
    //if (!token.equals(expected)) {
    //  throw new ParseException("Expected token [" + expected + "], got [" + token + "] instead");
    //}
  }

  /**
   * Returns a balanced-paren sequence of tokens. Called with ( as peek token.
   * Includes the parens in result. Tries to convert array(i) to array[i].
   */
  // Why is this different from exprlist?
  // These items don't have to be expressions. Not clear the separation is necessary, however.
  String getBalancedParentheses()  {
    StringBuilder balanced = new StringBuilder(getToken(true));
    while (!peek().equals(")") && !peek().isEmpty()) {
      if (peek().equals("(")) {
        balanced.append(getBalancedParentheses());
      } else if (peek().equals(".")) {
        balanced.append(globalState.getWithName()).append(getToken(true)).append(getName());
      } else if (tokenType.equals("ID")) {
        String name = getName();
        balanced.append(name);
        if (globalState.isArrayName(name) && peek().equals("(")) {
          balanced.append(setBrackets(getBalancedParentheses()));
        }
      } else {
        balanced.append(ConverterUtil.fixOperators(getToken(true)));
      }
    }
    balanced.append(getToken(true)); // adds terminating )
    return balanced.toString();
  }

  /**
   * Tacked onto output lines in gen().
   */
  String getComment() {
    return comment;
  }

  /**
   * Returns current token.
   */
  String getCurrentToken() {
    return token;
  }

  /**
   * Returns next expression from input. Handles .names, nested constructs etc.
   * Drops whitespace after , This has 6-7 levels of precedence. From the
   * bottom:
   * :=, logical ops, negation, relational ops, arith, unary, expon.
   * This isn't complete but it's simpler; assumes that the input is already
   * sensibly parenthesized so it doesn't generate spurious parens.
   */
  String getExpression()  {
    String expression = getArg();
    if (peek().equals(":=")) { // named argument
      getToken(true);
      expression = "\"" + expression + " :=\", " + getLogic();
    }
    return expression;
  }

  /**
   * Returns whatever is left of the line.
   */
  String getLine() {
    return converted.trim();
  }

  /**
   * Returns next name from input, with . expanded, () => [], etc.
   */
  String getName()  {
    if (peek().equals(".")) {
      return globalState.getWithName() + getToken(true) + getName();
    }
    if (!tokenType.equals("ID")) {
      return "";
    }
    StringBuilder name = new StringBuilder(getToken(true));
    if (peek().equals("(")) { // e.g., Range("A3")
      String expressions = getExpressionList();
      if (globalState.isArrayName(name.toString())) {
        expressions = setBrackets(expressions);
      }
      name.append(expressions);
    }
    if (peek().equals("(")) { // e.g., Range("A1")(cnt)...
      name.append(getExpressionList());
    }
    while (peek().equals(".")) { // e.g., Range("A3").Selection.Cells(1,j)
      name.append(getToken(true));
      name.append(getName());
    }
    return name.toString();
  }

  /**
   * Returns the trimmed original input.
   */
  String getOriginal() {
    return original.trim();
  }

  /**
   * Returns whatever remains of the current input line.
   */
  // Perhaps should do get_expr or the like to handle array subscripting?
  String getRest()  {
    StringBuilder rest = new StringBuilder();
    while (!peek().isEmpty() && !peek().equals(ConverterUtil.EOF)) {
      rest.append(ConverterUtil.fixOperators(getToken(true)));
    }
    return rest.toString();
  }

  /**
   * Returns next token.
   */
  // Not sure that string process is right yet, since a quoted string is
  // canonicalized by constructor but found here by a simple RE that doesn't
  // handle \" within a string.
  String getToken(boolean advance)  {
    if (original.trim().equals(ConverterUtil.EOF)) {
      return ConverterUtil.EOF;
    }

    converted = converted.trim();
    for (Pattern type : TYPES.keySet()) {
      Matcher m = type.matcher(converted);
      if (m.lookingAt()) {
        tokenType = TYPES.get(type);
        token = converted.substring(m.start(), m.end()); // the matching part

        if (tokenType.equals("TOSS")) {
          converted = converted.substring(token.length()); // left for next time
          continue;
        }
        if (tokenType.equals("STR")) { // re for strings isn't right so clean up
          token = getStr(converted);
        }
        if (tokenType.equals("DATE")) { // replace # by "
          token = "\"" + token.substring(1, token.length() - 1) + "\"";
        }
        if (tokenType.equals("HEX")) {
          token = token.replaceFirst("&H", "0x");
        }
        if (token.equals("!")) { // maybe too exuberant?
          token = ".";
        }
        if (advance) {
          converted = converted.substring(token.length()); // left for next time
        }
        if (tokenType.equals("NUM")) { // get rid of vb type indicator
          token = token.replaceFirst("[&#]$", "");
        }

        return toUpperCase(token);
      }
    }

    throw new ParseException("Unknown token, can't parse: " + converted);
  }

  boolean hasComment() {
    return !getComment().isEmpty();
  }

  boolean hasToken() {
    return !getCurrentToken().isEmpty();
  }

  /**
   * Tries to isolate a comment if any, while partially coping with horrors like
   * single quotes inside double, quotes in comments, etc.
   */
  Line parseLine(String line) {
    this.original = line;
    this.peekCount = 0;
    this.comment = "";
    this.converted = "";
    this.tokenType = "";

    while (!line.isEmpty()) {
      char first = line.charAt(0);
      if (first == '\'') {
        comment = line.substring(1);
        break;
      } else if (first == '"') {
        // getstring returns the quoted string and the residue as a tuple
        // in java, getstring can append to _str and return residue
        line = getString(line);
      } else if (first == '[') {
        // getbrack returns the quoted string and the residue as a tuple
        // in java, getbrack can append to _str and return residue
        line = getBracketed(line);
      } else {
        converted += first;
        line = line.substring(1);
      }
    }
    converted = canonicalize(converted.trim());

    return this;
  }

  /**
   * Returns next token without consuming it.
   */
  String peek() {
    if (original.trim().equals(ConverterUtil.EOF)) {
      return ConverterUtil.EOF;
    }

    // kludge to detect potential bad input
    // alternative is to decorate all calls with "and cur.peek() != EOF"
    ++peekCount;
    if (peekCount > MAX_PEEK_LIMIT) {
      throw new ParseException("Looping because of illegal input: " + original);
    }

    return getToken(false);
  }

  /**
   * Returns the type of the next token. (Assumes that peek() has just been
   * called).
   */
  String peekTokenType() {
    return tokenType;
  }

  /**
   * Add outer parens if !s appears to need them.
   */
  private String addParen(String str) {
    if (str.matches(".*[-+*/%^<>=!&|].*")) { // watch out: needs unanchored
      return "(" + str + ")";
    } else {
      return str;
    }
  }

  /**
   * Canonicalize some lexical stuff, like Public, that will simplify
   * subsequent processing.
   */
  private String canonicalize(String str) {
    return str.replaceAll("Property Get ", "Function Get")
              .replaceAll("Property Let ", "Function Let")
              .replaceAll("Property Set ", "Function Set")
              .replaceAll("End Property", "End Function")
              .replaceAll("(Public|Private|Friend) +Sub", "Sub")
              .replaceAll("(Public|Private|Friend) +Function", "Function")
              .replaceAll("(Public|Private|Friend) +Dim", "Dim")
              .replaceAll("(Public|Private|Friend) +Global", "Global")
              .replaceAll("(Public|Private|Friend|Global) +Const", "Const")
              .replaceAll("(Public|Private|Friend) +Declare", "Declare")
              .replaceAll("(Public|Private|Static)", "Dim");
  }

  /**
   * Returns next expression from input. Handles .names, nested constructs etc.
   * Drops whitespace after ,. This has about 6 levels of precedence, from the
   * bottom:
   * logical operators, negation, comparison, arith, unary, expon.
   * This isn't complete but it's simpler; assumes that the input is already
   * sensibly parenthesized.
   */
  private String getArg()  {
    StringBuilder arg = new StringBuilder(getLogic());
    while (LOGICAL_OPS.contains(peek())) {
      String op = ConverterUtil.fixOperators(getToken(true));
      arg.append(op).append(getLogic());
    }
    return arg.toString();
  }

  private String getArithmeticOp()  {
    String op = getFactor();
    while ("^".equals(peek())) {
      getToken(true);
      op = "exp(" + op + ", " + getArithmeticOp() + ")";
    }
    return op;
  }

  /**
   * Collect [...]
   */
  private String getBracketed(String str) {
    StringBuilder inside = new StringBuilder();
    String bracketed = str.substring(1);
    while (true) {
      char first = bracketed.charAt(0);
      if (first == ']') {  // the end
        bracketed = bracketed.substring(1);
        break;
      } else if (first == '!') {
        inside.append(".");
        bracketed = bracketed.substring(1);
      } else {
        inside.append(first);
        bracketed = bracketed.substring(1);
      }
    }
    converted += "Range(\"" + inside.toString() + "\")";
    return bracketed;
  }

  private String getCompare()  {
    StringBuilder expr = new StringBuilder(getUnary());
    while (ARITHMETIC_OPS.contains(peek())) {
      String op = ConverterUtil.fixOperators(getToken(true));
      expr.append(op).append(getUnary());
    }
    return expr.toString();
  }

  /**
   * Returns a list of expressions. Called with ( as peek token. Includes the
   * parens in result. Tries to convert array(i) to array[i]. Flags empty
   * exprs.
   */
  // Note: Logic is too convoluted: getFactor() should be smarter.
  private String getExpressionList()  {
    StringBuilder expressions = new StringBuilder(getToken(true)); // "("
    while (!peek().equals(")") && !peek().isEmpty()) {
      if (peek().equals(",")) { // empty expr
        expressions.append(ConverterUtil.EMPTY).append(getToken(true)).append(" ");
        if (peek().equals(")")) { // empty expr
          expressions.append(ConverterUtil.EMPTY);
        }
        continue;
      }
      expressions.append(getExpression());
      if (peek().equals(",")) {
        expressions.append(getToken(true)).append(" ");
        if (peek().equals(")")) { // empty expr
          expressions.append(ConverterUtil.EMPTY);
        }
      }
    }
    expressions.append(getToken(true)); // adds terminating )
    return expressions.toString();
  }

  /**
   * Returns single entity -- number, name, or (expr). This also returns things
   * like comma, which is a botch.
   */
  private String getFactor()  {
    StringBuilder expr = new StringBuilder();
    String peek = peek();
    if (tokenType.equals("ID")) {
      String name = getName();
      expr.append(name);
      if (globalState.isArrayName(name) && peek().equals("(")) {
        String bp = getBalancedParentheses();
        expr.append(setBrackets(bp));
      }
    } else if (tokenType.equals("NUM")) {
      expr.append(getToken(true));
    } else if (tokenType.equals("STR")) {
      expr.append(getToken(true));
    } else if (peek.equals(".")) { // .name
      expr.append(globalState.getWithName()).append(getToken(true)).append(getName());
    } else if (peek.equals("Not")) { // BUG?
      expr.append(getLogic());
    } else if (peek.equals("(")) {
      expr.append(getToken(true)).append(getExpression()).append(getToken(true));
    } else {
      expr.append(getToken(true));
    }
    return expr.toString();
  }

  private String getLogic()  {
    StringBuilder expr;
    if (!peek().equals("Not")) {
      expr = new StringBuilder(getNotOp());
    } else {
      expr = new StringBuilder();
    }
    while (peek().equals("Not")) {
      String op = ConverterUtil.fixOperators(getToken(true));
      expr.append(op).append(addParen(getLogic()));
    }
    return expr.toString();
  }

  private String getNotOp()  {
    String expr = getCompare();
    while (RELATIONAL_OPS.contains(peek())) {
      String op = ConverterUtil.fixOperators(getToken(true));
      if (op.equals("Like")) {
        expr = "Like(" + expr + "," + getCompare() + ")";
      } else {
        expr += op + getCompare();
      }
    }
    return expr;
  }

  // TODO(nikhil): Rename getStr() and getString()
  /**
   * Returns the real string, skipping embedded \"'s.
   */
  private String getStr(String str) {
    int i = 1;
    while (i < str.length()) {
      if (str.substring(i, i + 1).equals("\"")) {
        break;
      }
      if (str.substring(i, i + 1).equals("\\")) {
        ++i;
      }
      ++i;
    }
    return str.substring(0, i + 1);
  }

  /**
   * Collect quoted string, handle "" and \.
   */
  private String getString(String str) {
    StringBuilder parsed = new StringBuilder(str.substring(0, 1)); // the " at the front
    String input = str.substring(1);
    while (true) {
      if (input.substring(0, 1).equals("\\")) {
        parsed.append("\\").append(input.substring(0, 1));
        input = input.substring(1);
      } else if (input.substring(0, 1).equals("\"") &&
          (input.length() > 1) &&
          input.substring(1, 2).equals("\"")) {
        parsed.append("\\\"");
        input = input.substring(2);
      } else if (input.substring(0, 1).equals("\"")) {
        parsed.append("\"");
        input = input.substring(1);
        break;
      } else {
        parsed.append(input.substring(0, 1));
        input = input.substring(1);
      }
    }
    converted += parsed.toString();
    return input;
  }

  private String getUnary()  {
    String op = "";
    while (peek().equals("+") || peek().equals("-")) {
      op += getToken(true);
    }
    String expr = getArithmeticOp();
    expr = op + expr;
    return expr;
  }

  /**
   * Set brackets in s to convert from (...) to [...]""".
   */
  // Note: This is probably too aggressive: won't work if there are nested
  // commas. e.g., in function calls in subscripts, or in strings.
  private String setBrackets(String str) {
    String input = str.substring(1, str.length() - 1);
    if (input.indexOf('(') == -1) {
      input = input.replaceAll(", *", "][");
    }
    return "[" + input + "]";
  }

  /**
   * Canonicalizes the case of a likely keyword.
   */
  private String toUpperCase(String str) {
    String lowerCase = str.toLowerCase();
    if (KEYWORDS.containsKey(lowerCase)) {
      return KEYWORDS.get(lowerCase);
    } else {
      return str;
    }
  }
}