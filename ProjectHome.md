Java library for converting VB code to the equivalent JavaScript. The translator tries to do a best-effort syntactical conversion. It does not analyze/parse the input in any way, though this might be worth doing in the future. The translator can only handle Visual Basic upto VBA 6, and there is no current plan to support VB.NET. A very high-level list of constructs that can be translated are listed below:
Comments.
Variable declarations.
Constants (including dates and file handles).
Arrays (including the correct initialization of multi-dimensional arrays).
Expressions.
Control flow statements (loops and decision blocks).
Subroutines (in VB, these don't return a value) and functions (these return a value).
Exceptions (since VB's error handling allows for goto jumps to arbitrary sections of code, this is done on a best-effort basis).
Translation of VBA Types to JS classes.
VBA-isms like IIf (immediate If), the Array function and multi-line continuations using an underscore.