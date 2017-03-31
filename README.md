# EQAN-The Equation Analyser
Solves for the unknown variable of an equation or equation set irrespective of its location with respect to the "=" sign

# Setup
EQAN was initially written in Microsoft Excel’s VBA and later a standalone application was written in VB.Net using Visual Studio. Both the excel version and .net version are available for the user.  

# Usage
Consider an example : a=b+sin(c/d) where value of variable "c"  is unknwon.

![solarized dualmode](https://github.com/vaniceast/EQAN/blob/master/example.gif)

# Mathematical Operations
EQAN supports the following mathematical operations: 
* Addition/Subtraction (+/-)
* Exponentiation (^)
* Division (/)
* Multiplication (*)
* Natural Logarithm (LN)
* Exponential Function (EXP)
* Square Root (SQRT)
* Trigonometric Functions: 
  * Sine (SIN)
  * Cosine (COS)
  * Tangent (Tan)
  * Cosecant (ASIN)
  * Secant (ACOS)
  * Cotangent (ATAN)

# Algorithm 
Click on the link : [Algorithm.pdf](https://github.com/vaniceast/EQAN/blob/master/EQAN.pdf)

# Limitations
EQAN is far from perfect. Most of its imperfections are the results of time constraints and the fact this is my first time programming in visual basic. The excel version has some advantages over the standalone application and vice versa. Some limitations include:
* The LHS of the “=” sign should be a single variable. So an equation like "a/b=c/d" will result in an error.
* Trigonometric functions are evaluated in degrees only. Results in radians are not possible.
* The algorithm for selecting and arraigning equations results in an error for certain combinations.
* The UI should have been better. Shortcut keys are yet to be added.

# Contact
Feel free to contact me with queries, remarks, suggestions, etc. (vaniceast@outlook.com)
