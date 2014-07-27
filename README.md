VariableVisualizer
==================

A developer tool that will help take the guesswork out of refactoring. 

The program will parse VB.NET code looking for local variable definitions and usages.

You will then see a printout of which local variables are still needed (by line number).

The objective is to find a point where the number of local variables in use is low--you can then break up the function so
that everything before the low point is one child function, and everything after becomes a second child function.

You will then only need to pass a minimum of variables from one half to the other. There is also a good chance that breaking
up the master function in this way will result in 2 or more chunks that are conceptually coherent and distinct.
