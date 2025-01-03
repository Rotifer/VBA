# Excel VBA

I still occasionally use Excel VBA although I now try to use the Excel formula language and LAMBDAs if possible. 
However, sometimes it is easier and cleaner to implement something in VBA. 

## Set Operations - Implementing them in VBA 😬

Something I find painful in Excel is the absence of _set_ logic. In SQL we have operators such as UNION, EXCEPT, and INTERSECT and they are super useful. 
I have seen implementations of these in Excel using combinations of FILTER, MATCH and so on in LAMBDAs. They implementations might be clever 
but they are not clear or intuitive to me. 

As a challenge, I decided to try to implement them in VBA using dictionaries. The implementations are not yet complete but I am making progress.
Once I have finished, I will post an extensive blog on the subject.

### Guide to the files

I have uploaded two __.bas_ module files and two _.cls_ class files.

1. __clsSet.cls__: This contains the set logic code and is the key file.
1. __clsTestSet__: Testing code for the main class file. Trying to implement a poor man's unit testing in VBA 🤪 . I have tried RubberDuck but not working.
1. __modTest_clsSets__: Creates the class instance for testing and calls the test routines.
1. __SetUDFs.bas__: User-defined functions that use the main class and that can be called in Excel.

### TODO - lots!

- The code is not yet fully implemented or documented.
- I will write a longer explanation of how this works in a blog once I have finished the coding.


[Now blogging here](https://rotifer.github.io/) they


