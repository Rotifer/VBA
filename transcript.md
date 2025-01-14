Create custom functions in Excel
Selecting transcript lines in this section will navigate to timestamp in the video
- [Curt] Microsoft Excel is a powerful and versatile tool for analyzing data within your business or organization. The new LAMBDA and LET 
functions make it possible for you to create powerful and portable custom calculations you can use just like built-in Excel functions. In this 
course, I will show you the Excel formula language and provide real-world examples to demonstrate how you can apply the incredible power of 
custom functions to your Excel workbooks. I'm Curt Frye. Join me at LinkedIn Learning for an introduction to the essential skills that will let 
you unlock the power and flexibility that comes with creating your own custom functions in Microsoft Excel.

What you should know before starting
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Thanks for your interest in this course. Before I get started, I would like to give you some information about what you should 
know to get the most out of this course. First, you should definitely know how to create formulas in Excel. You don't need to be the most 
advanced user, but the more you know, the more you will get out of this course, at least at the start. You should also know a lot about your 
business and how to determine what information you need to solve a problem. Creating custom functions lets you streamline your work, so it's 
important to know what it is you need to do. Also, know that Lambda works best with other functions. You do not need to create everything by 
hand. Use the built-in functions in Excel, and you'll save yourself a lot of time. Also, be aware that Excel from Microsoft 365 updates 
regularly. That means that your screen might not look exactly like mine, but everything will still be there. And what's even more exciting is 
that new functions are added all the time, so you'll have even more capability than what I have today. Also, be aware that this is a course for 
beginning and intermediate users, so there will be some repetition to emphasize important concepts. Also, I will create named functions for 
some, but not all, examples. It all depends on how long I feel a particular movie is running. And I also have a chapter on scenarios that 
encourage you how to think about your own solutions. I have built some custom calculations for problems that I have spoken about in the past, 
and hopefully that will provide you with incentives and insight into the type of functions that you can create for yourself.

Explore custom functions in Excel VBA
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] The ability to create custom functions in Excel within a worksheet is an important step forward. However, I think it's important 
to know where we have come from. So in this movie, I will show you the old method of creating custom functions using Visual Basic for 
Applications, or VBA. My sample file is 01_01_VBAFunction, and that is a macro-enabled Excel workbook you can find in the chapter one folder of 
the exercise files collection. So the goal for this worksheet is to calculate a commission. I have a set of sales starting in cell B4 and then I 
want to calculate commissions. I can do that using a function that I've created using Visual Basic for Applications. To move to the Visual Basic 
editor, I will press alt-F11 and then I will go over to the left side in the project window and double-click Module 1. And there you can see the 
function that I created. I won't go through it in depth but you can see at the top, I have a function and then its name is Commission. That's 
what we'll use as the function when we create the formula. It takes a currency value as an input and then it looks to see if sales are less than 
500. If so, 5% commission, greater than 500 but less than 1,000, then 6%, anything 1,000 or greater, 7%, and, at the bottom, it returns the 
value that was calculated. I'll press alt-F11 to move back and now I can create my formula. So I will select cells C4 through C6, equal, and 
then I'll type in commission which is the new function that I defined, and then I'll type b4. That'll be for the first formula, and then 
Control+Enter, and you see, I get the commissions on each of those sales. One of the limitations of working with VBA might have happened to you 
when you were trying to open this file. Your macro security settings might have either prevented it silently or they might have flashed a 
warning, or your company's IT policy might not have allowed the file to open on your system. And those cases indicate the possibility that a 
macro could be written that would be harmful to your computer. So rather than do that, the Excel product team at Microsoft has allowed us to 
create our own functions using the Excel formula language and that is what we'll focus on from here on out.

Describe the Excel formula language
Selecting transcript lines in this section will navigate to timestamp in the video
- [Narrator] You can create custom functions in Excel now, using the Excel formula language. Existing functions accept inputs and generate 
results. And here you have three different ways to create a SUM formula. You use the SUM function, looking at the range, A1 through A15 or the 
individual cells A1 through A15 or two separate ranges, A1 to A5 and A11 to A15. So ignoring cells A6 to A10. The problem at least if you want 
to create your own functions is that the calculation is not visible to the user. One benefit of using LAMBDA is that you can define the inputs 
and the calculation and make them explicit. That is make them visible within your workbook. Here is what a LAMBDA function might look like. 
Amount and rate are inputs and amount times one plus rate is the calculation. And at the end you see that I have A2 and B2 in their own set of 
parentheses. Those are the cells that provide values for the amount and the rate that are used in the calculation. You can also use the Name 
Manager to create a new function. For example, you might name your function ApplyGrowth and accept the amount and rate as part of the formula 
that you create. With that as background, I'll switch to Excel and show you what it looks like within a workbook. I have switched over to Excel 
and the name of the workbook I'm using is 01_02_Describe and you can find it in the chapter one folder of the exercise files collection. In this 
workbook, I have two starting amounts and growth rates. And in cell C2 you can see that I have a LAMBDA function although I have it as text 
because I didn't put an equal sign in front. That basically creates the function that I showed you earlier in this moving. So if I want to see 
how that would work I will edit the formula in cell C2 and add an equal sign in front. So now it's a formula, enter and we get the value of 
104,750. And if we take a look back, we'll see that once again this function accepts its inputs from A2 and B2. If I were to take those away by 
editing them out and press enter, I would get a calc error. And that's because the formula as is defined right now doesn't have any inputs. So 
I'll press Control Z to bring them back. I've also created a function that will allow me to use this LAMBDA as a named formula. So if I go to 
the formulas tab and click Name Manager you can see in the Name Manager that I have a name called ApplyGrowth. And that is actually a custom 
function. And if you think it's weird that you have to go to the Name Manager, which is usually used to name ranges, to create a custom function 
then the Excel product team agrees with you. It's just this is the best solution they could find for right now. So it'll probably change in the 
future. But for now, this is where you can find it. And if you look at the bottom, you'll see in Refers to: that we have the formula that I 
defined before LAMBDA and then it accepts amount and rate and then it multiplies the amount by one plus the rate. And also because we are 
accepting amount and rate as inputs we don't have the A2 and B2 at the end. Okay, with that in mind, I'll click close. And then in cell C3 I'll 
type equal and ApplyGrowth as our function. And the amount is in cell A3, The rate is in B3, rare parenthesis to close and enter. And we get our 
result of $212,000.

Describe the goals of formula programming
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] When I mention Lambda to Excel users who haven't heard of it, they are interested in knowing what the goals are for formula 
programming in Excel or functional programming. And there are quite a few of them, but the four I want to focus on start with the ability to 
create custom functions used throughout a workbook. You can define a function, such as applied growth rate. And because you created it, you know 
exactly what it's going to do. It also allows the function to be copied to other workbooks. You can go to the name manager, copy the Lambda 
definition from the named range that refers to your custom function, and copy it into another workbook. Now, of course, if you want your 
formulas to work in that other workbook, you need to make sure that you give it exactly the same name in the name manager, but that's something 
that you'll figure out very quickly in case you make a mistake. The overarching goal is to create an analytical platform that does not rely on 
code. Visual Basic for Applications, the macro programing language, is incredibly powerful. But because it can do almost anything within Excel, 
it means that it can be used to violate your computer security. Macro viruses were a real problem for a long time, so that means the companies 
clamp down on your ability to use macros and VBA. That's understandable, and the formula programming language through Lambda takes away that 
vulnerability. But finally, going into the future, using Lambda will enable advanced applications, such as machine learning. The map and reduced 
functions are extremely powerful and are frequently used in machine learning applications. So even though I don't go into any of that in this 
course, if you want to use Excel as an advanced analytical platform, then you'll be able to do so going forward. Perhaps not right at the 
moment, at least not as much as you would like to, but going forward in the future, I think you'll see those capabilities come online very 
quickly. Manage formulas on the Formula Bar
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] When you create formulas, including those that use lambda and let functions you do so on the formula bar. In this movie, I will 
show you how to work with the formula bar in ways you might not have in your everyday Excel use. My sample file is 0104 formula bar and you can 
find it in the chapter one folder of the exercise files collection. I have created an IFS, that is a multiple condition IF formula and I am 
rounding the value up to zero digits so it will be a whole number. And you can see the formula, which is in cell B4 currently displayed on the 
formula bar and it's a long one. So let's say that I want to make that formula easier to read by expanding the formula bar. I can do that a 
little bit by going to the right edge of the formula bar and clicking its down arrow here. And that expands it to three lines of text. And one 
way that you can use that extra space is to add line breaks within the formula. So for example, I will click to the right of the first left 
parentheses and then press alt enter. Doing so adds a line break and I could do the same thing to the left of A4 and to the right of the next 
left parenthesis. So alt enter and I have run out of space but what I can do is hover the mouse pointer over the bottom edge and when it turns 
to an up and down pointing arrow, I can drag down and you see that it's adding space as I go. Now it's taking that space away from the worksheet 
itself. So you see that I only have a couple of rows left at the bottom, especially because I'm zoomed in. So you need to balance your need for 
space in the formula bar versus what you need to see in the worksheet. And now I will add lines. So I'll click to the right of the next comma 
after A4 times 2.5. So I have my condition and then my result. So alt enter, and then I'll go to the right of the comma after the next 
condition, alt enter. And I'll do the same thing here and I will save you the commentary because the reasoning and the actions are all the same. 
So there we go. And then I will go to the end of ROUNDUP. And there we have it. So this is one way to lay it out. And because I have some space 
left over I can drag the formula bar back up so it only shows as much as I need. And now if I press enter, the formula's entered and the formula 
bar stays in the same configuration. If I want to close the formula bar, that is to reduce its space, then I can click the up arrow here and 
we're all the way back up. But of course the issue is that I only see the first line and so I see ROUNDUP and then left parentheses and nothing 
else. So to see the rest of it, I need to expand the formula bar and it goes back to the condition that it was in before. So if you're creating 
a long formula that is hard to read you can use the expanded formula bar and alt enter to make it easier to read. And if it seems like you're 
not seeing all the formula on the formula bar, then you can expand it. And I bet that you'll see more lines laid out like you see here.

Define a function using LAMBDA
Selecting transcript lines in this section will navigate to timestamp in the video
- One great new capability in Excel is that you are now able to define a custom function in the program using lambda. In this movie, I will show 
you what that process looks like. My sample file is oh two oh one create lambda, and you can find it in the chapter two folder of the exercise 
files collection. In this workbook, I have a worksheet that displays values. So we have a starting amount, and then the growth rate. And then to 
the left, we have a number of years. If you look at the formula in cell B three, you'll see that the starting amount for the year 2024 is the 
ending amount in D two for 2023. So we're able to continue on with our growth. If I want to create a lambda function to calculate the interest 
that's been earned, then I can type an equal sign; I'll start in D two, and lambda is the name of the function. I need to provide it with two 
separate inputs, So I'll have start, underscore amount, then a comma, and then G underscore rate. So we have our starting amount and our growth 
rate, then a comma, and the calculation is start, amount and you can see it shows up in the auto complete list, and we'll multiply that by in 
parentheses one plus the growth rates that's G underscore R A T E. Then two right parentheses, and it's easy to think that we're done, but we'll 
see an error when I press enter. And that error is calc, and what that indicates is that there's a calculation engine error and the problem, if 
you look at the formula in the formula bar, is that we don't have any inputs. We just have definitions. So if I want to add those inputs, I can 
double click D two. And then in a new set of parentheses I need to type in the cell addresses for the inputs. So the starting amount is in B 
two, then a comma, and the growth rate is in C two. So I have those and parentheses, and enter, and there's my test value. I get the ending 
amount of 106,000 and then that's the starting amount for 2024. Now I'll go to sell D two and double click the fill handle at the bottom right 
corner and that copies the growth rate all the way down. So I have the ending amounts and I have the formulas that I need. So those are the 
basics of creating a function using lambda within Microsoft Excel.

Assign a function name to a LAMBDA
Selecting transcript lines in this section will navigate to timestamp in the video
- After you create a Lambda function you can assign it a name using the name manager. In this movie, I will demonstrate how to do that. My 
sample file is 0 2 0 2 name Lambda, and you can find it in the chapter two folder of the exercise files collection. I have cell D2 selected at 
the moment. I'll go ahead and double click it so we can see the formula that is inside of it. This formula, which was created using the Lambda 
function, has a start amount and a growth rate as the two arguments it receives. And then it calculates a result that is the amount earned or 
the amount at the end of the year by multiplying the start amount by one, plus the growth rate. And then at the end, again, we have inputs from 
cell B2, which is the first input, the start amount, and then C2 is the growth rate. If I want to create a named function based on this lambda, 
then I can copy everything up to the inputs at the end. So I will copy everything there. So I have the equal sign and the second right 
parenthesis at the end. So I'll go ahead and press control C to copy then escape to exit the cell editing mode. Now I need to go to the name 
manager, So I'll go to the formulas tab and then click name manager. And then in the name manager, I'll click new and I can enter in a new name. 
And for this, I will call it total amount. And one thing to note is that I used total with a capital T and then amount with a capital A, But all 
the other letters are lowercase. The reason I did that is because Excel's built-in functions have all capital letters as their function names, 
and this is a way of reminding myself that this is a function that I created. And the usual rules for named ranges apply. You can't start it 
with a number, you can't have any spaces, and so on. The scope is the workbook, and for the comment, I will just say calculate total amount 
after interest. And then, I'll go down to the refers to box and replace the existing text with the formula that I copied. Everything looks good, 
so I'll go ahead and click okay. And there I have my Lambda function called total amount. So I'll click close, and then with cell D2 still 
selected, I'll type equal total amount. And you can see that I have a little bit of indicator text saying that it calculates the total amount 
after interest. So I'll press tab, and then I also have a tool tip that asks for the two arguments. So I have the start amount, which for this 
row is in B2 and then the growth rate is in C2, right parentheses, and enter. And I get the same value as before, but as you can see, the 
formula is much more readable. Then I can double-click the fill handle at the bottom right of cell D2, and I get the same results, but as 
before, I have total amount that accepts two arguments from the cells in this particular row.

Define a variable using LET
Selecting transcript lines in this section will navigate to timestamp in the video
- A Lambda statement, lets you create, reusable custom functions in Excel. A Let statement, by contrast gives you the ability to define 
variables that can be used within those Lambda functions. In this movie, I will show you the first part of the process, which is how to define a 
variable using Let. My sample file is, 0 2 0 3 Let, and you can find it in the chapter two folder, of the Exercise Files collection. In this 
workbook I have a list of cases and also varieties. So let's say that we're in a warehouse, and I have cases, that contain varieties of seed, 
that I want to use on a farm, and they've been distributed randomly. And you see that in column B. Case one contains Variety four, so does 
column two. Case three contains Variety three, and you can see the rest. Let's say there are more than four varieties possible, and I want to 
have a list of all the unique varieties that actually were selected for this particular case. To do that, I can go to Cell D four and then type 
equal. And the function I'll use is fairly new, and that is unique. And this finds the unique values, within a data set. And my array is B four 
to B 12. I don't need to change anything else, 'cause I'm just looking for unique values, right Parentheses and enter. And the result of the 
formula starts in cell D four, and then spills to cells D five through D seven. So I have four unique varieties. This isn't all that easy to 
read though, because they are not listed in alphabetical order. I can change that by using the sort function, which is also relatively new. So I 
will double click cell D four, and I will edit the formula, so that I will start with the sort function. I have my array, which is the result of 
the unique formula, that I created earlier, then a comma. My sort index will just be the values. I don't need to define which column it is 
within the data set, and the sort order will be ascending. So lowest value first, highest value at the end. So I'll type a one and I don't need 
to indicate by column, so I'll just type a right parentheses and enter. And then I get an alphabetical list, of the unique varieties. If I were 
to change the value in B four, from variety four to variety five. So I'll do that now and enter, you can see that the formula results are 
updated. I'll press control Z, to go back to what I had before. Now let's assume, that I want to create a calculation that will provide a 
variable using Let. That calculates the number of unique varieties within a data set. So what I'm doing here, is calculating the percentage of 
the four unique varieties, within the nine total cases that I have in the warehouse. So I'll go to cell F four, type in equal sign, and I'll 
start with Let, which is the function, that we use to store intermediate calculations in a variable. The name will be Num, N U M, that'll be the 
first value. This variable contains the total number, of items in the set, so it'll be nine, then a comma. And the value for this will be to 
count, all the non empty cells in the range B four to B 12. So I'll do count A, which counts the number of cells that are not empty, left 
parenthesis and then the range is B four to B 12. All right, close out those parentheses in the comma. The second name that we'll use is U N Q, 
and that is short for unique, then a comma. And we'll count the number, of unique values within the same range. So that will be count A as 
before, left parentheses unique, and the range is B four to B 12. I'll just type that in. Then I have my right parenthesis there. And because I 
have unique nested, within Count A I need to type another right parenthesis, then a comma. And then finally the calculation. And that will just 
be the number of unique values. So U N Q, divided by the total number of items, and that is N U M. And you can see that both U N Q, and N U M 
appeared in the formula, auto complete list. So I'll do that, and then a right parenthesis, and everything looks to be balanced out. So I'll 
press enter, and there you have it. It counts all the unique varieties, divided by the total number of items in the set, and gives me a percent 
unique of 44. And again, if I edit the value in B four, from variety four to variety five and enter, it goes up to 56%.

Define a variable using LET
Selecting transcript lines in this section will navigate to timestamp in the video
- A Lambda statement, lets you create, reusable custom functions in Excel. A Let statement, by contrast gives you the ability to define 
variables that can be used within those Lambda functions. In this movie, I will show you the first part of the process, which is how to define a 
variable using Let. My sample file is, 0 2 0 3 Let, and you can find it in the chapter two folder, of the Exercise Files collection. In this 
workbook I have a list of cases and also varieties. So let's say that we're in a warehouse, and I have cases, that contain varieties of seed, 
that I want to use on a farm, and they've been distributed randomly. And you see that in column B. Case one contains Variety four, so does 
column two. Case three contains Variety three, and you can see the rest. Let's say there are more than four varieties possible, and I want to 
have a list of all the unique varieties that actually were selected for this particular case. To do that, I can go to Cell D four and then type 
equal. And the function I'll use is fairly new, and that is unique. And this finds the unique values, within a data set. And my array is B four 
to B 12. I don't need to change anything else, 'cause I'm just looking for unique values, right Parentheses and enter. And the result of the 
formula starts in cell D four, and then spills to cells D five through D seven. So I have four unique varieties. This isn't all that easy to 
read though, because they are not listed in alphabetical order. I can change that by using the sort function, which is also relatively new. So I 
will double click cell D four, and I will edit the formula, so that I will start with the sort function. I have my array, which is the result of 
the unique formula, that I created earlier, then a comma. My sort index will just be the values. I don't need to define which column it is 
within the data set, and the sort order will be ascending. So lowest value first, highest value at the end. So I'll type a one and I don't need 
to indicate by column, so I'll just type a right parentheses and enter. And then I get an alphabetical list, of the unique varieties. If I were 
to change the value in B four, from variety four to variety five. So I'll do that now and enter, you can see that the formula results are 
updated. I'll press control Z, to go back to what I had before. Now let's assume, that I want to create a calculation that will provide a 
variable using Let. That calculates the number of unique varieties within a data set. So what I'm doing here, is calculating the percentage of 
the four unique varieties, within the nine total cases that I have in the warehouse. So I'll go to cell F four, type in equal sign, and I'll 
start with Let, which is the function, that we use to store intermediate calculations in a variable. The name will be Num, N U M, that'll be the 
first value. This variable contains the total number, of items in the set, so it'll be nine, then a comma. And the value for this will be to 
count, all the non empty cells in the range B four to B 12. So I'll do count A, which counts the number of cells that are not empty, left 
parenthesis and then the range is B four to B 12. All right, close out those parentheses in the comma. The second name that we'll use is U N Q, 
and that is short for unique, then a comma. And we'll count the number, of unique values within the same range. So that will be count A as 
before, left parentheses unique, and the range is B four to B 12. I'll just type that in. Then I have my right parenthesis there. And because I 
have unique nested, within Count A I need to type another right parenthesis, then a comma. And then finally the calculation. And that will just 
be the number of unique values. So U N Q, divided by the total number of items, and that is N U M. And you can see that both U N Q, and N U M 
appeared in the formula, auto complete list. So I'll do that, and then a right parenthesis, and everything looks to be balanced out. So I'll 
press enter, and there you have it. It counts all the unique varieties, divided by the total number of items in the set, and gives me a percent 
unique of 44. And again, if I edit the value in B four, from variety four to variety five and enter, it goes up to 56%.

Refer to Excel tables in a LAMBDA
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] In many worksheets, you will refer to specific cell ranges in your lambda and lead functions. In other worksheets, you will have 
data that is stored in an Excel table. In this movie, I will show you how to use Excel tables as parameters for your lambda and lead functions. 
My sample file is oh two oh five Excel tables and you can find that in the chapter two folder of the exercise files collection. In this 
workbook, I have a single worksheet and I count the number of unique varieties from a list here on the left that's stored in an Excel table and 
then I calculate the percentage of unique varieties. So in other words, I have four separate unique varieties and I have nine items over here 
and that leads to a calculation of 44%. The formula in cell F4 that calculates this result is on the formula bar, and you can see that it looks 
at the range B4 to B12 twice, counting the total number of values and counting the number of unique values. However, let's say that I want to 
add values to the table. That would change the calculation and I want Excel to automatically update the formula. To do that, I can refer to the 
Excel table that stores the varieties that are currently in B4 to B12 with the header in B3. So if I click a cell in the table and then go up to 
table design and look all the way at the left, I see that the table name is Variety List. So that's the table name that I will use for my cell 
reference. I will go back to click F4 and in the formula bar I will edit B4 to B12 to Variety List and then I'll indicate the column by typing a 
left square bracket and Varieties is the name of that column and I'll highlight it and press tab and then a right square bracket to close out 
the reference and I'll do the same thing within the unique function here. So I have B4 to B12, and that was variety list. Left square bracket 
varieties, right square bracket to close it out and enter and we get the same value as a result, but you can see that the references have 
changed in the formula. Now if I add a row to the table, then the calculations will update automatically. So I'll click cell B12 and then press 
tab to add another row to the table and I currently have unique value of 56% and that's because blank is considered a new value. So I'll type 
variety three and enter. So now we've gone down to 40% unique values and that's because I have 10 entries and there are only four unique ones. 
If I edit the value in cell B13 from variety three to variety five and press enter, then we have five unique varieties and the percentage unique 
is 50%. Now one thing you'll note here is that I have my unique varieties and those are calculated using a formula and the formula result spills 
from the original cell of D4 down to, in this case, D8, and it's tempting to try to format this as an Excel table, but in fact you cannot and 
I'll show you what happens when you try. So in D4, I will, on the home tab of the ribbon, go up to format as table, and I'll click a style, 
whereas the data D4 to D8, my table has headers. I'll click okay and formulas are rich data types will be converted to static text. You want to 
continue? No. So in other words, what's happening is that I'm unable to create a table based on this data and have it remain as the result of a 
formula. So as you can see, Excel tables let you name your table and columns within them, which makes your formulas easier to read. If you're 
able to do so, and especially if you're bringing data into your workbook from an external source, use Excel tables in your formulas so they're 
easier for your colleagues to understand and for you to remember.

Create logical branches using IF and IFS
Selecting transcript lines in this section will navigate to timestamp in the video
- [Narrator] Many calculations you create using LAMBDA will require conditional calculations. For example, you might use one calculation if an 
input is less than $500 and another if the input is 500 or more. In this movie, I will show you how to check for conditions using the IF and IFS 
functions. My sample file is 03_01_if and you can find it in the chapter three folder of the exercise files collection. There is a single 
worksheet in this workbook and on it I have purchase amounts and the goal is to assign points based on those purchase amounts and we'll have 
bonuses. We're going to offer 150% points for purchases of more than $500 to start, and we'll change that in a moment. So to create the formula, 
I'll click in cell C4, already selected for me, and I'll just go in and create the LAMBDA directly. So I'll do equal LAMBDA. Our only parameter 
will be the amount which I will abbreviate as amt and now I can create my IF statement. So I use if, and this is just like the function that we 
use normally in workbooks. So that's AMT greater than or equal to 500 than a comma. If that's the case, then the points will be the purchase 
amount multiplied by 1.5. So AMT asterisk 1.5 in a comma, and otherwise it's just amt which is the amount of the purchase right parentheses to 
close out the if, right parentheses to close out the LAMBDA. And to test, we'll use the value from B4. So I'll use that as an input and 
parenthesis sell B4, close it out and enter, and there are the points. And if I drag the formula down I can see that it works properly and that 
I have bonus points which are greater than the purchase amount for 510 and 1400. Let's say that I want to add additional conditions and for that 
I would use the IFS statement. So I will double click in cell C4 and I'll change the function name to ifs. And this allows me to create multiple 
pairs of condition and output. So I have ifs, we'll still use amount, but I'm going to delete everything else I have behind it. I'm going to 
start with the most restrictive condition that will be less than 500, comma and that is just the amount of the purchase, so AMT then a comma. 
Next is if the amount is less than 1000, so AMT less than 1000, comma. If it is, the amount would be multiplied by 1.5, multiplied by 1.5. And 
we're going to add an additional condition where we get double bonus points for any purchase of a thousand or more. So if AMT is greater than or 
equal to 1000 which covers all the other possible values then a comma and we're going to multiply AMT by two. Couple of right parentheses to 
close out and again, I will use cell B4 as the input test, enter right, 350. Now I should see 765 for cell C5 and 2,800 instead of 2,100 for 
cell C6. So I'll go ahead and copy the formula down and those are the results that we expected. Now I can create a LAMBDA based on the formula 
that I just created. So I'll double click C4 and I'll copy everything over within the LAMBDA itself, not including the input. Then control C to 
copy and escape. Then I'll go to formulas and click the name manager, new, and I'll leave the name as points, that's fine. And then for the 
comment I'll say calculate purchase points, including bonuses. And then I will delete the text in refers to and control V to paste what we 
created before. Click OK, everything's good and points is the name of the function. So I'll click close and then I'll delete what's in the 
cells, C4 to C6 and equal points calculate purchase points, including bonuses. Excellent and the amount is before right parenthesis and enter 
350 is good. And we should see the values that we had before. And we do 765 and 2,800.

Select a value using CHOOSE
Selecting transcript lines in this section will navigate to timestamp in the video
- The CHOOSE function lets you assign a value to an input. For example, you could assign the name of Monday to a date with a value of one for 
its weekday. In this movie, I will show you how to use CHOOSE to assign the text name of a weekday to a date. My sample file is 0302 CHOOSE and 
you can find it in the chapter three folder of the exercise files collection. The CHOOSE function selects a value from a list based on an index 
number. We have a seven day week in the US. So, that means that the days are numbered one to seven. So I can work with the dates that I have in 
column B and get the text name of a weekday based on the dates in that column. So I'll start in column C. Type equal and we're creating a 
lambda, so I'll use that function. The input is the date, comma, and then we'll choose from a list that I define based on the index number of 
the weekday of the date in cell B4. I do need to be careful though about what value is returned. So I'm going to make sure the return type is 
correct. I'll type a comma and I see the first option. Number one has number one for the weekday is Sunday and number seven is Saturday. I tend 
to think of Monday as the start of the week. So I will choose option number two. So Monday is number one, and Sunday is number seven. Then a 
right parentheses and that closes out weekday then a comma and now I can add in the text values that we'll choose from and there will be seven 
of them. Again, corresponding to the weekdays. So in double quotes, because it's text Monday, comma, Tuesday, comma, Wednesday. That one's 
always tricky, comma. Then Thursday, comma, Friday, comma. The quotes Saturday and Sunday and double quotes again, and that will close out the 
CHOOSE function. And then right parentheses again to close out the Lambda. And I will use the text or rather the date from cell B4. So I'll type 
that as an input in a separate set of parentheses and enter. And it worked. October the seventh was indeed a Friday. And I can copy that formula 
down. And you see that we have all of the days of the week represented and they all appear to be spelled correctly. Now I can define a function 
using this Lambda that I just created. So I'll double click cell C4 and I'll copy everything except the input at the end. So I've selected it. 
Control C, escape. Then I'll go to the formulas tab, click the name manager, click new, and for the name, I'll call it weekday text. And for the 
comment, return text of a weekday with Monday one and Sunday seven. And then it will refer to the formula that I just copied. So in refers to I 
will replace the existing text with what I copied earlier. Click okay. And just remember it's weekday text, W-K-D-A-Y. Click close. Then I'll 
select cell C4 through C10 and delete the contents, and with them still selected, type equal. And then weekday text, and the date is then B4, 
then right parentheses to close out, and to enter the text into all of these selected cells at the same time, I'll press control enter, and 
there we go. We get the results that we expected and they match what we saw earlier.

Select a value or display a default value using SWITCH
Selecting transcript lines in this section will navigate to timestamp in the video
- The switch function lets you display specific results for specific inputs, but it also lets you display a result for any input that you did 
not list. In this movie, I will show you how to handle inputs that don't match your listed values using switch. My sample file is oh three oh 
three switch, and you can find it in the chapter three folder of the exercise files collection. Over on the left, I have a set of product 
categories and then in cell C4, I have a data validation list that I can use to select from those categories. So, if I click sell C4 and click 
the down arrow you see a list of the values from over on the side and I'll switch to batteries just to show you how it works. If you want to 
create that kind of list, you can click the cell where you want to create it. Go to the data tab, and then in the data tools group click the 
data validation button, select list from the allow box, and then click the collapse dialogue button and select the cells that you want to use as 
the data source. And when you're done, click okay and your list will be in place. My goal with this worksheet is to determine whether a discount 
is available for a particular purchase category. So in cell E4, I will have that indication and I will use the switch function. So I'll type 
equal and I'm going to build it as a lambda directly. So lambda, the parameter will be the category, which I will abbreviate as cat, and then I 
can create my switch function. And now I indicate the data that it's going to receive, and that will be cat, which is the lambda specific 
variable that I created before. Now the first value, and I'll just go in alphabetical order will be batteries and it's tax, so it's in double 
quotes. And I'll say, yes comma, actually, I'll do a quote, 10% and then double quotes, and comma again. And then the next item will be light 
bulbs, and I'll make sure that I have it spelled correctly. I do. If you're unsure of your spelling, you can always copy the data from the data 
validation list if you have one. Then a comma, and the text for this is yes colon, 15%, then a comma, and now I can have the default value, 
which is what will be returned for anything else that's put into cell C4. So I'll put that in double quotes and no discount available, and 
double quotes again. And right parenthesis once to close out switch. Right parenthesis again to close out the lambda. And I will use the value 
from C4. So I have that in parenthesis. Enter and we get a discount of 10%. And then if I change to solar panels, nope. And then light bulbs 
there is. And then power converters, no. So I have all of my cases covered, and if I were to type a value into C4 directly, such as cables and 
enter, you can see that my data validation list doesn't allow that to happen. So I'll go ahead and click cancel, and now I can create a lambda 
based on what I just created. So I will double click cell E4, and copy everything from the equal sign through the rest of the Lambda, not 
including the input. Then control c to copy. Escape. Then I'll go to the formulas tab. Click name manager. I already have my table here called 
product categories. That's the list of categories on the left. I want to create a new name, so I'll click new, and for the name, I'll just type 
discount. And in the comment check if discount available for category. And then I will delete the text in the refers to box and control v to 
paste in the lambda I created earlier. And I'll click okay. So I have discount as my function. I'll click close, and then I will delete the text 
in E4 and type equal discount. And the category will be in C4 so I'll click there, right parentheses and enter. No discount available for Power 
Converters. But, if I go back to C4 and click batteries, then I do get one. If you're trying to decide whether to use the choose function or the 
switch function, remember that the switch function lets you have a value returned in case you don't have a match with your user's input.

Scenario: Calculate Economic Order Quantity
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Placing an order with a supplier and holding the inventory you receive cost money. In this movie, I will show you how to create a 
Lambda to calculate the economic order quantity for your product, which is the number of items that minimizes your total cost of ordering. My 
sample file is 0401EOQ, and you can find it in the chapter four folder of the exercise files collection. This workbook contains a single 
worksheet, and on it I have all the information we need to calculate EOQ. Starting out with Setup Cost. And Setup Cost is the total cost of 
placing an order, that could be any fees for ordering that your supplier charges you, and also the cost of doing business. That could also 
include what you have to pay your employees for their time in creating the order. Flow rate per week is the number of units that you typically 
sell. So in this case, we assume that you would sell 1800 of this particular item per week. And finally, your holding cost. That is the cost of 
keeping an item in inventory for a week. So that is in the same units as the flow rate. And the equation is over on the right. And how this was 
derived is a long story. So I'll just ask you to accept that it is what you need to do to calculate the basic EOQ quantity. So in cell B6, I'll 
create the EOQ formula. I'll type in equal sign, and then we're taking the square root of the quantity that we see on the right, that'll be two 
multiplied by the setup cost, that's in B3, multiplied by the flow rate, that's in B4, and dividing all of that by holding cost, and that is in 
B5. And I don't need to use any extra parenthesis because multiplication and division are on the same level within the order of operations for 
Excel. So I can do it this way. If I had a plus in there, then I might need to add parenthesis, but in this case I don't. So I'll type a right 
parenthesis here to close out the arguments for square root and Enter. And I get an EOQ of 2939 and a bit, and we'd probably round that up to 
2940. Now with that in place, I can create a Lambda based on that information. So I will click B6 and then I will open it for editing. And then 
to the left of square root and to the right of the equal sign, I'll type Lambda. And we're going to be using three different inputs or 
parameters. The first one is setup, so that's a setup cost, then a comma. The second is the flow rate. So I'll do flow underscore rate, then a 
comma. And finally the holding cost, and I'll just call that hold, then a comma. Now I have my calculation but I need to replace the cell 
addresses with the parameters. So two remains the same. B3 is the setup, and you can see that it appears in the formula auto-complete list. And 
then B4 is the flow rate. And then B5 is the holding cost or hold. Now I will add a right parenthesis to the end to close out the Lambda, and to 
test I'm going to use B3, B4, and B5 as inputs. So I need to put them in a separate set of parentheses. So that's B3, B4, B5. And right 
parentheses and Enter. And we get the same answer that we did before. So the Lambda is working. Now I can create a function in this workbook 
using the name manager. So I will double click cell B6, and then copy everything for the Lambda itself, but not the arguments to the right. Then 
Control+C to copy and Escape. Now I'll go to the formulas tab on the ribbon, click Name Manager, click New and the name is EOQ, that's already 
there so I'll just leave it. For the Comment; Calculate Economic Order Quantity, and then in Refers to, I will replace the existing text with 
the text that I copied. All right, everything is good, so I'll click OK. And I have EOQ, it's indicated as a Lambda. So I'll go ahead and click 
Close. Now, I'll go back and edit the value in cell B6 to equal EOQ, there's the name of the function, left parenthesis, the setup is B3, flow 
rate is B4, holding cost is B5 and Enter. And we get the value that we did before. So now you can copy this EOQ custom function to other 
workbooks and use them there as well.

Scenario: Calculate quality of service
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Most businesses set goals for the quality of service they offer to their customers. One scenario would be a repair company that 
sends parts to wind farms in the American Midwest. In this movie, I'll show you how to create a LAMBDA function to calculate the quality of 
service given customer demand and warehouse location. My sample file is 04_02_QualityOfService, and you can find it in the Chapter 04 folder of 
the Exercise Files collection. Rather than start with the description of the workbook's contents, I'm going to talk about the goal, and that 
goal is found here at the bottom right. Let's suppose that for our quality of service, we want to ship parts no more than 250 miles to a 
particular customer. Now, that's not going to be possible. Some parts will almost certainly have to go more than 250 miles. So we have set a 
goal of sending no more than 70% of all parts more than 250 miles. And you can see that given the current configuration, we have a quality of 
service level of 70.02%. So we're meeting our goal, and these numbers were generated with the idea of minimizing cost. So how do we get to those 
calculations? Well, at the top left, I have a table that shows the distance between our customers, which are here in column B, and the 
warehouses, which are here in row three. And the distance is in miles. And below that, we have a matching array with zeros and ones. Any value 
that is greater than 270 gets a zero, and any value that is less than 270 gets a one. So you can see for Amarillo that the first item or 
distance is 270. So that gets a zero because it's more than 250. And then the second is 210. So that gets a one. And the same throughout the 
rest of the array. On the right we have units shipped. So in this case, we're sending 72 from Amarillo to Abilene. And then if you look down, to 
save costs and meet our goal, we had to send three from Amarillo to Dodge City and another 397 from Kansas City. And we're doing that even 
though Amarillo is about 90 miles closer to Dodge City than Kansas City. So you always have to make some trade-offs. And at the bottom, we 
calculate the quality of service using a SUMPRODUCT formula. SUMPRODUCT multiplies two arrays of the same size element by element and then adds 
up the total. So here, we have 72 times zero and then 528 times zero, which means that none of these units count as being below 250 miles. 
However, this 325 from Amarillo to Lawton does because 325 is multiplied by one instead of zero. Then, the rest of the formula gets the sum of 
all the units shipped, and that's the array here from H2 to J11 and divides that to find the percentage. So our goal is to transfer our current 
formula to a LAMBDA. So I'll start by double-clicking in cell H18, and the function will be a LAMBDA. And we'll have two parameters, shipped and 
dist, D-I-S-T. So we have the number shipped and then the distance. For the SUMPRODUCT, instead of our ranges, I'm going to use the parameters, 
so the SUMPRODUCT of shipped multiplied element-wise and added with D-I-S-T. And then we'll have the SUM of shipped. I will close out this part 
of the formula with a right parenthesis. So we have our LAMBDA. And now I need to tell it which items to use, and those'll be the ranges. So 
we're using this as an input to make sure that our LAMBDA's working properly. So then we will have H4 to J11. That is the number shipped, then a 
comma. And then the distance, whether it's below the maximum or not, is C16 to E23. There we go. Right parenthesis and Enter, and we get the 
same value we did before. And now we can turn it into a named function. So I will double-click cell H18 and copy everything through the LAMBDA. 
But I'm not going to include the parameters because we won't need them. Then, Control + C and exit. Now, I'll go up to the ribbon and click 
Formulas and Name Manager. Click New, and the name will be QOS_Shipping. And I won't put in a comment this time because this movie's running a 
little bit long. And then press Control + V to copy the LAMBDA function that I created earlier, and click OK. Right, that looks good. I'll click 
Close. And now I'm going to replace the LAMBDA that I have in H18 with a function that uses the ranges that I described earlier. So I will have 
QOS_Shipping, and then we just need the range for the units shipped. And that again is H4 to J11, then a comma, and the distances and whether 
they are 250 or lower. So that's C16 to E23, right parenthesis, and Enter. And we get our result.

Scenario: Calculate quality of service
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Most businesses set goals for the quality of service they offer to their customers. One scenario would be a repair company that 
sends parts to wind farms in the American Midwest. In this movie, I'll show you how to create a LAMBDA function to calculate the quality of 
service given customer demand and warehouse location. My sample file is 04_02_QualityOfService, and you can find it in the Chapter 04 folder of 
the Exercise Files collection. Rather than start with the description of the workbook's contents, I'm going to talk about the goal, and that 
goal is found here at the bottom right. Let's suppose that for our quality of service, we want to ship parts no more than 250 miles to a 
particular customer. Now, that's not going to be possible. Some parts will almost certainly have to go more than 250 miles. So we have set a 
goal of sending no more than 70% of all parts more than 250 miles. And you can see that given the current configuration, we have a quality of 
service level of 70.02%. So we're meeting our goal, and these numbers were generated with the idea of minimizing cost. So how do we get to those 
calculations? Well, at the top left, I have a table that shows the distance between our customers, which are here in column B, and the 
warehouses, which are here in row three. And the distance is in miles. And below that, we have a matching array with zeros and ones. Any value 
that is greater than 270 gets a zero, and any value that is less than 270 gets a one. So you can see for Amarillo that the first item or 
distance is 270. So that gets a zero because it's more than 250. And then the second is 210. So that gets a one. And the same throughout the 
rest of the array. On the right we have units shipped. So in this case, we're sending 72 from Amarillo to Abilene. And then if you look down, to 
save costs and meet our goal, we had to send three from Amarillo to Dodge City and another 397 from Kansas City. And we're doing that even 
though Amarillo is about 90 miles closer to Dodge City than Kansas City. So you always have to make some trade-offs. And at the bottom, we 
calculate the quality of service using a SUMPRODUCT formula. SUMPRODUCT multiplies two arrays of the same size element by element and then adds 
up the total. So here, we have 72 times zero and then 528 times zero, which means that none of these units count as being below 250 miles. 
However, this 325 from Amarillo to Lawton does because 325 is multiplied by one instead of zero. Then, the rest of the formula gets the sum of 
all the units shipped, and that's the array here from H2 to J11 and divides that to find the percentage. So our goal is to transfer our current 
formula to a LAMBDA. So I'll start by double-clicking in cell H18, and the function will be a LAMBDA. And we'll have two parameters, shipped and 
dist, D-I-S-T. So we have the number shipped and then the distance. For the SUMPRODUCT, instead of our ranges, I'm going to use the parameters, 
so the SUMPRODUCT of shipped multiplied element-wise and added with D-I-S-T. And then we'll have the SUM of shipped. I will close out this part 
of the formula with a right parenthesis. So we have our LAMBDA. And now I need to tell it which items to use, and those'll be the ranges. So 
we're using this as an input to make sure that our LAMBDA's working properly. So then we will have H4 to J11. That is the number shipped, then a 
comma. And then the distance, whether it's below the maximum or not, is C16 to E23. There we go. Right parenthesis and Enter, and we get the 
same value we did before. And now we can turn it into a named function. So I will double-click cell H18 and copy everything through the LAMBDA. 
But I'm not going to include the parameters because we won't need them. Then, Control + C and exit. Now, I'll go up to the ribbon and click 
Formulas and Name Manager. Click New, and the name will be QOS_Shipping. And I won't put in a comment this time because this movie's running a 
little bit long. And then press Control + V to copy the LAMBDA function that I created earlier, and click OK. Right, that looks good. I'll click 
Close. And now I'm going to replace the LAMBDA that I have in H18 with a function that uses the ranges that I described earlier. So I will have 
QOS_Shipping, and then we just need the range for the units shipped. And that again is H4 to J11, then a comma, and the distances and whether 
they are 250 or lower. So that's C16 to E23, right parenthesis, and Enter. And we get our result.
Scenario: Calculate process capacity given a batch size
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Many companies create multiple versions of the same product, perhaps changing the flavor of ice cream, or making different colors 
of the same product. Because they need to change from one product type to another, companies will divide production into batches. In this movie, 
I'll describe how to create a lambda function to calculate the capacity of a process given various constraints. My sample file is 0403 Batch 
Size, and you can find it in the chapter four folder of the exercise files collection. This calculation uses three inputs. So I have my batch 
size, which is 2,400, so that's the number of items that will be created in one run of this process. The setup time in hours is 0.4, so that is 
0.4 times 60 minutes, so that's 24 minutes. And the processing time per unit, which is also in hours, is 0.05, and that is 0.05 times 60, or 
three minutes, 1/20th of an hour. Now we can create the formula that I have written over to the right, and that's the batch size divided by the 
setup time, plus the batch size, times time per unit. So we count the setup time once, and then we add the total of the time that's needed to 
create all the items. And then we divide the batch size by that quantity. So that gives us our capacity per hour. So in B6, I'll type equal, the 
batch size is in B3. Then we'll divide that by the quantity of the setup time, which we only count once, that's in B4. But then we will add the 
batch size by the time per unit to get the total time required. So there I have a right parenthesis, everything is good. Enter. And we get a 
capacity per hour of 19.9. So just below 20. One way to look at this calculation is that we can make 20 units per hour. So that is the upper 
limit for our capacity if there is zero setup time. However, the setup time takes 24 minutes, or four tenths of an hour, and that reduces our 
capacity. Because we're going to be running for a while, it means that the setup cost is low in comparison to the overall time required. Now say 
that I want to create this calculation as a lambda that we can use here and in other workbooks. So I will double click cell B6, and then I'll 
click, to the right of the equal sign and type lambda. We'll use three parameters. That will be batch, comma, and then set up, comma, and then 
P-R-O-C for process. And our calculation, and I'll just delete everything I have here. And that will be batch, which again is our batch size. 
And that will be divided by the quantity of the setup time plus the batch size, times processing time per unit. And that is PROC. Then I will 
add a right parenthesis to the outside to close out the lambda, and I'll provide three inputs in parentheses. Those will be B3 for the batch 
size, B4 for the setup time, and B5 for processing time per unit. Right parentheses and enter. And we get the same value as before. And now I 
can create a lambda function that we can use in this workbook and copy elsewhere. So I will double click B6, and then copy everything to the 
last parentheses of the lambda. Then CTRL+C, then escape. Then on the formulas tab of the ribbon, I'll click name manager, click new, and the 
name will be batch capacity. And then I'll skip adding a comment for now. And then I will paste the text that I just copied from the formula 
into the refers to box, replacing what's already there. Now I'll click okay, batch capacity is there, so I'll click close. And then in cell B6 
I'll test it by typing equal batch capacity. And then we have B3, comma, B4, comma, B5, right parentheses and enter. And we get the value that 
we did before.

Scenario: Clean up imported text
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] When you bring in data from an outside source, such as a scanned PDF file, you might end up with characters you don't want and 
spaces everywhere. In this movie, I will show you how to build a lambda function to clean up that text. My sample file is 04_04_Text, and you 
can find it in the chapter four folder of the exercise files collection. In column B, I have a set of company names that I brought in from an 
outside source. And you can see that I have extra spaces at the beginning and also in the middle of some of these words, also that everything is 
lowercase. And in a few cases, I have added spaces at the end, just to make it to be a little bit of a more realistic exercise. I decided not to 
include any non-printing characters. And the reason I did that is because you wouldn't be able to see them, but you could see that there were 
additional spaces at the end of the value in B3. Because company names typically have capital letters at the start of each of their words, or at 
least most of them, you can correct the text in three different ways. So the first will be to clean all the text, that is, to take out the 
non-printing characters, then to trim the text, which removes all white space, spaces, tabs, and so on, except for one space between each of the 
words, and then also proper, which makes the first letter of each word capital. So I'll show you how I do that in cell C3. I need to start from 
the outside in, so I'll type =. The last step that I want in my formula is to change the remaining text so that it's initial caps, so that will 
be proper. And then the text will actually be coming from two other functions. The first one is trim, and trim is the function that removes 
extra spaces from your data. And then, finally, we'll have clean, and clean removes any non-printing characters from your data. And the source 
for this will be cell B3. And then 1, 2, 3 right parentheses to close out each of my nested functions and Enter. And I get Bartley Brothers, and 
it looks correct. Now I'll copy that formula down. So I'll click cell C3 and then double-click the fill handle at the bottom right corner. When 
I move my mouse pointer over it, I know I'm in the right place when the pointer changes to a black crosshair. So I'll double-click, and it goes. 
Now I can look at my data, and everything looks okay except for Addison and Clark. And I see here that the word and also has a capital A. And 
that's not surprising because I used proper, and that means that every word will have an initial capital letter. So that means I need to change 
my rule a little bit to account for this error. So I will double-click cell C3, and I'll put another function on the outside, and this will be 
substitute. And this substitutes one string of text for another. So the first argument is the text, and that is everything here, proper, trim, 
and clean. Then I'll click after the trailing right parentheses, then a comma. The old text will be And with a capital A, and it is 
case-sensitive, so that's in quotes, and then a comma, and then double quote and. So that will be a lowercase and for Addison and Clark. And if 
you're reading ahead, you know that something interesting is about to happen. So I'll type a right parentheses and Enter, and then I'm going to 
double-click again to copy the formula down. Addison and Clark is correct, but now Anderson Industries is wrong. And you can see that there's a 
lowercase a at the start of Anderson. So that means we need to refine our rule a little bit. So I'll go up to C3. And then instead of just 
having and, uppercase A plus and lowercase a, I'm going to add spaces. So space And, and then space and space. Great. So now we're only going to 
be changing and if it is a word that occurs on its own, with a space on either side. And remember, we will have exactly one space between all of 
our words because we use trim to get rid of the extra white space, so then Enter. Bartley Brothers, no change. It's not affected by the rule. 
And now we're going to watch the values in C5 and C7 to see how they do. So double-click the fill handle again. Addison and Clark is correct, 
and Anderson Industries is correct. Now I can create a lambda that will clean up this particular version of the text. So I will double-click 
cell C3. I will click to the right of the equal sign and type lambda. And then we're going to use our text as an input. So I'll type text and 
then a comma, and then I'm going to replace the cell reference of B3 with text. And then I'm going to go to the end of the lambda formula, type 
another right parentheses to close it out, and then to test, I will give it the input cell of B3. So left parentheses, our input is B3, and 
Enter. Looks good. And then I will click cell C3 and double-click the fill handle, copy down. Everything looks great. So now we can go to the 
Formulas tab and use the Name Manager to create our function. So I'll double-click in cell C3 and copy the lambda, but not the input, Control + 
C and Exit or Escape, then Formulas, Name Manager, click New. And then the name will be CompName 'cause I know it's always going to be used to 
clean up company names. I won't add a comment now, to save time. And then I will select and replace the text in Refers to by pasting in using 
Control + V, there's my formula. I'll click OK. Everything looks good for the lambda. I'll click Close. And then in cell C3, which I already 
have selected, I'll type =CompName, so that's company name. The text is B3, right parentheses and Enter. Looks good. So I'll click C3, 
double-click, and the values appear exactly as they should. One reason I wanted to give you a text handling example is because when you bring in 
data from outside sources, it can be very hard to work with it. In most cases, you'll have to do some changes by hand, but if you're creative 
with your formulas, and you watch out for accidental replacements, you can go a long way automatically.

Update values using MAP
Selecting transcript lines in this section will navigate to timestamp in the video
- [Presenter] Some data analysis tasks require you to perform a mathematical operation on a set of numbers and write the results to another cell 
range. You can use lambda to define that operation and then apply it using the map function. In this movie, I will show you how to do that in a 
scenario analysis worksheet. My sample files 0501 map and you can find it in the chapter five folder of the exercise files collection. As I 
mentioned, this workbook contains a scenario analysis worksheet and you can see at the top that I have three cases, best, middle and worst. And 
then in cell B3 I have a dropdown that I created using a data validation list. So I can select whether I want to display the best case, the 
middle case or the worst case and I'll press escape to get out of that display. I have adjustments for each one so if I were to select best then 
the numbers in my base scenario which are in row eight and nine would be multiplied by 125%. If I go to worst case, they'd be multiplied by 75%. 
And then the middle case, of course is the numbers themselves multiplied by a 100%. Let's say that I want to use the map function in combination 
with a lambda to display the adjusted values in row 13. So below the second set of headings and below the adjusted values top level heading. I 
can do that by going to cell B13 which I've already selected, typing equal. And then the formula I'll use is lambda. I need to enter my 
parameter or calculation in this case it will be two parameters. So I'll have the range which is the range of values, then a comma then another 
parameter that will be the factor. And the factor will be the value in B4. That's the adjustment. And now I can enter the calculation, which 
will be range. And you can see it in the auto complete list because I've defined it as an argument within the lambda. And multiply that by 
factor factor. There we go. And that's correct. So I will type A right parenthesis, but now I want to give some input values so that Excel is 
able to test it. So I'll type A left parenthesis, and we'll use the range of B9 to E9 that contains the original base scenario values, then a 
comma, and we're going to multiply them by the value in B4. All right, everything is good. I'll press enter and I get the values. I'm currently 
in the middle case at 100%, so I'll go up to cell B3 and I'll change it to the best case. And you can see that everything is multiplied by 125%. 
If I go down to cell B13 and click it, you can see the formula as I entered it. And if I go to C13, the formula has spilled from cell B13 over 
to B14. One thing to notice is that the formula on the formula bar is displayed in gray so you can't actually edit it here. You have to go back 
to the cell in which you entered the original lambda formula and that is cell B13. So our lambda function is working correctly. And now we can 
define a custom function using the name manager. So with cell B13 still selected, I will go up to the formula bar and I will copy everything 
from equal lambda to the right parenthesis after factor. So I'm leaving out the arguments, so I have that selected. I'll press control C, then 
escape to exit the cell. Then I'll go to the formulas tab on the ribbon click name manager, and then I will click new. And then in the new name 
dialogue box I will change the name to SCENAdj for scenario adjustment. And then for the comment it will be adjust scenario values by a factor. 
And then in the refers to box, I will select the existing text and press control V and paste in the lambda. So then I'll click okay and click 
close. And our lambda has been created. Now I will select cells B13 through E13 and delete the contents. Then click cell B13 again and then I'll 
type equal and then scenario adjustment. So this is the function that we just created. The range will be B9 through E9, then a comma and the 
factor is in cell B4. So I'll click there, write parenthesis, and enter. And you can see that the formula spills over and we have the values 
that we have before but instead of this long lambda that we have to edit every time we have a simple compact formula and just to make sure 
everything works I'll go back up to cell B3 and select worst everything changed, go back to middle. And we have the values we expect with our 
lambda function working as expected.

Summarize values using REDUCE
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] One common data analysis task is to summarize a set of values using a specific calculation. For example, you could multiply every 
value in an array by 1/2, or 0.5, and add up the results of the individual operations. One way to apply more complicated calculations is to use 
the reduce function in combination with a lambda calculation, and in this movie I will show you how to do it. My sample file is 05_02_Reduce and 
you can find it in the chapter five folder of the exercise files collection. In this workbook, I have a worksheet that has a number of routes, 
eight of them, in fact, and then the distance for each of those routes. So all that is obvious enough. In column C, I have the adjusted 
distance. And the reason I have that column is because I want to assume this business pays any driver who goes 100 miles or more in a day an 
additional 50%. So if you offer them a bonus by mile and they go over 100, then you would multiply the distance traveled by 1.5 to get the 
adjusted distance. First, I will calculate the values individually and then show you how to do it in a single calculation using your reduce. So 
in cell C4, which I've already selected, I'll type equal and I'll use an if function. So if and then a left parenthesis. And the logical test is 
whether the value in B4, that is the distance of this route, is greater than or equal to 100, then a comma. If it is, multiply B4 by 1.5, and if 
not, we'll just return B4. So everything's good. I've got a right parenthesis, enter, and we get 156. I'm going to copy that formula down. So 
I'll click cell C4, and then at the bottom right corner of the cell you can see a green square. That's the fill handle. So I'll move my mouse 
pointer over top of it and when it changes to a black cross hair, I will double click, and there we go. So any value that is 100 or more is 
multiplied by 1.5. Anything that is 100 or less is just the original value. And I have a formula that gets the sum of all those individual 
values here. But now let's suppose that you want to do all of this in one formula in one cell instead of adding up all of the adjusted distances 
in a separate column. To do that, you can use the reduce function. So I will go to cell C14 and then equal, and then I'll use reduce. We start 
by adding the initial value, and you can leave this blank. If you do, Excel will assume it's zero, but I like to put it in just so the formula 
is easier to read. So zero, and then a comma. The array is B4 to B 11, and, remember, we are using the distance and not adjusted distance. If 
you multiply the values in adjusted distance by 1.5, then you would be multiplying any value of 100 or more twice. So we have B4 to B 1, and 
then a comma. And now we can enter the lambda function. So I will do lambda, and then we'll start with the accumulator value, which will be A, 
then a comma. Then we'll use the distance, which I will call DIST, and this is a variable, as the same as A that are internal to this lambda, 
then a comma, and now we can create the same "if" function or formula that we did before, except we'll use the values A and DIST. So if, left 
parenthesis, logical test is if the distance, DIST, is greater than or equal to 100, then a comma, then we will have the accumulator value, or 
A, plus DIST times 1.5, then a comma. And if not, it'll just be A, the accumulated value, plus the distance, DIST. And then I need one, two, 
three right parentheses to close everything out. I'll press enter and we get the same value that we did before of 862.5, except that now it is 
all contained within one cell. And if I wanted to, I could take the lambda that I used here and go through the name manager to create a named 
function that I could use elsewhere.

Calculate intermediate values using SCAN
Selecting transcript lines in this section will navigate to timestamp in the video
- [Narrator] The SCAN function, which I'll describe in this movie applies a LAMBDA to a data set and calculates intermediate values as it works 
through that set. One common example for this type of calculation is a running total which I will demonstrate in this movie. My sample file is 
05_03_scan, and you can find it in the chapter five folder of the exercise files collection. In this workbook, I have a single worksheet and I 
have a table of eight routes that could be driven by employees of a particular company and then the distance in column B for each of those 
routes. And the idea is that this company offers a bonus to its drivers for the number of miles driven in a day. If the number of miles is 100 
or greater then it multiplies the total by 1.5. If not, in other words, if it's below 100 then they just pay the driver for the actual miles 
driven. I will use the SCAN function to create formulas to calculate both the adjusted distance and an adjusted running total. So I'll start 
with just the adjusted distance. I'll go to cell C4 and then type in equal sign and I will use the SCAN function. The initial value for the 
accumulator is zero so we'll start there, then a comma, the array of values that we're going to summarize is in cells B4 through B11, then a 
comma and the function we'll use is a LAMBDA. So I'll go with the LAMBDA and we'll use two parameters. The first will be the accumulator value 
which I'll just call A then a comma, And then the distance, which I will call dist and A and dist are variables that are used only within this 
LAMBDA function. So they're not used anywhere else in the workbook. So then a comma, and now we can put in the calculation. So for this we'll 
use an IF function. So if the distance, dist is greater than or equal to 100 then a comma, we'll multiply the distance by 1.5 and a comma, and 
if it's below 100, we'll just return the original value. And I need 1, 2, 3, right parentheses, and then I'll press enter. And there we get our 
values. And you can see that the formula spilled to include every cell in a row that is next to the cells in the original range of B4 through 
B11. And we get our adjusted distance total at the bottom of 862.5. Now let's say that I want to calculate an adjusted running total. So instead 
of having the value for each individual route I want to have the accumulated total as we move down through the routes. So I'll go to cell D4 and 
this function will be very similar to the previous one. So I'll move through it a little more quickly, only explaining the differences. So in 
D4, I'll type equal, and again we'll use SCAN initial value zero range B4 to B11, then a comma, and we're doing a LAMBDA again. And the first 
parameter will be A for the accumulator then a comma and distance. And you can see that dist doesn't show up in the auto complete list. That's 
because, again it's not something that applies to the entire workbook it's just within this function, then a comma, and now the calculation. 
Again it will be an if then left parenthesis if the distance is greater than or equal to 100, comma, if it's true, then we will multiply the 
distance by 1.5 and add that to the accumulator value, then a comma and if it's not true then we will just have the distance, unadjusted added 
to the accumulator. So everything is good here. So 1, 2, 3, right parenthesis to close out and enter. And there we see we get our adjusted 
running total. And I have 156 plus 53 is 209, that's correct. Plus 195 is 404 also correct. And at the bottom the last cell gives me an adjusted 
running total of the entire set for 862.5. And that means that I don't have to add up all the values because the formula that I created has done 
that for me already. Generate an array of values using MAKEARRAY
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Data analysis tasks often use array of values to generate results. In this movie, I will show you how to use the MAKEARRAY 
function in combination with a LAMBDA to generate a random set of values. My sample file is o5_o4_MakeArray, and you can find it in the chapter 
five folder of the exercise files collection. Let's say that I have a goal to plant crops randomly in a field, and in rows three and four, you 
can see that I have my number of rows in B3 and the number of columns in B4. And my goal is to plant four varieties randomly, allowing repeats 
within a three by three grid. And I can do that using the MAKEARRAY function. So I'll click on cell D3, type an equal sign and we'll use 
MAKEARRAY. The number of rows is the value in B3. The number of columns is in B4 and know they do not have to be the same number. Now I can 
create my function as a LAMBDA, so I'll do LAMBDA and the first parameter will be the number of rows. So that will be called R, then a comma C 
for columns, then a comma. And now I want to choose, let's say from among four varieties and I'll just call them varieties one through four. I 
will use CHOOSE, to select a variety at random. The index number will be generated randomly, so I'll use RANDBETWEEN, and we have four different 
varieties, so I'll type one as the bottom number than a four as the top number. Close that with the right parenthesis, and then a comma and now 
a list of the varieties. Those will be variety one through variety four and I'll type them in with the names between double quotes. So double 
quote variety one, and I'm separating each value by commas. Then double quote variety two, and a comma variety three as before, and comma then 
variety four and a double quote. And I have to use double quotes, because the values are text and I have spaces, so that indicates that I have 
text, right? Everything's good. So I will type one to three right parenthesis to close out and enter. And there I have my list of varieties. If 
I want to recalculate it, I can press F9. And that gives me different layout, F9 again. And you can see that some values are changing and some 
values aren't. So, I don't have these values changing in the background. While I do my next task, I will select cells D3 through F5, press 
control C, and then I will paste in the values using a set of keyboard shortcuts. So, I'll press alt that opens up the key tabs, then H for the 
home tab, V four paste and then V again to paste values. So the formulas are gone in cells D three through F5 and I just have the values, so 
they'll stay the same. Another task that I wanted to show you that isn't actually related to LAMBDA but is very useful and fits in here. So, I 
thought I would go ahead and do it. And the idea is to generate a random ordering for the values in cells A8 through A16. And I can do that by 
creating a set of random decimal values and then sorting based on them. So, I'll click on cell B8 and actually drag down to include the range 
from B8 to B16. Then in B eight we'll type equal. And then the function is RAND, R A N D. And it doesn't take any arguments, so I'll just type 
an open and close parenthesis. And then to enter the same formula in every selected cell, I'll press control enter. And there I have a set of 
random decimal values. And if I increase the size of column B, you can see that there are quite a few decimal values displayed. And if I go to 
the home tab and the number group and click increased decimal, you can see that there are a lack actually a lot of decimal places. And see if I 
increase. Now I believe I finally got to the point where I get repeating zeros. Yes, I have. So if I decrease decimal, and there we go. So that 
is how many digits are randomly generated and you will probably, and in fact I'm willing to guarantee that you will never see two values the 
same consecutively or even within the same data set. So I want to sort based on these values and I will copy them. So I'll press control C and 
paste them in. So alt H, V, V as I did before, and those key strokes are independent, so you don't hold down alt and then press H. It's alt 
release H, release V, V. So there are my values, and now I can select the range from B8 through A16. Then again, on the home tab, I'll go to 
sort and filter and I will sort smallest to largest. So, I have sorted based on the values that were in the first column I selected. That's why 
I selected column B values first. And so I have all those values and you can see that they go in ascending order. And over here I have a random 
set of varieties from the values that I created before. So, if I want to create a randomly ordered set of values, all I need to do is type them 
in, generate random decimal numbers and then sort based on the random values.

Apply a LAMBDA to an array by column using BYCOL
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Lambda functions can be applied to entire arrays of data or just parts of them. In this movie, I will show you how to apply a 
Lambda to individual columns of an array. My sample file is 0505 by column, and you can find it in the chapter five folder of the exercise files 
collection. This worksheet summarizes the number of packages handled by a company and it's broken down both by day. So you see, I have the days 
of the week, Monday through Friday, and also by route. I also assume that I pay my drivers and employees based on the number of packages that 
they handle. And if they handle 500 or more on a particular day then I want to give everyone a bonus per package of 110%. So, instead of paying 
them say a dollar per package, I would pay them $1.10. And I can calculate that using a Lambda. So, I'll start with just creating a Lambda that 
will perform the calculation for a single column. So, I'll click in cell B10, type equal and the Lambda that I create will use the column of 
data. So, I'll just type in column as my argument and a comma. And then if the sum of column, which is the data that's been entered is greater 
than or equal to 500 comma, then the value if true will be sum of column times 1.1. And if not, it'll just be the total. So comma sum of column. 
And three right parentheses to close out. I'm using the colors to know when I get to black which indicates that I've reached balance in my 
equation with parentheses and enter and I get calc, and that's because I forgot to tell Excel and the Lambda which range of cells to use. So, I 
will double click cell B10, go to the left of the formula, and then left parenthesis, and then we're using B5 through B9. So, that's the cell 
range there and enter. And now I get my value. And I can copy this formula to the right. So, I will click cell B10 and then drag its fill handle 
to the right and it's been applied. So, any value below 500 remains the same. Any value that started as greater than 500 is now multiplied by 
1.1. So, I've shown that the Lambda works, but now I'm going to change it so that I work on the entire array of data and go by column. So, I 
will erase the formulas in cells C10 through F10 so I've selected them. Just press delete and then I will double click cell B10 and start 
editing. Here I want to go by column. So, I will put the BYCOL function to the left of Lambda and the range that I want to use will be B5 
through F9. So, those are all the values in the array. Then I'll type a comma, all right, everything has been selected. So, I have B5 through B9 
and now I want to delete the range at the end so I'm not operating on B5 through B9 nine anymore. And also, I can tell by the colors of the 
parenthesis, the right parenthesis here at the end, that I need a black parenthesis. So, if you see here at the start on the left I have a black 
left parentheses and then the second one is red, and then you work through the colors and then they balance out. So green and green, green and 
green again, green and green. And then we go to the outer areas where we have purple, red, and that means I need black, so everything is 
balanced. I'll press enter and I get the same formulas that I did before except now they've been entered with only one formula entry into cell 
B10. One other thing that I can do to clean up the data is to round up to the next highest integer value and that uses the round up function. 
So, I'll double click B10 again and then to the left of BYCOL and to the right of the equal sign I'll type roundup, left parenthesis, and then 
we're rounding up the value generated by the formula here. And then we need to tell it the number of digits to the right or left of the decimal 
point. So, I want it to be a whole number and that will be zero digits and then right parentheses to close out and balance out the parenthesis 
and enter. And there we get our whole numbers where any fraction or decimal value has been rounded up to the next highest number.

Apply a LAMBDA to an array by row using BYROW
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Lambda functions can be applied to entire arrays of data or just parts of them. In this movie, I will show you how to apply a 
Lambda to individual rows of an array. My sample file is 05 06 by Row, and you can find it in the chapter five folder of the exercise files 
collection. I'm using the same data that I used in the previous movie, but this time I want to focus on the rows, which calculate or summarize 
the number of packages that have been delivered on a route for a given week. So I have route one in row five, route two in row six and so on. 
I'm assuming that I am paying my employees per package and I want to offer them a bonus of 10% if they handle more than 500 packages on a route 
within a week. And that will be the calculation that we do in column G for adjusted route total. So in cell G5, I'll type in an equal sign, and 
we're creating a Lambda function. The parameter will be the values in a row, so I'll just use row as that name. And if the sum of the values in 
the row are greater than or equal to 500, we'll multiply that sum by 1.1. So I'll have sum, row times 1.1, and if not we'll just return the same 
value. So that'll be sum of row. Now I need to close out my parentheses. So there's one, I can see that it's red, and I need another for black. 
And you can see that the first left parentheses is black in color, and the second is red, and then it comes out at the other end. So we have 
purple and then red and then black. That's how you know you're balanced out. The values were summarizing are in cells B5 through F5, so I'll 
input those as arguments. So another left parentheses, and then B5 through F5. And right parentheses looks good and enter. And there we get a 
total of 456, and it's under 500, so it wasn't multiplied. And then I'll just go ahead, and copy this formula down. And I can see the three of 
the five routes handled more than 500 packages in this week. If I don't want to copy the formula, then I can use the by row function. So I will 
delete the formulas in G6 through G nine, and then double-click cell G5 to edit the formula. And here I will use the by row function. So I'll 
click just to the right of the equal sign and type by a row. Then a left parenthesis. The array that we're going to use is B5 through F9. So all 
the data then a comma, and I don't need B5 through F5 again. So I'll go ahead and delete that. And I see that my parenthesis are out of balance 
so I'll type another right parenthesis. Looks good and enter. And there we get the same results as before, but only entering a formula in one 
cell. One last thing to do would be to round up the values and we'll give our employees the benefit of rounding up to the next highest number, 
even if the decimal value is less than 0.5. So I'll double-click G5 again, and we will use the round up function. So round up, left parentheses, 
and then I can keep the rest of this in place because that's the result or number that we're rounding. Then click to the right of what is now 
the final red parentheses, then a comma. And we need to indicate the number of digits to the left or right of the decimal point. I want to use 
whole numbers, so that will be zero. So we have no digits to the right of the decimal point, right parenthesis and enter. And we get our results 
rounded up to the next whole number.

Manage LAMBDA output
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Lambda functions can return more than one value. They can be in rows, in columns, or in arrays. In this movie, I will show you 
four functions that you can use to manage output from Lambdas that return multiple values. My sample file is 05_07 manage output and you can 
find it in the chapter five folder of the exercise files collection. I have two sets of values that I work with in this movie in the range B3 to 
D4. I have a number of randomly selected varieties and then I will perform other operations wrapping values into arrays using columns or rows in 
the table here from A7 through A13. But I'll start working with the values here in my upper range. My goal is to output this data as a single 
row. And what I want Excel to do is to take the first row of data, so that would be row three and then to the right of it, use the data from row 
four of the worksheet. So I'll click and cell G2, type equal. And the function I'll use is in cell F2 that is two row and the array is B3 
through D4. And I don't need to set any of my other parameters here. So I'll type right parenthesis and enter and I get the values. So you can 
see that it takes the first row and then immediately to the right of that, it writes the second row. In the same way, you can write an array 
into a single column. So I'll click and cell G4, type equal. And this time, as you might have guessed, the function is TOCOL or to column and my 
array is the same as before, B3 to D4. Don't need to change anything else. So I'll type right parenthesis and enter. And there I have variety 
four, variety of six and variety one, and then variety three, variety two and variety three. Notice that it did not go by columns. Instead it 
went by rows. So it took row one, rendering it as a column and then did the same for row two. You can also go in the other direction. So if your 
Lambda returns a single row or a column of data, you can wrap it into an array. So let's work with the data that I have in A8 through A13. So, 
and cell C8, I will start typing my formula. So equal, and the function name is wrap columns. The vector and vector is a fancy word for a single 
column or row will be the values in this table here. So I'm going to move my mouse pointer over the center of the header of my table. And when 
it's a black downward pointing arrow, I will click. That selects the entire table column and that's variety list varieties, then a comma and my 
wrap count will be the number of rows that I want. So I will wrap it into three rows. I don't need any padding, so type right parentheses and 
enter. And I get three rows and two columns. I can do the same thing for rows. So I'm going to wrap the rows by using =WRAPROWS. The vector is 
once again my varieties, then a comma and the wrap count will be three. Right parenthesis and enter. And you can see that I get three columns 
spread out over two rows. Now let's see what would happen if I did have a number of varieties that did not fit easily into the array that I 
created. So to demonstrate that, I'll go to cell A13, click it, press tab, and then variety seven and enter. And you can see that I have errors 
here, N/A and N/A and also the same for wrap rows. And that's because I don't have any values to go there. So it says it's not available. To fix 
that, I can edit the original formula. So I will double click cell C13. And then after the three, indicating where I want to wrap my rows, I 
will type a comma, and then I'll enter the string empty. So that will be empty in double quotes and enter. And I have the error fixed. So I have 
the varieties that you see here and then the last two cells are empty. And just for completeness, I'll go up to the top and then edit the 
formula in cell C8, have an empty string again and enter. And there we go. So as you can see, you can use wrap columns and wrap rows to create 
an array out of a single vector, either horizontal or vertical. And in the same way, if you have an array of data that you want to change to a 
row or column, you can use TOROW or TOCOL to make that happen.

Troubleshoot LAMBDA output
Selecting transcript lines in this section will navigate to timestamp in the video
- [Instructor] Like most other Excel functions, Lambda can accept single cells or cell ranges as input. However, different types of ranges can 
cause interesting problems. So in this movie, I will demonstrate one surprising behavior in Excel and describe a way to work around it with the 
goal of giving you an approach to solving problems when you encounter them in your own work. My sample file is 05 08 Troubleshoot, and you can 
find it in the chapter five folder of the Exercise Files Collection. Now, before I get started, I want to point out that this is not something 
that is unique to Lambda, Instead, this is something that I discovered when I was creating a previous movie, and I thought it was interesting 
enough to share. So my goal is to take the data, and the variety list table here, and to count the number of unique items, and to tell you 
what's going to happen. There are seven unique values. So I have nine values in this table, and three repeats and eight repeats. Otherwise there 
are no other repetitions. So there will be seven unique values. So to count the unique items in the table column there, I will type equal and 
then count A. This function counts the number of cells in a range that are not empty, so not just numbers. So I have count A, and then I want to 
find the number of unique values. So I will use the unique function, and my array is my table column here. So I will move my mouse pointer over 
its header, and when it's a downward-pointing black arrow, I'll click to select. There it goes. And I don't need to change anything else. So I 
will type two right parenthesis to close out and enter, and I get seven unique items, so that's correct. I have the same data laid out as an 
array in the range D4 through F6, and I want to show you an interesting behavior. So if I create the same type of formula, so I'll have equal, 
and that would be count A again and then unique. And the array is D4 through F6, and then close out with two right parentheses and press enter. 
I get the number nine, even though three repeats and eight repeats. So that is an interesting behavior that I discovered when I was preparing 
another movie in this course. So the question is, "How do you get around it?" Well, it turns out that if you within the formula represent the 
data as a single column, then Excel counts them accurately. So I'm going to edit the formula in I6 and instead of just counting the unique 
items, I'm going to instead use two call, which changes the data configuration from an array to a single column, a single vertical column. So I 
will type a right parenthesis, and then two column. And then I need another right parenthesis at the end to close it out and enter, and it 
counts it accurately. All right, so it works with columns. I was wondering if the same thing would work with rows. So I will edit the two-column 
function to two row, right? All my parentheses are still good, so I'll press enter and it does not work. But if I press Control + Z to go back 
to two column, then it does. So this is interesting, and again, you probably won't have the specific problem, but you might find that you aren't 
able to get the result that you want from an array. If that is the case, try transforming the data into a single column using two-column TOCOL 
to see if that fixes it.





