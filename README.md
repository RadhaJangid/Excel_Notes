# Excel_Notes
## Excel - 
when we don't have any workbook open, this is called Excel Application. We can make multiple applications in the computer.

WorkBook - WorkBook is the workbook within an Excel Application.

Worksheet -  we can keep create as many sheets as we want.

Cell - Within each worksheet we have got cells. Each cell has its own notation. we can also rename or change the cell notation. Each cell has got different attributes.

Ranges - Combination of multiple cells are called "ranges" and single one is called "cell". we can also named the range.

We would need to enable the Developer Tab. Steps -
Options -> Customize Ribbon -> Click on Developer(On the right side it is available)

How do we go to VBA Application? How can we write code? Where we can write code? What are different windows over there? What are the different pannels are available? 
How to go to visual basic application?

One way to go is go to developer tab and then click on visual basic. This will open the visual basic application window. Shortcut - Alt + F11. Another option is right click on the sheet and click on view code. which will actually take you to the visual basic window.

- Immediate Window - is a debugging window. you want to check anything, you can do it with immediate window.How do we do that?? Put a ? and try the code what you want to see the result.
- Local Window
- Watch Window
- Property Window - Property Window basically  gives the properties all these whatever we select(In VBA project)
  
Inside the VBA project we have got excel which is our excel and we have got all the sheets which is there in our excel file. Once we delete one of the sheet it  goes from here also. we can create a code within a sheet(double click the sheet). We can create new module to write set of codes. So what's a difference??
So when we write code in the sheet and say you delete the sheet the entire code goes off because we are writting it within the sheet.

In Excel we have got options to hide a sheet. we can basically hide any sheet we want and if i go to unhide we can see the sheet there. The same thing can be done with VBA as well. So if i go to sheet2, I have got the visibility option(In Property Window).
Three visibility option: visible, hidden, very hidden(require VBA to unhide).

Every sheet has the sheet name. There are two different sheet names. The internal name of the sheet is not changed when we change the name of sheet outside.how it is important?? Once we create the Excel VBA script and give it to someone anyway anyone can go and change the name of the sheet and if we change the name of the sheet we donot want to mashup your code and change the name of the sheet. so that's where the internal name of the sheet helps. we can change the internal name of the sheet as well and that can be used in your code whenever we want and so basically if  some goes here and change the name of the sheet to first name, this will not affect our code because we will bw still using sheet name.

We write the code within the sheet only when it is necessary.The reason is if someone goes to sheet and delete it, the entire code that we've written is gone. Whenever we need to do programming. we always do it within a module.
How do we create a module??
project -> right click insert -> module.
If user deletes any other sheet it is not affected the module still remains.







