# An-Analysis-of-Refactoring-VBA-code

## Overview of Project

We have learned that the workbook sent to Steve was very useful to him and he is very grateful for its practicality.
However, we are concerned that the code he has will be useful for his future analyzes that could contain more data than what we have analyzed ..
For this reason we have decided to refactor the code considering the weight of the script as its timer so that we guarantee its usefulness for any scenario to which it is subjected.

### Purpose

The purpose of this analysis is to provide Steve with a code that will serve as a tool for his future stock analysis that is practical, flexible, understandable and efficient.

## Analysis and Challenges

The main challenge we face is to delve into the technical aspects of the code so that we can reduce the space and weight it occupies to make it more efficient without altering the results.

The main modifications that were made to the code with respect to the original were the following:

1. Integrate everything in a single code: the update time and the format for the cells were added to the code of the analysis of all the actions.

![Timer in the code](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/Timer%20in%20the%20code.png)

![Formatting in the code](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/Formatting%20in%20the%20code.png)

2. Changing the type of variables without altering their participation in the code:

We define the variable Dim tickerVolumes As Long since a long data type accepts a wider range of integers. According to Docs.Microsoft.com The Long data type widens to Decimal, Single, or Double. This means you can convert Long to any one of these types.

Because declaring a smaller variable type reduces update time.
However, it is important to say that this difference is only noticeable if you perform many thousands of operations which is what we are looking for for Steve's needs.

- We change the property to call the cells from .Value to .Value2 that gives the underlying value of the cell.
As it involves no formatting, .Value2 is faster than .Value. .Value2 is faster than .Value when processing numbers (there is no significant difference with text).

![Value2 property](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/Value2%20property.png)

3. We reduced the number of lines by listing the new variables in a single statement using commas.

![List variables using commas](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/List%20variables%20using%20commas.png)

4. We use the With statement to format the same range of variables. The above Reduces memory space.

![With statement in the code](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/With%20statement%20in%20the%20code.png)


Once the above was done, we verified that the outputs did not change with respect to the original code and the execution times were compared.

In the following image you can see the outputs with the refactored code for both years of analysis.

![Outputs 2017](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/Outputs%202017.png)
![Outputs 2018](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/Outputs%202018.png)

Regarding the execution time, we had the following times registered for the original code:

![Module2_VBA_Time](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/Module2_VBA_Time.png)

For the refactored code we have:

![VBA_Challenge_2017](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/VBA_Challenge_2018.png)

The comparison allows to observe an improvement. It should be noted that when carrying out the activity I realized that the times change every time I carry out the update. This means that it is a random indicator.
In addition, it also depends on the level of fatigue of the machine, that is, how hot are the electronic components of the machine that I am using.

## Analysis of the obtained results

Regarding the data we have the following:

We can be seen in the graph below that the returns of almost all stocks fell in 2018 compared to 2017.
In the year 2017 it seemed that the best option to invest was DQ followed by SEDG, however the year 2018 shows us that these actions seem to be volatile which means that there is a high risk when investing in them.
For its part, the comparative graph also allows us to observe that the action that had a more stable behavior is ENPH, since in both years it obtained positive returns and without radical changes.
With the information we have at hand, we can recommend Steve to propose to his parents to invest in the most stable action.
It is recommended to carry out an analysis with more years of comparison to see the behavior pattern of the shares and give a more robust proposal. 


![All Stocks Returns](https://github.com/LAURYMEOW/An-Analysis-of-Refactoring-VBA-code/blob/main/All%20Stocks%20Returns.png)


## Summary: In a summary statement, address the following questions.

What are the advantages or disadvantages of refactoring code?

- The refactoring of the code has more advantages than disadvantages since it allows the code to be more efficient.
- The flexibility that we guarantee when making our primer template allows precisely this refactoring.
- Refactoring in turn leads us to go deeper into the code and therefore improve it.
- The downside may be that it is less detailed, but the ability to leave comments largely eliminates this disadvantage.

How do these pros and cons apply to refactoring the original VBA script?

- The first step to refactoring is to understand each line of code so that when making changes we know what is being done and what is expected to be.
- As I already mentioned in the answer to the previous question, the flexibility of the template that we made in the module, which is nothing more than the well-established sequence and the concise comments, it allowed us an easier handling of it.
- Having everything required in a single code allows for a cleaner macro.
