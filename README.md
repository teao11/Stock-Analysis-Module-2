# Stock-Analysis-Module-2

<h2>Overview of Project</h2>

The goal of this project was to utilize the skills we had learnt throughout the VBA week 2 module coursework to create VBA code, allowing a user to perform an analysis on a list of stocks. This code was meant to be flexible, enabling the user to select the desired year, outputting the results in a nicely formatted table!

We built on the work that we had done by refactoring the original code, in an attempt to make the new macro faster!

<h2>Results of the Project</h2>

The original code that we created throughout the coursework did a great job of teaching us the fundamentals of VBA -- for loops, declaring variables, conditional formatting, control statements etc. however the way in which we addressed the problem was NOT the most efficient. The main issue, is that we did not break down the code into bite sized sections! Instead, we had one massive for loop (with nested for loop) tackling the majority of the code:

<img width="1050" alt="Original Code - One For Loop" src="https://user-images.githubusercontent.com/46773181/118440704-ec448900-b69c-11eb-8345-eed5d690062d.png">

And the for loop continues below!

<img width="1049" alt="Original Code - For Loop Cont" src="https://user-images.githubusercontent.com/46773181/118440736-f797b480-b69c-11eb-9b19-769030dcb0ca.png">

The speed of the original macro is showcased below for 2017:

<img width="1419" alt="Original Code - 2017 Speed" src="https://user-images.githubusercontent.com/46773181/118441095-860c3600-b69d-11eb-9a9f-f3b1dc5957bd.png">

And again for 2018:

<img width="1419" alt="Original Code - 2018 Speed" src="https://user-images.githubusercontent.com/46773181/118441114-8c9aad80-b69d-11eb-90c5-4d33b69b535e.png">

Although this code ran relatively fast, there was certainly room for efficiency. This was the meat of our module 2 challenge! We addressed these inefficiencies in a couple of ways. Firstly, we broke our massive nested for loop into two distinct sized chunks, that completed the various blocks of the original for loop into two sections (assigning the volume + values for return AND outputting the values into the table)

<img width="1003" alt="Refactored Code - Breaking Out the For Loop" src="https://user-images.githubusercontent.com/46773181/118441482-1cd8f280-b69e-11eb-9409-43cd28d1f1e1.png">

Secondly, we created 3 different arrays to hold the various outputs for our table. This would make it very easy to output, as we could simply loop through the array and reference the desired index seperately, versus handling this in one massive nested loop.

<img width="1029" alt="Refactored Code - Table Referencing Arrays" src="https://user-images.githubusercontent.com/46773181/118441527-2a8e7800-b69e-11eb-90c6-dfbd1a6b4e4d.png">

The result was increased speed for our macro! 2017 ran much faster with the refactored code:

<img width="1415" alt="Refactored Code - 2017 Time Taken" src="https://user-images.githubusercontent.com/46773181/118442005-d9cb4f00-b69e-11eb-930e-1e25c9cac64a.png">

As did 2018 with the refactored code:

<img width="1422" alt="Refactored Code - 2018 Time Taken" src="https://user-images.githubusercontent.com/46773181/118442032-e2238a00-b69e-11eb-8b26-38bd18db0627.png">

<h2>Summary</h2>

<h4>What are the advantages or disadvantages of refactoring code?</h4>

The advantages of refactoring code certainly include addressing inefficiencies that creep in when tackling the problem initially. When we first write code, we may not have handled the solution as elequently as possible, rather trying to get a solution in place that works. Refactoring allows us to learn where the inefficiencies are, and how we can build on them by implementing more logical code.

We may also have designed the code without considering other dependencies, as they may not have existed when we first generated our solution. Refactoring gives us the advantage of making our code more suitable to our current needs, while keeping the functionality originally in place.

Disadvantages certainly include the risk of misinterpreting code. It is possible that the original person that created the code is no longer at the company, and if the code is poorly documented, refactoring can take a LONG TIME and will be an extremely painful process!

You also run the risk of removing certain functionality that is required, and possibly impacting other areas of the company / solution that leveraged the original code!

<h4>How do these pros and cons apply to refactoring the original VBA script?</h4>

Thankfully, the pros are certainly relevant here. As we became more familiar with VBA, we learnt that what we initially been taught and instructed was not necessarily the most efficient way to tackle the problem (although the point here was to TEACH us VBA fundamentals, and NOT to be efficient!).

The cons do not really apply, as we were the ones that generated the original code, meaning it should have been easy to interpret and well documented!! However, we must realize that this might not always be the case (:
