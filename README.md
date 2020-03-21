# Applied-Research-Method-via-VBA
Excel VBA coursework

Week 1
TASK SET 1
(a)
(i) Run the “Hello World!” program on VBA. Experiment with leaving out parts of the program,
to see what error messages you get. Experiment to find out what happens when more arguments
are added in the function msgbox().
(ii) Write a VBA subroutine program that prints out the phrase “hello, world:” several times
following the effect of a variety of specifications in printing ``hello, world'' (15 characters) .
The dialog box output should be as follows:
:hello, world:
:hello, world :
:hello, wor:
:hello, world :
:hello, world :
:hello, world :
: hello, wor:
:hello, wor :
(iii) Write a VBA modular subroutine that prints out using the msgbox() a previously initialized
floating-point variable representing a bond’s yield with in ways following various output
formatting specifications (10 characters wide). The output should be as follows:
:Bond yield: 0.045__:
: Bond yield: 0.045:
: Bond yield: 0.04 :
:Bond yield: 4.5%_ _:
:Bond yield: 4.50%_ :

(b)
Write a VBA modular subroutinme that computes the future value of $100.00 received today and
values it a year from now. Assume an annual compound interest rate of 0.0245. The computed
value should be output right adjusted, with two decimal places, and with the £ symbol appended.

TASK SET 2
(a)
(i) Write a VBA modular subroutine that prompts the user (i.e. use InputBox()) for a monetary
amount to be received at time t (you may input this interactively or initialize statically), the
discounting rate and computes the discounted value.
(ii) Rewrite the solution of (i) so that the frequency of compounded is included and specified
through an input prompt.

(b)
Write a VBA modular subroutine that computes the price of a single equity share, assuming that
the stock has just paid a dividend of £0.60, and future dividend payouts are expected to be
precisely at this level forever. The required rate of return of this stock is 10 per cent.

TASK SET 3
(a)
i.Write a VBA modular subroutine that prompts the user to enter a real number, computers the
square root of the number and prints it out in two decimal places.
ii.Write a VBA modular subroutine that prompts the user to enter a real number and an integer
number n, computers the number in power n and prints it out in two decimal places.
iii.Write a VBA modular subroutine that generates a random number between 0.0 and 1.0 and
prints it out in two decimal spaces.

(b)
i.Write a VBA modular subroutine that prompts the user for his/her initials, and prints them out
in upper case.
ii.Write a VBA modular subroutine that prompts the user for his/her name and prints it out
(dialog box or excel) in lower case.

TASK SET 4
(a)
As sume you plan to retire in t years and want to accumulate enough by then to provide
yourself with $30,000 a year for 15 years. The interest rate is 10 percent. Write a program
that prompts the user for the time t to retirement and computes the amount accumulated
by the time you retire. The time entered should be positive.

(b)
As sume you plan to retire in t years and want to accumulate enough by then to provide
yourself with $30,000 a year for 15 years. The interest rate is 10 percent. Write a VBA
modular subroutine that prompts the user for the time t to retirement and computes the
amount accumulated by the time you retire.

Week2
TASK SET 1
(a)
Wr i te a VBA program that prompts the user for his/her initials, validates the case of
the characters (upper or lower) and prints them out in upper case.

(b)
Write a VBA program that it prompts the user for the time t validated to be positive only,
an amount to be received at time t validated to a positive real number, the discounting
rate validated to be between 0.0 and 1.0 and computes the discounted value. Make use of
arithmetical, short cut assignment, and logical operators where appropriate.

(c)
Write a VBA program that prompts the user for a character, checks to see whether it is a
white space character. White space characters are ‘ ‘, ‘\n’, ‘\t’. If the non-white character
is entered in upper case it converts it to lower case. The output could be as follows:
Enter a character: X
The character in lower case is x
or
Enter a character:
You entered a white space character!

TASK SET 2
(a)
Consider an equity based call option with an exercise price of K= 100. Write a VBA
program that prompts the user the for market stock price at time t, it validates the stock
price against the exercise and computes the option payoff and a comment on whether the 
“right to buy” is exercised or not. Use the time t option payoff formula C = MAX[S - K,0].
Create various versions of the program using different conditional statement.

(b)
Assume a simple case of a two-fund separation problem with stock x with return of 15%
and stock y with a rate of return on 21% combined in a portfolio. Write a program that
prompts the user for the capital weights and computes the portfolio’s rate of return. Add
validation safeguards where appropriate.

TASK SET 3
(a)
Write a VBA program that prompts the user to enter his/her name one character at a time
as part of a loop and assigns them to a string variable. The program should also print the
name one character at a time using another for loop.

(b)
Write a VBA program that prompts the user to enter a sentence and it should use a for
loop to pick up and counts the number of times the character ‘a’ or ‘A’ is present in the
sentence. The output could be as follows:
Enter a sentence: Hello John! Are you going to school today?
The frequency of the character ‘a’ (and ‘A’) is 2.
