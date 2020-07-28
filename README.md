# super-yahtzee

A simple twist to the classic dice game.

###### Note: **This is just a scorecard!** You can print the [image file](https://github.com/brianferguson/super-yahtzee/blob/master/SuperYahtzeeScorecard.png) or use the [Excel file](https://github.com/brianferguson/super-yahtzee/blob/master/SuperYahtzee.xlsm) as your scorecard. The Excel file is locked to prevent any formulas from accidentally getting overwritten. The Excel file is macro-enabled to provide some simple functions like clearing the scorecard and changing 0's to scratch marks. It also allows for double clicking on the bonus Yahtzee cells to add/remove a checkmark. You can view the VBA code [here](#VBA-files).

#

#### How to play:
Super Yahtzee is just like regular Yahtzee, but with an extra die (for a total of **6**) and an **extra roll** of the dice on your turn (for a total of **4**). Because of the extra die, there are a few changes to the scorecard.

##### Differences from regular Yahtzee:
1. The bonus for the Upper section is now 45 points (instead of 35), and you have to achieve 84 points to score the bonus (instead of 63 points).
2. `3 of a kind` is replaced with `4 of a kind`.
3. `4 of a kind` is replaced with `5 of a kind`.
4. A `Full House` now can be 4 of the same number and 2 of a different number OR 3 of the same number and 3 of a different number.<br />
Example 1: `4 Ones, 2 Threes`<br />
Example 2: `3 Fives, 3 Sixes`
5. A new category called `3 Pair` that scores 35 points. `3 Pair` consists of 3 **distinct** pairs of numbers. 2 of the 3 pairs cannot be of the same number (since that would be a `Full House`).<br />
Valid example: `2 Ones, 2 Threes, 2 Fours`<br />
Invalid example: <code>2 <b><em>Ones</em></b>, 2 <b><em>Ones</em></b>, 2 Threes</code> (this is a `Full House`)
6. A new category called `Ex Lg Straight` or extra large straight that scores 50 points. This is a sequence using all 6 dice.
7. Yahtzee using all 6 dice instead of 5. Yahtzee is now worth 75 points.

Check out the [scorecard](https://github.com/brianferguson/super-yahtzee/blob/master/SuperYahtzeeScorecard.png)!

#

#### Background:
My family loves to play classic board and dice games. One of our favorites is Yahtzee. Sometime in the mid 2000's, our family had an incredible streak of bonus Yahtzee's and high scores. Our scores were consistently in the 600-700 range during this time. Once this streak came to an end, we started to miss all the excitement of rolling all those Yahtzee's. We even moved onto other board games for a while. At some point we got a new set of dice and for some reason it came with 6 dice instead of 5. Everytime we played Yahtzee, we had to leave 1 die out. One day, when we started to play, we forgot to take out that extra die. On the very first roll, an extra large straight came out of the cup. We joked that this should count for some extra points...and that's when it hit. Why not play with all 6 dice? So, after a little experimentation, Super Yahtzee is what we came up with. This is the only way we play Yahtzee now!

I hope you and your family and friends enjoy Super Yahtzee as much as my family does!

#

#### VBA files:
* [SuperYahtzeeClass.cls](https://github.com/brianferguson/super-yahtzee/blob/master/SuperYahtzeeClass.cls)
* [SuperYahtzeeModule.bas](https://github.com/brianferguson/super-yahtzee/blob/master/SuperYahtzeeModule.bas)

#

#### Credits:
* [Yahtzee logo](https://github.com/brianferguson/super-yahtzee/blob/master/resources/logo.png) from [Wikipedia.org](https://en.wikipedia.org/wiki/Yahtzee#/media/File:Yahtzee_logo.svg) [rotated using Paint.NET]
* [Dice images](https://github.com/brianferguson/super-yahtzee/blob/master/resources/) from [ClipArtMax](https://www.clipartmax.com/)
