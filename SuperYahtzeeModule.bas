Attribute VB_Name = "SuperYahtzeeModule"
'        Name: SuperYahtzeeModule.bas
'      Author: Brian Ferguson
'     Website: https://github.com/brianferguson/super-yahtzee/
'        Date: 2020.07.28
'     Version: 1.0
'     License: CC BY-NC-SA 4.0  https://creativecommons.org/licenses/by-nc-sa/4.0/
' Description: This defines the "click" event for the "Clear Scorebard" button.
'    Platform: Microsoft Office Professional Plus Excel 2013

Option Explicit

Sub Button_Clear_Scorecard_Click()

    Application.EnableEvents = False

    SuperYahtzeeClass.ClearOldValue

    Range("Upper_Section").ClearContents    ' Defined as a named range in the Workbook
    Range("Lower_Section").ClearContents

    Application.EnableEvents = True

End Sub
