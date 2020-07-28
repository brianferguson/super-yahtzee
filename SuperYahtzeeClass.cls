VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuperYahtzeeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'        Name: SuperYahtzeeClass.cls
'      Author: Brian Ferguson
'     Website: https://github.com/brianferguson/super-yahtzee/
'        Date: 2020.07.28
'     Version: 1.0
'     License: CC BY-NC-SA 4.0  https://creativecommons.org/licenses/by-nc-sa/4.0/
' Description: This file defines the methods that insert check and scratch marks on the board.
'    Platform: Microsoft Office Professional Plus Excel 2013

Option Explicit
Dim oldValue As String                      ' Represents a cell's text when selected

Public Sub ClearOldValue()

    ' Clear the old value when clearing the scorecard
    ' Called when clicking the "Clear scorecard" button (See: SuperYahtzeeModule.bas)
    oldValue = ""

End Sub

Private Function GetCheckMark() As String

    GetCheckMark = ChrW(&H2713)             ' Fancy check mark

End Function

Private Function GetScratchMark() As String

    GetScratchMark = ChrW(&H2E3B)           ' Three-Em Dash

End Function

Private Sub SetCellCheckMark()

    ' Sets a check mark on double clicking the box, or removes it
    If ActiveCell.Value2 = GetCheckMark() Then
        ActiveCell.ClearContents
    Else
        ActiveCell.Value2 = GetCheckMark()
        oldValue = GetCheckMark()           ' Prevent manually changing the check mark
    End If

End Sub

Private Sub SetExtraYahtzees(ByVal cells As Range, ByVal total As Range)

    ' Calculates the bonus for any extra yahtzee's
    ' Note: If more than 3 extra yahtzees are rolled, just manually add "100"
    ' to the bonus box. If you accidentally double click the check mark box
    ' after manually editing the bonus box, it will re-calcuate the bonus
    ' based ONLY on the check marks, so keep track of your extra-extra yahtzees!
    Dim result As Integer
    Dim cell As Range

    For Each cell In cells
        If cell.Value2 = GetCheckMark() Then
            result = result + 100
        End If
    Next cell

    If result > 0 Then
        total.Value2 = result
    Else
        total.MergeArea.ClearContents
    End If

End Sub

Private Sub CalculateExtraYahtzees(ByVal Target As Range)

    If Not Intersect(Target, Range("E27:G27")) Is Nothing Then
        SetExtraYahtzees Range("E27:G27"), Range("E28")
    ElseIf Not Intersect(Target, Range("H27:J27")) Is Nothing Then
        SetExtraYahtzees Range("H27:J27"), Range("H28")
    ElseIf Not Intersect(Target, Range("K27:M27")) Is Nothing Then
        SetExtraYahtzees Range("K27:M27"), Range("K28")
    ElseIf Not Intersect(Target, Range("N27:P27")) Is Nothing Then
        SetExtraYahtzees Range("N27:P27"), Range("N28")
    ElseIf Not Intersect(Target, Range("Q27:S27")) Is Nothing Then
        SetExtraYahtzees Range("Q27:S27"), Range("Q28")
    ElseIf Not Intersect(Target, Range("T27:V27")) Is Nothing Then
        SetExtraYahtzees Range("T27:V27"), Range("T28")
    End If

End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Application.EnableEvents = False

    If Intersect(Target, Range("E27:V27")) Is Nothing Then
        Application.EnableEvents = True
        Exit Sub
    End If

    SetCellCheckMark
    Cancel = True
    CalculateExtraYahtzees Target

    Application.EnableEvents = True

End Sub

Private Sub Worksheet_Calculate()

    ' Replace any 0's with the scratch mark
    Range("E8:T28").Replace 0, GetScratchMark(), 1

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    ' Do some validation in case check mark boxes are edited manually on accident
    Application.EnableEvents = False

    If Intersect(Target, Range("E27:V27")) Is Nothing Then
        Application.EnableEvents = True
        Exit Sub
    End If

    If Not IsEmpty(Target.Value2) And oldValue = GetCheckMark() Then
        Target.Value2 = GetCheckMark()
        Application.EnableEvents = True
        Exit Sub
    End If

    If Not Target.Text = GetCheckMark() Then
        Target.ClearContents
        CalculateExtraYahtzees Target
    End If

    Application.EnableEvents = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    ' Set the "old" value before any changes, for validation later
    If Not IsNull(Target.Text) Then
        oldValue = Target.Text
    End If

End Sub
