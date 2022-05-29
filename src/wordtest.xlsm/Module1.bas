Attribute VB_Name = "Module1"
Option Explicit

'This macro is written by Taisei FUJISAWA to realize easy output of English word test.
'All rights reserved.
'Last revise:2019/01/06

'Note:Refer when starting new wordbook
'*Write header(name of edited wordbook) at three praces(Eng->Jap,Jap->Eng,Answer)
'*White the number of words in wordbook at constant number in the program at two praces(This Workbook and Module1)


'This macro runs when Macrobutton is clicked.This macro carries out printing of contents of the test.


Sub Test_Print()


  Const AllWords = 1900     'Costant Number,corresponding the number of words in wordbook,ONE MORE PLACE AT This Workbook.
  
  
  If AllWords < Worksheets(1).Scope.Text / 2 * 2 + Worksheets(1).Min.Text - 1 Then
    MsgBox ("選択範囲が不適当です。")
    Exit Sub
  End If
  'Error message about setting of unappropriate Scope and Min.;/2*2 is necessary to consider Worksheets(1).Scope.Text to be not letters but value.
  
  
  Dim Set_Number As Integer     'the number of print sets
  For Set_Number = 1 To Worksheets(1).Sets.Text     'repeat printing;FOR
    Dim i As Integer    'for carry out VLOOKUP Function
    Dim num As Integer      'for preventing RANDOM NUMBER from being repeated
    Dim Flag(1 To AllWords) As Integer      'set all numbers of word's number;this parameter is ARRAY
    Randomize
    
    For num = 1 To AllWords
        Flag(num) = 0
    Next
    'reset Flag
    
    
    For i = 1 To 25
    
        Do
            num = Int(Rnd() * Worksheets(1).Scope.Text + Worksheets(1).Min.Text)
        Loop Until Flag(num) = 0
        Worksheets(4).Cells(2 + i, 1).Value = num
        Flag(num) = 1
        'RANDOM NUMBER output;not to be repeated
        
        Worksheets(4).Cells(2 + i, 2).Value = WorksheetFunction.VLookup(Worksheets(4).Cells(2 + i, 1), Range(Worksheets(1).Cells(11, 1), Worksheets(1).Cells(AllWords + 10, 4)), 2, False)
        Worksheets(4).Cells(2 + i, 3).Value = WorksheetFunction.VLookup(Worksheets(4).Cells(2 + i, 1), Range(Worksheets(1).Cells(11, 1), Worksheets(1).Cells(AllWords + 10, 4)), 3, False)
        Worksheets(4).Cells(2 + i, 4).Value = WorksheetFunction.VLookup(Worksheets(4).Cells(2 + i, 1), Range(Worksheets(1).Cells(11, 1), Worksheets(1).Cells(AllWords + 10, 4)), 4, False)
        'VLOOKUP Function;English,Japanese and part of speech
    Next
    'display contents at worksheets(4)'s left rows
    
    
    For i = 1 To 25
    
        Do
            num = Int(Rnd() * Worksheets(1).Scope.Value + Worksheets(1).Min.Value)
        Loop Until Flag(num) = 0
        Worksheets(4).Cells(2 + i, 7).Value = num
        Flag(num) = 1
        'RANDOM NUMBER output;not to be repeated
        
        Worksheets(4).Cells(2 + i, 8).Value = WorksheetFunction.VLookup(Worksheets(4).Cells(2 + i, 7), Range(Worksheets(1).Cells(11, 1), Worksheets(1).Cells(AllWords + 10, 4)), 2, False)
        Worksheets(4).Cells(2 + i, 9).Value = WorksheetFunction.VLookup(Worksheets(4).Cells(2 + i, 7), Range(Worksheets(1).Cells(11, 1), Worksheets(1).Cells(AllWords + 10, 4)), 3, False)
        Worksheets(4).Cells(2 + i, 10).Value = WorksheetFunction.VLookup(Worksheets(4).Cells(2 + i, 7), Range(Worksheets(1).Cells(11, 1), Worksheets(1).Cells(AllWords + 10, 4)), 4, False)
        'VLOOKUP Function;English,Japanese and part of speech
    Next
    'display contents at worksheets(4)'s right rows
    
    
    Worksheets(4).Cells(1, 2).Value = Worksheets(1).Min.Text
    Worksheets(4).Cells(1, 4).Value = Worksheets(1).Scope.Text / 2 * 2 + Worksheets(1).Min.Text - 1
    'display test scope;/2*2 is necessary to consider Worksheets(1).Scope.Text to be not letters but value.
    
    
    Worksheets(4).Range("A1:K27").Copy (Worksheets(2).Range("A1"))
    Worksheets(4).Range("A1:K27").Copy (Worksheets(3).Range("A1"))
    'copy worksheets(4) to worksheets(2) and (3)
    
    
    Worksheets(2).Range("D3:F27").ClearContents
    Worksheets(2).Range("J3:K27").ClearContents
    Worksheets(3).Range("B3:B27").ClearContents
    Worksheets(3).Range("H3:H27").ClearContents
    'clear unnecessary contents in worksheets(2) and (3)
    
    
    If Worksheets(1).Kind.Text = "日本語→英語" Then
        Worksheets(3).Range("A1:K27").PrintOut
    Else
        Worksheets(2).Range("A1:K27").PrintOut
    End If
    'specify the kind of test pattern;printing right one
    
    
    Worksheets(4).Range("A1:K27").PrintOut
    'print answer
    
    
  Next      'repeat printing;NEXT
  
End Sub

