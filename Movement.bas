Attribute VB_Name = "Movement"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'**********************************************************************************'
'Ok a bit about myself... My name is Ray and i enjoy programming. I am not the best'
'but i do my best and i always am willing to try and experience new things. I don't'
'get everything right in programming but i don't often give up unless i can't teach'
'myself as i have taught myself basically everything i know......School gave me the'
'comeplete basics and i read help file after help file and other things from planet'
'source code which have helped me out. I made this little program for a girl in USA'
'she asked me for this. i made her do most of the graphics cos i really suck at em '
'i don't often make something unless i am asked to or have a need to....I used the '
'same basics from one of my earlier programs called rayman a pacman game which did '
'pretty well i thought. this has the same basics and is just a little more advanced'
'i had never used a database EVER before so that i had to learn quite a lot about. '
'But i managed and i think it worked out well. Other than that most of it is pretty'
'basic but i quite proud of it considering it is almost all my own work if not it's'
'all edited to suit my needs. and as far as i know it all works except for the stuf'
'that i didn't understand but that i couldn't be bothered with and the girl from US'
'wanted it quicker than i could have figured it out. So i left it as is and it was '
'good enough for me. Anyway Hope you enjoy it. And by all means vote for me at PSC '
'if you liked it...Hell even if you didn't like it vote...Anything i can improve on'
'         tell me. I would love some feedback either on PSC or e-mail me @         '
'                            brindleyray@hotmail.com                               '
'**********************************************************************************'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'here we get into the big and meaningful hard shit...this part is the backbone. of
'virtually everything else there is in this project.
'i will only comment on the first and last of the movements cos the other 2 are the
'same sort of thing

Public Sub MoveUp()
  frmMaze.tmrTime.Enabled = True    'start the basics
  frmMaze.tmrBonuses.Enabled = True
  NewTop = (frmMaze.imgChar.Top - Pos + 9)  'ok find where it is going next... - pos (pos is declared in declarations as 19) + 9 = the top of the image - 19 + 9 making it where it will end up
  NewLeft = frmMaze.imgChar.Left            'not using this really because it is only moving up
  Gridrow = Int((NewTop - 10) / 10 + 1)     'finds the actual place in the grid where it will end up.
  GridCol = Int((NewLeft - 10) / 10 + 1)    'same as previous
    If Grid(Gridrow, GridCol) <> "1" And (NewLeft - 10) Mod 10 = 0 Then   'checks the grid to see whether or not the space it is going to move to is appropriate to move to
      If Grid(Gridrow, GridCol) = "5" Then  'if they are near a question
        frmMaze.imgBeforeQu.Visible = True  'show the near question picture
        frmMaze.imgQuestion.Visible = False 'hide the on question picture
      ElseIf Grid(Gridrow, GridCol) <> "5" Then 'make sure to hide the image if it isn't near a question
        frmMaze.imgBeforeQu.Visible = False
      End If
      frmMaze.imgChar.Top = (frmMaze.imgChar.Top - 10)  'finally move the character after checking where it is moving to
      If Grid(Gridrow, GridCol) = "2" Then  'if they land upon a question do the following
        frmMaze.lblCover.BackStyle = 1
        frmMaze.imgBeforeQu.Visible = False
        frmMaze.imgQuestion.Visible = True
        frmMaze.tmrTime.Enabled = False
        frmMaze.Enabled = False
        frmQuestions.tmrTime.Enabled = True
        frmQuestions.Show
      End If
    End If
End Sub

Public Sub MoveLeft()
  frmMaze.tmrTime.Enabled = True
  frmMaze.tmrBonuses.Enabled = True
  NewTop = frmMaze.imgChar.Top
  NewLeft = (frmMaze.imgChar.Left - Pos + 9)
  Gridrow = Int((NewTop - 10) / 10 + 1)
  GridCol = Int((NewLeft - 10) / 10 + 1)
    If Grid(Gridrow, GridCol) <> "1" And (NewTop - 10) Mod 10 = 0 Then
      If Grid(Gridrow, GridCol) = "5" Then
        frmMaze.imgBeforeQu.Visible = True
        frmMaze.imgQuestion.Visible = False
      ElseIf Grid(Gridrow, GridCol) <> "5" Then
        frmMaze.imgBeforeQu.Visible = False
      End If
      frmMaze.imgChar.Left = (frmMaze.imgChar.Left - 10)
      If Grid(Gridrow, GridCol) = "2" Then
        frmMaze.lblCover.BackStyle = 1
        frmMaze.imgBeforeQu.Visible = False
        frmMaze.imgQuestion.Visible = True
        frmMaze.tmrTime.Enabled = False
        frmMaze.Enabled = False
        frmQuestions.tmrTime.Enabled = True
        frmQuestions.Show
      End If
    End If
End Sub

Public Sub MoveRight()
  frmMaze.tmrTime.Enabled = True
  frmMaze.tmrBonuses.Enabled = True
  NewTop = frmMaze.imgChar.Top
  NewLeft = (frmMaze.imgChar.Left + Pos - 1)
  Gridrow = Int((NewTop - 10) / 10 + 1)
  GridCol = Int((NewLeft - 10) / 10 + 1)
    If Grid(Gridrow, GridCol) <> "1" And (NewTop - 10) Mod 10 = 0 Then
      If Grid(Gridrow, GridCol) = "5" Then
        frmMaze.imgBeforeQu.Visible = True
        frmMaze.imgQuestion.Visible = False
      ElseIf Grid(Gridrow, GridCol) <> "5" Then
        frmMaze.imgBeforeQu.Visible = False
      End If
      frmMaze.imgChar.Left = (frmMaze.imgChar.Left + 10)
      If Grid(Gridrow, GridCol) = "2" Then
        frmMaze.lblCover.BackStyle = 1
        frmMaze.imgBeforeQu.Visible = False
        frmMaze.imgQuestion.Visible = True
        frmMaze.tmrTime.Enabled = False
        frmMaze.Enabled = False
        frmQuestions.tmrTime.Enabled = True
        frmQuestions.Show
      End If
    End If
End Sub

Public Sub MoveDown()
  frmMaze.tmrTime.Enabled = True              'same as others
  frmMaze.tmrBonuses.Enabled = True           '
  NewTop = (frmMaze.imgChar.Top + Pos - 1)    '
  NewLeft = frmMaze.imgChar.Left              '
  Gridrow = Int((NewTop - 10) / 10 + 1)       '
  GridCol = Int((NewLeft - 10) / 10 + 1)      '
    If Grid(Gridrow, GridCol) <> "1" And (NewLeft - 10) Mod 10 = 0 Then
      If Grid(Gridrow, GridCol) = "5" Then    '
        frmMaze.imgBeforeQu.Visible = True    '
        frmMaze.imgQuestion.Visible = False   '
      ElseIf Grid(Gridrow, GridCol) <> "5" Then
        frmMaze.imgQuestion.Visible = False   '
      End If                                  '
      frmMaze.imgChar.Top = (frmMaze.imgChar.Top + 10)
      If Grid(Gridrow, GridCol) = "2" Then    '
        frmMaze.lblCover.BackStyle = 1
        frmMaze.imgBeforeQu.Visible = False   '
        frmMaze.imgQuestion.Visible = True    '
        frmMaze.tmrTime.Enabled = False       '
        frmMaze.Enabled = False               '
        frmQuestions.tmrTime.Enabled = True   '
        frmQuestions.Show                     '
      End If
      If Grid(Gridrow, GridCol) = "4" Then    'only finishes when you press down so i didn't bother putting it in the others
        frmMaze.tmrTime.Enabled = False       'stop the timer
        CheckTime                             'call the procedure the add bonuses for finishing in a certain time
        Response = MsgBox(("You have finished level " & Level & " with a time of " & frmMaze.lblMin.Caption & ":" & frmMaze.lblSec1.Caption & frmMaze.lblsec2.Caption & " Congratulations!"), vbOKOnly)
        If Response = vbOK Then               'response is a fancy way to do a message box...it gives you the options of what happens if different buttons are pressed...i like using it on different occasions
          frmMaze.lblProgress.Caption = frmMaze.lblProgress.Caption + frmMaze.lblScore.Caption    'add the score to the total score
          If Level = 3 Then                   'if the mazes are finished then
            Congratulations
            HighScores                        'check to see how well the went
            frmHighScores.Show                'show them how well they went
          End If
          If Level = 3 Then                   'reset the maze for next time round
            Level = 0
          End If
          Level = Level + 1                   'go to the next maze
          MakeGrid                            'and make the next maze
        End If
      End If
    End If
End Sub

'bigtime complications here
'i don't understand it all myself hehehe. well maybe i do

Public Sub MakeGrid()
  Unload frmQuestions   'i had a problem with changing the questions from a different database till i relised when it would need to be loaded

  Levelfile = FreeFile  'choose the file for freefile
  Open "maze" & Level & ".mze" For Input As #Levelfile  'open the file corresponding with the current level
   
   For I = 1 To 50  'from the beggining to the end of the lines
    Line Input #Levelfile, FullGrid 'Read it all
    For J = 1 To 50 'from the beggining to the end of the width
    
      K = K + 1 'go to the next one
      SingleGrid = Mid(FullGrid, J, 1)  'saying what exactly is the singlegrid and where it is situated in other words just one number at a time
      Grid(I, J) = SingleGrid           'the single number is where? where the current I and J are
      
      Select Case CInt(LCase(SingleGrid)) 'what the singlegrid is and what to do with it
        Case 7  'if it's a 7
          pictype = 1 'what picture to put in it's place
          frmMaze.imgApple.Top = (I * 10)   'this and 6 aren't needed because i didn't get it actually working
          frmMaze.imgApple.Left = (J * 10)
        Case 6
          pictype = 1
          frmMaze.imgBanana.Top = (I * 10)
          frmMaze.imgBanana.Left = (J * 10)
        Case 5
          pictype = 1 'this just needs to be invisable so it is just a blank. thats what the picture is in the picture on the maze form
        Case 4  'the end of the maze
          pictype = 1
          frmMaze.lblFinish.Left = (J * 10) + 10  'puts the label into place of the finish line
          frmMaze.lblFinish.Top = (I * 10) + 25
          frmMaze.lblFinish.Visible = True
          frmMaze.tmrStart.Enabled = True
        Case 3  'the beginning of the maze
          pictype = 1
          frmMaze.lblStart.Left = (J * 10) + 10 'puts the label into place of start
          frmMaze.lblStart.Top = (I * 10) - 25
          frmMaze.lblStart.Visible = True
          frmMaze.imgChar.Visible = True  'putting the character in place
          frmMaze.imgChar.Top = (I * 10)
          frmMaze.imgChar.Left = (J * 10)
        Case 2  'these are the questions
          pictype = 1
        Case 1  'if it's a wall
          pictype = 0
        Case 0  'if its a blank
          pictype = 1
      End Select
        'i didn't come up with this idea either someone showed it to me...much easier way to do it.
        'first of all it says where to paint a picture of what is has just read
        'next it checks where to paint that picture
        'what the width and height of the picture to be painted
        'what picture it is painting in this case it gets all of the pictures out of one picture
        'and just extra bits and pieces at the end to keep it all problem free
        frmMaze.PaintPicture frmMaze.pic.Picture, (J * DestPicWidth), (I * DestPicHeight), DestPicWidth, DestPicHeight, pictype * 8, 0, Source_Pic_Width, Source_Pic_Height, vbSrcCopy
    
    Next J  'repeat the long process again
  Next I
  
  Close #Levelfile  'close the file
  
  If Level = 1 Then 'depending on the level what to do about it
    frmMaze.lblMin.Caption = 5
  ElseIf Level = 2 Then
    frmMaze.lblMin.Caption = 5
  ElseIf Level = 3 Then
    frmMaze.lblMin.Caption = 5
  End If
  frmMaze.lblSec1.Caption = 0
  frmMaze.lblsec2.Caption = 0
  frmMaze.lblScore.Caption = 0
End Sub

Public Sub CheckTime()  'this is just a basic checker to see what bonus they get depending on how much time they have left to finish the maze
  If frmMaze.lblMin.Caption >= 5 And frmMaze.lblMin.Caption < 6 Then
    frmMaze.lblProgress.Caption = frmMaze.lblProgress.Caption + Int(300 / 4)
  ElseIf frmMaze.lblMin.Caption >= 3 And frmMaze.lblMin.Caption < 4 Then
    frmMaze.lblProgress.Caption = frmMaze.lblProgress.Caption + Int(240 / 4)
  ElseIf frmMaze.lblMin.Caption >= 2 And frmMaze.lblMin.Caption < 3 Then
    frmMaze.lblProgress.Caption = frmMaze.lblProgress.Caption + Int(180 / 4)
  ElseIf frmMaze.lblMin.Caption >= 1 And frmMaze.lblMin.Caption < 2 Then
    frmMaze.lblProgress.Caption = frmMaze.lblProgress.Caption + Int(120 / 4)
  ElseIf frmMaze.lblMin.Caption >= 0 And frmMaze.lblMin.Caption < 1 Then
    frmMaze.lblProgress.Caption = frmMaze.lblProgress.Caption + Int(60 / 4)
  End If
End Sub
