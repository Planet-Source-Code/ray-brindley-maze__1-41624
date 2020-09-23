Attribute VB_Name = "Declarations"
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
'For Maze form
Public Images                       'Did actually use this one in the end couldn't get it working
Public Response
'For Movement
Public Grid(50, 50)                 'how big the grid is
Public Gridrow, GridCol
Public FullGrid, SingleGrid, I, J, K, Level, OldLevel
Public NewTop, NewLeft, Pos
Public DestPicWidth As Integer
Public DestPicHeight As Integer
Public Const Source_Pic_Width = 8
Public Const Source_Pic_Height = 8
'For Database
Public DB As DAO.Database           'Holds the database
Public WS As Workspace              'Holds the DB Workspace
Public RS As Recordset              'Holds the Record Set
Public intTotalRecs As Integer      'Holds the total number of entrys
Public intRndValue As Integer       'Holds random number upto total entrys
'For Question Form
Public Question, Answer1, Answer2, Answer3, Answer4, Answer   'everything in one sometimes easier to read other times not your choice makes no difference
'For Highscores Form and Sub
Public Names(11), Scores(11)
Public Scoresfile
'For About Form
Public Cap
Public AboutFile
'For an extra sub that was not used in the end
Public InFile As Integer
Public PauseTime As Integer
Public Count As Integer
Public TotalTime As Integer
Public Current As Integer
Public X As Integer
Public Start As Single
Public Finish As Single

Public Sub Organize()
  frmMaze.imgChar.Visible = False   'set everything up the way it should start like
  DestPicWidth = 10
  DestPicHeight = 10
  Pos = 19
  Level = 1
End Sub

Sub HighScores()  'here's a slightly more advanced file reading example
  Scoresfile = FreeFile
  Open "highscores.hss" For Input As Scoresfile   'choose the file to read and copy whatever you call it
    
    List = 0  'start from the beginning
      
      While Not EOF(Scoresfile)   'if it isn't finished KEEP GOING
        List = List + 1           'so that it doesn't repeat exactly
        Input #Scoresfile, Names(List), Scores(List)  'put the shit in
      Wend      'while end (end that little thing like end if
    scoreposition = 0     'put the score at the top and work down
      
      For I = 10 To 1 Step -1   'here it is figuring out whether it is good enough to be in the top ten
        If Int(frmMaze.lblScore.Caption) > Scores(I) Then scoreposition = I 'if the score is greater than that of one of the scores then it goes in its place
      Next I
    If scoreposition > 0 Then
        For J = 10 To scoreposition Step -1   'counts down until it is satisfied that you weren't good enough to get on the score table
          Scores(J) = Scores(J - 1)   'makes sure it counts down both name and score
          Names(J) = Names(J - 1)
        Next J
    Scores(scoreposition) = Int(frmMaze.lblScore.Caption)   'saying what is to be pit in if good enough
    Names(scoreposition) = frmMaze.LblName.Caption
  
  Close #Scoresfile   'close file
  
  Scoresfile = FreeFile
  Open "highscores.hss" For Output As Scoresfile    'open the file to edit/change
    
    For I = 1 To 10
      Write #Scoresfile, Names(I), Scores(I)  'write the name and score if good enough in the correct position
    Next I
    
    End If
  Close #Scoresfile     'close file
End Sub
  
Public Sub Maths()
  'Opens the workspace
  Set WS = DBEngine.Workspaces(0)
  'Open the database
  Set DB = WS.OpenDatabase("maths" & Level & ".mdb")
  'Open the record set
  Set RS = DB.OpenRecordset("Questions", dbOpenTable)
  'Set Entry Total
  intTotalRecs = RS.RecordCount
  'Randomizes it all
  Randomize
  'Create Random Integer
  intRndValue = Int((Rnd * 50) + 1)
  'Move to random entry
  RS.Move intRndValue
  With frmQuestions
    .txtFields(0).Text = RS("Question") 'put in the question from the database
    .txtFields(1).Text = RS("Answer1")  'put in the first answer
    .txtFields(2).Text = RS("Answer2")  'put in the second answer
    .txtFields(3).Text = RS("Answer3")  'and so on
    .txtFields(4).Text = RS("Answer4")  'and so forth
    .txtFields(5).Text = RS("Correct")  'etc...
  End With
  'Close record set, DB connection, and Workspace
  RS.Close: DB.Close: WS.Close
  'Destory record set, DB connection, and Workspace
  Set RS = Nothing: Set DB = Nothing: Set WS = Nothing
End Sub

Public Sub QuestionTimer()    'didn't like this it really slowed shit down...good coding i didn't do i though...dunno who did i forgot sorry i didn't find it someone else did and said i should use it...i did edit it a bit so it isn't plagarism...soz i forgot your name dude if you ever find out :)
  Do
    PauseTime = 1  ' Set duration.
    Do While Timer < Start + PauseTime
      DoEvents    ' Yield to other processes.
    Loop
    Finish = Timer  ' Set end time.
    Count = Int(Finish - Start)  ' Calculate total time.
    Current = TotalTime - Count
    frmQuestions.lblTime.Caption = Current
    If Count = 60 Then
      X = MsgBox("You took too long to answer. Sorry!", vbOKOnly)
      If X = vbOK Then
        Unload frmQuestions
      End If
    End If
    DoEvents ' Yield to other processes.
  Loop While Count < 60
End Sub

Public Sub Gameover()
  Organize
  frmGameOver.Visible = True
  frmGameOver.Caption = "Game Over!"
  frmGameOver.Picture = frmGameOver.imgGameover.Picture
  frmGameOver.tmrGameover.Interval = 2000
  frmGameOver.tmrGameover.Enabled = True
End Sub

Public Sub Quit()
  Response = MsgBox("Are you sure you want to exit?", vbYesNo, "Quit?") 'more complicated way of doing a message box
  If Response = vbYes Then    'doing the quit process this way is simpler than doing it each time it needs to be done
    frmGameOver.Visible = True  'simple stuff setting up the single form to do different things. as i use the one form for gameover, goodbye and congratulations
    frmGameOver.Caption = "GoodBye!"
    frmGameOver.Picture = frmGameOver.imgGoodbye.Picture
    frmGameOver.tmrGameover.Interval = 1000
    frmGameOver.tmrGameover.Enabled = True
  ElseIf Response = vbNo Then
    Load frmMenu
  End If
End Sub

Public Sub Congratulations()
  Organize
  frmGameOver.Visible = True
  frmGameOver.Caption = "Congratulations you have finished!"
  frmGameOver.Picture = frmGameOver.imgCongrat.Picture
  frmGameOver.tmrGameover.Interval = 1500
  frmGameOver.tmrGameover.Enabled = True
End Sub

