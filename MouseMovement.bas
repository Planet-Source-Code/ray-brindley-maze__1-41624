Attribute VB_Name = "MouseMovement"
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
'ok this i didn't do the person i made this for made it
'like this only in the form which just made it difficult
'to find what you wanted to edit...it a bit easier this way

Public Sub Char1_MouseMove()
  With frmMenu      'so that  you don't have to write frmmenu over and over and over again
  If .imgCharO2.Visible = True Then   'this basically says that if the mouse is moved
    .imgChar2.Visible = False         'over this picture then things are visable or not
    .imgCharI2.Visible = False        'depending on there original state
    .imgCharO2.Visible = True
    .imgChar1.Visible = False
    .imgCharI1.Visible = True
    .imgCharO1.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  ElseIf .imgCharO3.Visible = True Then
    .imgChar3.Visible = False
    .imgCharI3.Visible = False
    .imgCharO3.Visible = True
    .imgChar1.Visible = False
    .imgCharI1.Visible = True
    .imgCharO1.Visible = False
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
  Else
    .imgChar1.Visible = False
    .imgCharI1.Visible = True
    .imgCharO1.Visible = False
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  End If
  End With
End Sub

Public Sub Char2_MouseMove()
  With frmMenu
  If .imgCharO1.Visible = True Then     'same as above. i sure there would be another
    .imgChar1.Visible = False           'way to do this stuff but i didn't bother...it works
    .imgCharI1.Visible = False
    .imgCharO1.Visible = True
    .imgChar2.Visible = False
    .imgCharI2.Visible = True
    .imgCharO2.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  ElseIf .imgCharO3.Visible = True Then
    .imgChar3.Visible = False
    .imgCharI3.Visible = False
    .imgCharO3.Visible = True
    .imgChar2.Visible = False
    .imgCharI2.Visible = True
    .imgCharO2.Visible = False
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
  Else
    .imgChar2.Visible = False
    .imgCharI2.Visible = True
    .imgCharO2.Visible = False
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  End If
  End With
End Sub

Public Sub Char3_MouseMove()
  With frmMenu
  If .imgCharO1.Visible = True Then   'same again
    .imgChar1.Visible = False
    .imgCharI1.Visible = False
    .imgCharO1.Visible = True
    .imgChar3.Visible = False
    .imgCharI3.Visible = True
    .imgCharO3.Visible = False
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
  ElseIf .imgCharO2.Visible = True Then
    .imgChar2.Visible = False
    .imgCharI2.Visible = False
    .imgCharO2.Visible = True
    .imgChar3.Visible = False
    .imgCharI3.Visible = True
    .imgCharO3.Visible = False
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
  Else
    .imgChar3.Visible = False
    .imgCharI3.Visible = True
    .imgCharO3.Visible = False
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
  End If
  End With
End Sub

Public Sub CharI1_MouseDown()
  With frmMenu
  .imgChar1.Visible = False       'if you click on a certain image then another image
  .imgCharI1.Visible = False      'is visable. like the previous ones only clicking
  .imgCharO1.Visible = True       'this time
  .imgChar2.Visible = True
  .imgCharI2.Visible = False
  .imgCharO2.Visible = False
  .imgChar3.Visible = True
  .imgCharI3.Visible = False
  .imgCharO3.Visible = False
  End With
  frmMaze.imgPucca.Picture = frmMenu.imgChar1.Picture   'make everything change for the
  frmMaze.imgChar.Picture = frmMaze.imgBoy              'actual maze. all the right
  frmMaze.imgTop.Picture = frmMaze.imgGaTo.Picture      'colors in other words
  frmMaze.imgLeft.Picture = frmMaze.imgGaSi.Picture
  frmMaze.imgRight.Picture = frmMaze.imgGaSi.Picture
  frmMaze.imgBottom.Picture = frmMaze.imgGaBo.Picture
  frmMaze.pic.Picture = frmMaze.picGaru.Picture
  'frmMaze.lblFinish.BackColor = RGB(153, 255, 153)
  'frmMaze.lblStart.BackColor = RGB(153, 255, 153)
End Sub

Public Sub CharI2_MouseDown()
  With frmMenu
  .imgChar2.Visible = False       'same again
  .imgCharI2.Visible = False
  .imgCharO2.Visible = True
  .imgChar1.Visible = True
  .imgCharI1.Visible = False
  .imgCharO1.Visible = False
  .imgChar3.Visible = True
  .imgCharI3.Visible = False
  .imgCharO3.Visible = False
  End With
  frmMaze.imgPucca.Picture = frmMenu.imgChar2.Picture
  frmMaze.imgChar.Picture = frmMaze.imgRabbit.Picture
  frmMaze.imgTop.Picture = frmMaze.imgMaTo.Picture
  frmMaze.imgLeft.Picture = frmMaze.imgMaSi.Picture
  frmMaze.imgRight.Picture = frmMaze.imgMaSi.Picture
  frmMaze.imgBottom.Picture = frmMaze.imgMaBo.Picture
  frmMaze.pic.Picture = frmMaze.picMisha.Picture
  'frmMaze.lblFinish.BackColor = RGB(255, 255, 153)
  'frmMaze.lblStart.BackColor = RGB(255, 255, 153)
End Sub

Public Sub CharI3_MouseDown()
  With frmMenu
  .imgChar3.Visible = False   'and again
  .imgCharI3.Visible = False
  .imgCharO3.Visible = True
  .imgChar1.Visible = True
  .imgCharI1.Visible = False
  .imgCharO1.Visible = False
  .imgChar2.Visible = True
  .imgCharI2.Visible = False
  .imgCharO2.Visible = False
  End With
  frmMaze.imgPucca.Picture = frmMenu.imgChar3.Picture
  frmMaze.imgChar.Picture = frmMaze.imgGirl.Picture
  frmMaze.imgTop.Picture = frmMaze.imgPuTo.Picture
  frmMaze.imgLeft.Picture = frmMaze.imgPuSi.Picture
  frmMaze.imgRight.Picture = frmMaze.imgPuSi.Picture
  frmMaze.imgBottom.Picture = frmMaze.imgPuBo.Picture
  frmMaze.pic.Picture = frmMaze.picPucca.Picture
  'frmMaze.lblFinish.BackColor = RGB(255, 192, 203)
  'frmMaze.lblStart.BackColor = RGB(255, 192, 203)
End Sub

Public Sub GettingStarted()   'this makes everything start up visible and everything
  With frmMenu                'so it all look the way it should, very round about way of doing it but it works
  If .imgCharO1.Visible = True And .imgChar2.Visible = True And .imgChar3.Visible = True Then
    .imgChar1.Visible = False
    .imgCharI1.Visible = False
    .imgCharO1.Visible = True
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  ElseIf .imgCharO2.Visible = True And .imgChar1.Visible = True And .imgChar3.Visible = True Then
    .imgChar2.Visible = False
    .imgCharI2.Visible = False
    .imgCharO2.Visible = True
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  ElseIf .imgCharO3.Visible = True And .imgChar1.Visible = True And .imgChar2.Visible = True Then
    .imgChar3.Visible = False
    .imgCharI3.Visible = False
    .imgCharO3.Visible = True
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
  ElseIf (.imgCharO1.Visible = True And .imgCharI2.Visible = True And .imgChar3.Visible = True) Or (.imgCharO1.Visible = True And .imgCharI3.Visible = True And .imgChar2.Visible = True) Then
    .imgChar1.Visible = False
    .imgCharI1.Visible = False
    .imgCharO1.Visible = True
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  ElseIf (.imgCharO2.Visible = True And .imgCharI1.Visible = True And .imgChar3.Visible = True) Or (.imgCharO2.Visible = True And .imgCharI3.Visible = True And .imgChar1.Visible = True) Then
    .imgChar2.Visible = False
    .imgCharI2.Visible = False
    .imgCharO2.Visible = True
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  ElseIf (.imgCharO3.Visible = True And .imgCharI1.Visible = True And .imgChar2.Visible = True) Or (.imgCharO3.Visible = True And .imgCharI2.Visible = True And .imgChar1.Visible = True) Then
    .imgChar3.Visible = False
    .imgCharI3.Visible = False
    .imgCharO3.Visible = True
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
  Else
    .imgChar1.Visible = True
    .imgCharI1.Visible = False
    .imgCharO1.Visible = False
    .imgChar2.Visible = True
    .imgCharI2.Visible = False
    .imgCharO2.Visible = False
    .imgChar3.Visible = True
    .imgCharI3.Visible = False
    .imgCharO3.Visible = False
  End If
  End With
End Sub
