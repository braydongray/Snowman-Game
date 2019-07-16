'Snowman Game
'Program is a snowman remake of the hangman game.
'Written by Braydon Gray Novemeber 22, 2018
'Modified by Braydon Gray December 12, 2018

Option Explicit On
Option Infer Off

Public Class frmSnowman
    'declare variables
    Private intMistakes As Integer
    Private strWordList(12) As String
    Private strLetter As String = ""
    Private strUsername As String = frmWelcome.strUsername
    'guess letter procedure
    Private Sub GuessLetter()
        'declare boolean to count mistakes
        Dim blnCorrect As Boolean = False
        'if the letter guessed is equal to the character location than display it 
        If lblOne.Text = strLetter Then
            lblOne.Visible = True
            blnCorrect = True
        End If
        If lblTwo.Text = strLetter Then
            lblTwo.Visible = True
            blnCorrect = True
        End If
        If lblThree.Text = strLetter Then
            lblThree.Visible = True
            blnCorrect = True
        End If
        If lblFour.Text = strLetter Then
            lblFour.Visible = True
            blnCorrect = True
        End If
        If lblFive.Text = strLetter Then
            lblFive.Visible = True
            blnCorrect = True
        End If
        If lblSix.Text = strLetter Then
            lblSix.Visible = True
            blnCorrect = True
        End If
        If lblSeven.Text = strLetter Then
            lblSeven.Visible = True
            blnCorrect = True
        End If
        If lblEight.Text = strLetter Then
            lblEight.Visible = True
            blnCorrect = True
        End If
        If lblNine.Text = strLetter Then
            lblNine.Visible = True
            blnCorrect = True
        End If
        If lblTen.Text = strLetter Then
            lblTen.Visible = True
            blnCorrect = True
        End If
        If lblEleven.Text = strLetter Then
            lblEleven.Visible = True
            blnCorrect = True
        End If
        If lblTwelve.Text = strLetter Then
            lblTwelve.Visible = True
            blnCorrect = True
        End If
        'count mistakes
        If blnCorrect = False Then
            intMistakes = intMistakes + 1
        End If

    End Sub
    Private Sub CreateWordList()
        strWordList(0) = "DETAIL      "
        strWordList(1) = "SNOWMAN     "
        strWordList(2) = "CARROT      "
        strWordList(3) = "CANVAS      "
        strWordList(4) = "OBJECT      "
        strWordList(5) = "VARIABLE    "
        strWordList(6) = "ARRAY       "
        strWordList(7) = "VISUAL      "
        strWordList(8) = "STUDIO      "
        strWordList(9) = "WINDOWS     "
        strWordList(10) = "CHROME      "
        strWordList(11) = "WEBSITE     "
        strWordList(12) = "INTERNET    "
    End Sub
    Private Sub NewGame()
        CreateWordList()
        Dim strWord As String = strWordList(Int(13 * Rnd()))
        Dim chrWord() As Char = strWord.ToCharArray
        lblOne.Text = chrWord(0)
        lblTwo.Text = chrWord(1)
        lblThree.Text = chrWord(2)
        lblFour.Text = chrWord(3)
        lblFive.Text = chrWord(4)
        lblSix.Text = chrWord(5)
        lblSeven.Text = chrWord(6)
        lblEight.Text = chrWord(7)
        lblNine.Text = chrWord(8)
        lblTen.Text = chrWord(9)
        lblEleven.Text = chrWord(10)
        lblTwelve.Text = chrWord(11)

        lblOne.Visible = False
        lblTwo.Visible = False
        lblThree.Visible = False
        lblFour.Visible = False
        lblFive.Visible = False
        lblSix.Visible = False
        lblSeven.Visible = False
        lblEight.Visible = False
        lblNine.Visible = False
        lblTen.Visible = False
        lblEleven.Visible = False
        lblTwelve.Visible = False

        If chrWord(0) = " " Then
            linOne.Visible = False
            lblOne.Visible = True
        Else
            linOne.Visible = True
        End If
        If chrWord(1) = " " Then
            linTwo.Visible = False
            lblTwo.Visible = True
        Else
            linTwo.Visible = True
        End If
        If chrWord(2) = " " Then
            linThree.Visible = False
            lblThree.Visible = True
        Else
            linThree.Visible = True
        End If
        If chrWord(3) = " " Then
            linFour.Visible = False
            lblFour.Visible = True
        Else
            linFour.Visible = True
        End If
        If chrWord(4) = " " Then
            linFive.Visible = False
            lblFive.Visible = True
        Else
            linFive.Visible = True
        End If
        If chrWord(5) = " " Then
            linSix.Visible = False
            lblSix.Visible = True
        Else
            linSix.Visible = True
        End If
        If chrWord(6) = " " Then
            linSeven.Visible = False
            lblSeven.Visible = True
        Else
            linSeven.Visible = True
        End If
        If chrWord(7) = " " Then
            linEight.Visible = False
            lblEight.Visible = True
        Else
            linEight.Visible = True
        End If
        If chrWord(8) = " " Then
            linNine.Visible = False
            lblNine.Visible = True
        Else
            linNine.Visible = True
        End If
        If chrWord(9) = " " Then
            linTen.Visible = False
            lblTen.Visible = True
        Else
            linTen.Visible = True
        End If
        If chrWord(10) = " " Then
            linEleven.Visible = False
            lblEleven.Visible = True
        Else
            linEleven.Visible = True
        End If
        If chrWord(11) = " " Then
            linTwelve.Visible = False
            lblTwelve.Visible = True
        Else
            linTwelve.Visible = True
        End If
        picHat.Visible = True
        picEyeOne.Visible = True
        picEyeTwo.Visible = True
        picNoseOne.Visible = True
        picNoseTwo.Visible = True
        picMouthOne.Visible = True
        picMouthTwo.Visible = True
        picMouthThree.Visible = True
        picMouthFour.Visible = True
        picMouthFive.Visible = True
        picHead.Visible = True
        picArmOne.Visible = True
        picArmTwo.Visible = True
        picTorso.Visible = True
        picBottom.Visible = True
        picPuddle.Visible = False
        intMistakes = 0
        lblMessage.Text = ""
        cmdA.Enabled = True
        cmdB.Enabled = True
        cmdC.Enabled = True
        cmdD.Enabled = True
        cmdE.Enabled = True
        cmdF.Enabled = True
        cmdG.Enabled = True
        cmdH.Enabled = True
        cmdI.Enabled = True
        cmdJ.Enabled = True
        cmdK.Enabled = True
        cmdL.Enabled = True
        cmdM.Enabled = True
        cmdN.Enabled = True
        cmdO.Enabled = True
        cmdP.Enabled = True
        cmdQ.Enabled = True
        cmdR.Enabled = True
        cmdS.Enabled = True
        cmdT.Enabled = True
        cmdU.Enabled = True
        cmdV.Enabled = True
        cmdW.Enabled = True
        cmdX.Enabled = True
        cmdY.Enabled = True
        cmdZ.Enabled = True
    End Sub
    Private Sub cmdNewGame_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNewGame.Click
        NewGame()
    End Sub
    Private Sub cmdLetter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdA.Click, cmdB.Click, cmdC.Click, cmdD.Click, cmdE.Click, cmdF.Click, cmdG.Click, cmdH.Click, cmdI.Click, cmdJ.Click, cmdK.Click, cmdL.Click, cmdM.Click, cmdN.Click, cmdO.Click, cmdP.Click, cmdQ.Click, cmdR.Click, cmdS.Click, cmdT.Click, cmdU.Click, cmdV.Click, cmdW.Click, cmdX.Click, cmdY.Click, cmdZ.Click
        Dim cmdLetterClicked As Button = sender
        strLetter = cmdLetterClicked.Text
        cmdLetterClicked.Enabled = False
        GuessLetter()
        If lblOne.Visible = True And lblTwo.Visible = True And lblThree.Visible = True And lblFour.Visible = True And lblFive.Visible = True And lblSix.Visible = True And lblSeven.Visible = True And lblEight.Visible = True And lblNine.Visible = True And lblTen.Visible = True And lblEleven.Visible = True And lblTwelve.Visible = True Then
            lblMessage.Text = strUsername & " has won!"
            cmdA.Enabled = False
            cmdB.Enabled = False
            cmdC.Enabled = False
            cmdD.Enabled = False
            cmdE.Enabled = False
            cmdF.Enabled = False
            cmdG.Enabled = False
            cmdH.Enabled = False
            cmdI.Enabled = False
            cmdJ.Enabled = False
            cmdK.Enabled = False
            cmdL.Enabled = False
            cmdM.Enabled = False
            cmdN.Enabled = False
            cmdO.Enabled = False
            cmdP.Enabled = False
            cmdQ.Enabled = False
            cmdR.Enabled = False
            cmdS.Enabled = False
            cmdT.Enabled = False
            cmdU.Enabled = False
            cmdV.Enabled = False
            cmdW.Enabled = False
            cmdX.Enabled = False
            cmdY.Enabled = False
            cmdZ.Enabled = False
        End If
        If intMistakes = 1 Then
            picHat.Visible = False
        ElseIf intMistakes = 2 Then
            picNoseOne.Visible = False
            picNoseTwo.Visible = False
            picEyeOne.Visible = False
            picEyeTwo.Visible = False
            picMouthOne.Visible = False
            picMouthTwo.Visible = False
            picMouthThree.Visible = False
            picMouthFour.Visible = False
            picMouthFive.Visible = False
        ElseIf intMistakes = 3 Then
            picHead.Visible = False
        ElseIf intMistakes = 4 Then
            picArmOne.Visible = False
            picArmTwo.Visible = False
        ElseIf intMistakes = 5 Then
            picTorso.Visible = False
        ElseIf intMistakes = 6 Then
            picBottom.Visible = False
            picPuddle.Visible = True
            lblMessage.Text = "You melted!"
            lblOne.Visible = True
            lblTwo.Visible = True
            lblThree.Visible = True
            lblFour.Visible = True
            lblFive.Visible = True
            lblSix.Visible = True
            lblSeven.Visible = True
            lblEight.Visible = True
            lblNine.Visible = True
            lblTen.Visible = True
            lblEleven.Visible = True
            lblTwelve.Visible = True
            cmdA.Enabled = False
            cmdB.Enabled = False
            cmdC.Enabled = False
            cmdD.Enabled = False
            cmdE.Enabled = False
            cmdF.Enabled = False
            cmdG.Enabled = False
            cmdH.Enabled = False
            cmdI.Enabled = False
            cmdJ.Enabled = False
            cmdK.Enabled = False
            cmdL.Enabled = False
            cmdM.Enabled = False
            cmdN.Enabled = False
            cmdO.Enabled = False
            cmdP.Enabled = False
            cmdQ.Enabled = False
            cmdR.Enabled = False
            cmdS.Enabled = False
            cmdT.Enabled = False
            cmdU.Enabled = False
            cmdV.Enabled = False
            cmdW.Enabled = False
            cmdX.Enabled = False
            cmdY.Enabled = False
            cmdZ.Enabled = False
        End If
    End Sub
    Private Sub frmSnowMan_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        NewGame()
    End Sub

    Private Sub filNewUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles filNewUser.Click
        frmWelcome.Show()
        Me.Close()
    End Sub

    Private Sub filNewGame_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles filNewGame.Click
        NewGame()
    End Sub
End Class
