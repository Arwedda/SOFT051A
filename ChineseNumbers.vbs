Option Explicit 
Dim score 
Dim currentRnd
Dim previousRnd
Dim btnCollection
Dim starCollection
Dim englishNumbersCollection
Dim i
Dim btnPressed
Dim inputInt
Dim answerSubmitted

Sub window_onLoad()
    Randomize
    currentRnd = 0
    score = 0
    answerSubmitted = true 
    btnCollection = Array(btnOne, btnTwo, btnThree, btnFour, btnFive, btnSix, btnSeven, btnEight, btnNine, btnTen)
    starCollection = Array(starOne, starTwo, starThree, starFour, starFive)
End Sub

Sub btnPlay_onClick()
    If (answerSubmitted = true) Then 
        previousRnd = currentRnd    
        Do While previousRnd = currentRnd
            currentRnd = Int(Rnd() * 10) + 1 
        Loop 
        For i = 0 To 9 
            btnCollection(i).disabled = false
            btnCollection(i).style.backgroundColor = ""
        Next
        parFeedback.style.visibility = "hidden"
        answerSubmitted = false
    End If
    sndPlayer.src = "ChineseNumber" + CStr(currentRnd) + ".wav"
End Sub 

Sub btnOne_onClick()
    btnHandler(1)
End Sub 

Sub btnTwo_onClick()
    btnHandler(2)
End Sub 

Sub btnThree_onClick
    btnHandler(3)
End Sub 

Sub btnFour_onClick
    btnHandler(4)
End Sub 

Sub btnFive_onClick()
    btnHandler(5)
End Sub 

Sub btnSix_onClick()
    btnHandler(6)
End Sub 

Sub btnSeven_onClick()
    btnHandler(7)
End Sub 

Sub btnEight_onClick()
    btnHandler(8)
End Sub 

Sub btnNine_onClick()
    btnHandler(9)
End Sub 

Sub btnTen_onClick()
    btnHandler(10)
End Sub 

Sub btnHandler(btnPressed)
    answerSubmitted = true
    sndPlayer.src = "ChineseNumber" + CStr(btnPressed) + ".wav"
    If (currentRnd = btnPressed) Then 
        btnCollection(btnPressed - 1).style.backgroundColor = "Green"
        parFeedback.innerText = "Correct"
        parFeedback.style.backgroundColor = "Yellow"
        score = score + 1
        starCollection(score - 1).src = "StarOn.gif"
        If (score = 5) Then 
            starFive.src = "StarOn.gif"
            MsgBox "Congratulations! You have won the game."
            window.location.reload()
        End If 
    Else
        btnCollection(btnPressed - 1).style.backgroundColor = "Red"
        parFeedback.innerText = "Sorry, it was " + CStr(currentRnd)
        parFeedback.style.backgroundColor = "Red"
        btnCollection(currentRnd - 1).style.backgroundColor = "Green"
    End If 
    parFeedback.style.visibility = "visible"
    For i = 0 To 9
        btnCollection(i).disabled = true
    Next
End Sub
