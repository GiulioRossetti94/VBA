'CODE FOR AUTOMATICALLY CALLIING THE GRAB FUNCTION IN BLOOMBERG AND SENDING IMAGES TO THIS PC'S OUTLOOK ADDRESS
'
'grabMyA__()---> types commands in the bloomberg terminal to grab the performance of the fund
'
'grabSend() ---> routin called by the command button. it grabs and sends a bloomberg screenshot to the users's email

Sub grabMyA__()

GrabFirstPic = Application.Run("BRUNCmd", "FIERITA IM <Equity>GP<GO>GRAB<GO>SGR03<TABR>grab<TABR>PROVA INVIO EMAIL1<TABR>SGR03<GO>1<GO><pause>0002<pause>", 1, "FIERITA IM Equity", "", "")


'r2 = Application.Run("BRUNCmd", "PRTU <GO>16<GO>4<GO><TABR>NESSUNO<TABR> ", 1, "", "", "")
End Sub

Sub grabSend()
'Grab = Application.Run("BRUNCmd", "GRAB<GO>SGR03<TABR>grab<TABR>PROVA INVIO EMAIL1<TABR>SGR03<GO>1<GO>", 1, "", "", "")
Grab = Application.Run("BRUNCmd", "GRAB<GO>SGR03<pause>0002<pause><TABR>grab<TABR>PROVA INVIO EMAIL1<TABR>", 1, "", "", "")
DoEvents
    Application.Wait (Now + TimeValue("0:00:003"))

g1 = Application.Run("BRUNCmd", "1<GO><pause>0002<pause>PORT<GO>", 1, "", "", "")
End Sub