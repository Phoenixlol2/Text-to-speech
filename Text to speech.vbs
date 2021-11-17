Dim message, sapi
message=InputBox("What do you want me to say?","Text to speech version 0.1")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message
