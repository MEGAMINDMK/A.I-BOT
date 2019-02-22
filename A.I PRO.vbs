dim speechobject
set speechobject=createobject("sapi.spvoice")
speechobject.speak "Welcome,         to your system human"
speechobject.speak "I am at your service"
speechobject.speak "My name is"
speechobject.speak "HP  bot"
speechobject.speak "your personal assistant  "
speechobject.speak "my opinion will be"
speechobject.speak "that"
speechobject.speak "you may surf the net"
set Sapi = Wscript.CreateObject("SAPI.SpVoice")
Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
dim str
 if hour(time) < 12 then
 Sapi.speak "Good Morning "
 else
 if hour(time) > 12 then
 if hour(time) > 16 then
 Sapi.speak "Good evening "
 else
 Sapi.speak "Good afternoon "
 end if
 end if
 end if
 Sapi.speak "The current time is"
if hour(time) > 12 then
 Sapi.speak hour(time)-12
 else
 if hour(time) = 0 then
 Sapi.speak "12"
 else
 Sapi.speak hour(time)
 end if
 end if
if minute(time) < 10 then
 Sapi.speak "o"
 if minute(time) < 1 then
 Sapi.speak "clock"
 else
 Sapi.speak minute(time)
 end if
 else
 Sapi.speak minute(time)
 end if
if hour(time) > 12 then
 Sapi.speak "P.M."
 else
 if hour(time) = 0 then
 if minute(time) = 0 then
 Sapi.speak "Midnight"
 else
 Sapi.speak "A.M."
 end if
 else
 if hour(time) = 12 then
 if minute(time) = 0 then
 Sapi.speak "Noon"
 else
 Sapi.speak "P.M."
 end if
 else
 Sapi.speak "A.M."
 end if
 end if
 end if

set wshshell = wscript.CreateObject("wscript.shell")

dim Input

Sapi.speak "I can only gain access to the following."
Sapi.speak "Commands."
Sapi.speak "Youtube, facebook, google, command prompt, calculator, notepad."
Sapi.speak "please input your command here."

do

Sapi.speak "How may I help you?"
Input=inputbox ("A.I PRO BY M.RAAHIM KHATTAK                 Hello! What can I do for you?---->just input the name of app you are trying to open. To exit BOT click on x")

if Input = "youtube" OR Input = "Youtube"then

Sapi.speak "Opening youtube"

wshshell.run "www.youtube.com"

elseif Input = "facebook" OR Input = "Facebook"then

Sapi.speak "Opening facebook"

wshshell.run "www.facebook.com"

elseif Input = "google" OR Input = "Google"then

Sapi.speak "Opening google"

wshshell.run "www.google.com"

elseif Input = "command prompt" OR Input = "Command prompt"then

Sapi.speak "Opening command prompt"

wshshell.run "cmd"

elseif Input = "calculator" OR Input = "Calculator"then

Sapi.speak "Opening calculator"

wshshell.run "calc"

elseif Input = "notepad" OR Input = "Notepad"then

Sapi.speak "Opening notepad"

wshshell.run "notepad"

elseif Input = "" then

Sapi.speak "Goodbye"

else

Sapi.speak "I didn't recognized your input, please try other commands"

end if

loop until Input = ""