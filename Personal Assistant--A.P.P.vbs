Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
set wshshell = wscript.CreateObject("wscript.shell")

dim input


Sapi.speak "Hello! My name is Advanced Analysis Prototype or APP.  What shall I do for you today?"

do
Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
set wshshell = wscript.CreateObject("wscript.shell")



input=inputbox ("Hello my name is Advanced Analysis Prototype or A.A.P.  How may I serve you today!")


if input = "youtube" or input = "Youtube"then
Sapi.speak "Opening youtube"
wshshell.run "www.youtube.com"

else
if input = "github" or input = "Github"then
Sapi.speak "Here you go sir"
wshshell.run "https://github.com/"

else
if input = input = "Git hub" or input = "git Hub" or input = "Git Hub"then
Sapi.speak "Opening Github"
wshshell.run "https://github.com/"

else
if input = input = "GitHub" or input = "git hub"then
Sapi.speak "Here is your code"
wshshell.run "https://github.com/"

else
if input = "gmail" or input = "Gmail"then
Sapi.speak "Opening G.mail"
wshshell.run "www.gmail.com"

else
if input = "internet explorer" or input = "Internet explorer" or input = "internet Explorer" or input = "Internet Explorer"then
Sapi.speak "Opening internet explorer"
wshshell.run "iexplore.exe"


else
if input = "twitter" or input = "Twitter"then
Sapi.speak "Opening twitter"
wshshell.run "www.twitter.com"

else
if input = "what is the weather" or input = "What is the weather" or input = "what is the Weather" or input = "weather" or input = "Weather"then
Sapi.speak "Here is the weather"
wshshell.run "www.theweathernetwork.com"


else
if input = "facebook" or input = "Facebook"then
Sapi.speak "Opening facebook"
wshshell.run "www.facebook.com"

else
if input = "my name is Siddharth" or input = "My name is Siddharth"then
Sapi.speak "Hello Sire Siddharth"
input=inputbox ("I hope I am serving you well")

else
if input = "yahoo" or input = "Yahoo"then
Sapi.speak "Opening yahoo"
wshshell.run "www.yahoo.com"

else
if input = "waz up" or input = "Waz up" or input = "what's up" or input = "What's up" or input = "whats up" or input = "Whats up" or input = "what up" or input = "What up"then
Sapi.speak "The sky obviously HAHAHAHAHAHAHAHA"
input=inputbox ("HAHAHAHAHAHA the anwser is so simple (The anwser is the sky)!")


else
if input = "notepad" or input = "Notepad"then
Sapi.speak "Opening notepad"
wshshell.run "notepad.exe"

else
if input = "pinterest" or input = "Pinterest"then
Sapi.speak "Opening pinterest"
wshshell.run "www.pinterest.com"

else
if input = "bing" or input = "Bing"then
Sapi.speak "Opening bing"
wshshell.run "www.bing.com"

else
if input = "google drive" or input = "Google drive" or input = "Google Drive" or input = "google Drive"then
Sapi.speak "Opening google drive"
wshshell.run "www.drive.google.com"

else
if input = "hello" or input = "Hello" then
Sapi.speak "Hello Sire"
input=inputbox ("Hello Master how are you today!")

else
if input = "how are you" or input = "How are you" then
Sapi.speak "I am fine, thank you for asking"
input=inputbox ("Thank you for asking! I am fine, how are you!")

else
if input = "wikipedia" or input = "Wikipedia" or input = "wiki" or input = "Wiki"then
Sapi.speak "Opening wikipedia"
wshshell.run "www.wikipedia.org"

else
if input = "google" or input = "Google" then
Sapi.speak "Opening google"
wshshell.run "www.google.com"

else
if input = "command prompt" or input = "Command prompt" or input = "cmd" then
Sapi.speak "Opening command prompt"
wshshell.run "cmd"

else
if input = "open calculator" or input = "open Calculator" or input = "Open calculator" or input = "Open Calculator" then
Sapi.speak "Opening calculator"
wshshell.run "calc"

else
if input = "notepad" or input = "Notepad" then
Sapi.speak "Opening notepad"
wshshell.run "notepad"

else
if input = "date" or input = "Date" or input = "what is the date" or input = "What is the date" or input = "what is the Date" or input = "What is the Date" then
Sapi.speak "Here is todays date!"
msgbox (Date)

else
if input = "time" or input = "Time" or input = "what is the time" or input = "What is the time" or input = "what is the Time" or input = "What is the Time"then
Sapi.speak "Here is the time!"
msgbox (Time)

else
if input = "" then
else


Sapi.speak "I don't recognize your input, please try again"
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
loop