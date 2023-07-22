' maxbybee
' 2020/7
set shell = createobject ("wscript.shell")
strtimes = inputbox("Amount of messages to spam")
strtdelay = inputbox("Delay between messages")
if not isnumeric(strtimes) then
wscript.quit
end if
if not isnumeric(strtdelay) then
wscript.quit
' makes sure imput values are valid
end if
msgbox "You have 5 seconds to get to your inputbox."
wscript.sleep( 5000 )
for i=1 to strtimes
shell.sendkeys("^V")
shell.sendkeys("~")
' after pasting the clipboard, sends an enter key
wscript.sleep(strdelay)
next
