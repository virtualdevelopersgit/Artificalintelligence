Dim sapi
set sapi = CreateObject("sapi.spvoice")
Set sapi.Voice = sapi.GetVoices.Item(1)
sapi.Speak "Hello sir, I am FRIDAY your Personal Assistant"
sapi.Speak "Starting Systems"
sapi.Speak "Checking Networks"
sapi.Speak "Establishing Network to the main Server"
sapi.Speak "Checking System Files"
sapi.Speak "Scanning Complete"
sapi.Speak "We are Ready to go sir."

