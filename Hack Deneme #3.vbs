Do
StrText=("Ben Kaan Kayserili")
set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Rate=-3
ObjVoice.Speak StrText
Loop