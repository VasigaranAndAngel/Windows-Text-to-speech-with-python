import win32com.client
import sys
import random
from datetime import datetime

noargvs = len(sys.argv)

if noargvs > 1:
    text = sys.argv[1]
    texts = text.split("|")
    notexts = len(texts)
    rnd = random.sample(range(1, notexts+1), 1) #rnd = random.randint(1, notexts) #rnd = random.sample(range(1, notexts+1), 1)
    rnd[0] = rnd[0] - 1
    speech = texts[rnd[0]]
    if "GoodMAE" in speech:
        th = datetime.now().strftime("%H")
        if int(th) < 12:
            speech = speech.replace("GoodMAE","Good-morning")
        if int(th) > 12 and int(th) < 15:
            speech = speech.replace("GoodMAE","Good-afternoon")
        if int(th) > 15:
            speech = speech.replace("GoodMAE","Good-evening")
    if "TimeH" in speech:
        th = datetime.now().strftime("%H")
        if int(th) > 12:
            th = int(th) - 12
        speech = speech.replace("TimeH",str(th))
    if "TimeM" in speech:
        th = datetime.now().strftime("%M")
        speech = speech.replace("TimeM",str(th))
    if "TimeS" in speech:
        th = datetime.now().strftime("%S")
        speech = speech.replace("TimeS",str(th))
    if "TimeAP" in speech:
        th = datetime.now().strftime("%H")
        if int(th) > 12:
            th = "PM"
        else:
            th = "AM"
        speech = speech.replace("TimeAP",str(th))

    if "DateD" in speech:
        date = datetime.now().strftime("%D")
        speech = speech.replace("DateD",str(date))

else:
    speech = "Unknown speak call."

if noargvs > 2:
    setvoice = sys.argv[2]
else:
    setvoice = 0

if noargvs > 3:
    voicerate = sys.argv[3]
else:
    voicerate = 0

if noargvs > 4:
    voicevolume = sys.argv[4]
else:
    voicevolume = 100


voice = win32com.client.Dispatch("SAPI.SPVoice")

#print(voice.getvoices().item(1))
#voice.voice = voice.getvoices.item(setvoice)

voice.rate = voicerate
voice.volume = voicevolume
voice.speak (speech)