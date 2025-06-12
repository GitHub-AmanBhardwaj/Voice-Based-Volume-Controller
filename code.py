import speech_recognition as sr
import re
import pythoncom
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import keyboard

recognizer=sr.Recognizer()
pythoncom.CoInitialize()
devices=AudioUtilities.GetSpeakers()
interface=devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume=cast(interface, POINTER(IAudioEndpointVolume))

with sr.Microphone() as m:
    # recognizer.adjust_for_ambient_noise(m,duration=2)   ## Use only if you want to identify background noises.
    while True:
        if keyboard.is_pressed('q'):
            print('\nEnded X_X')
            break
        try:
            print('Listening :)')
            audio=recognizer.listen(m,timeout=2,phrase_time_limit=4)
            text=recognizer.recognize_google(audio).lower()
            if 'volume up' in text:
                volume.SetMasterVolumeLevelScalar(1.0, None)
                print(f'Command: "{text}" Done :)')
            elif 'volume down' in text:
                volume.SetMasterVolumeLevelScalar(0.1, None)
                print(f'Command: "{text}" Done :)')
            elif 'set volume to ' in text:
                pattern=r'set volume to (\d{1,2})'
                match=re.search(pattern, text)
                if match:
                    n=int(match.group(1))
                    volume.SetMasterVolumeLevelScalar(n/100, None)
                    print(f'Command: "{text}" Done :)')
            else:
                print('No Command!')
        except:
            print('Nothing Heard!')
        
pythoncom.CoUninitialize()
