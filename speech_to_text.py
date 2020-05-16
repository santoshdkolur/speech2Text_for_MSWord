
# ### Dependencies to be installed:
#         - pyautogui            ( pip install pyautogui )
#         - speech_recognition   ( pip install SpeechRecognition )
#         - win32gui             ( pip install win32gui )
#         - pyaudio              ( pip install pyaudio )
#         
# The below code is written for Windows and will not work on other platforms. 
# Changes to lines using procedures from the library win32gui have to be made in accordance to your machine.
# 
# ##### Note : Requires active internet connect to work and might have lags. I will try to update this with a better version when I get time.


import pyautogui as p
import speech_recognition as sr
import os
from time import sleep
from win32gui import GetWindowText, GetForegroundWindow

def SpeechToText(r):
    
    #The code can be used to enter data into any text field by making the appropriate changes.
    #Refer the readme.md file on github to know what lines to change.
    
    flag=0
    p.alert(text='Are you ready?', title='START', button='Yes')
    
    while(True):
        path=input("Enter the path to your word file. Ex C:\\Users\\Meow\\Documents\\test.docx \nType Here:  ")
        path='\\\\'.join(path.split('\\'))
        try:
            os.startfile("meow")
        except:
            print("File not found or invalid path!\nTry again: ")
            if(flag<=2):
                flag+=1
                continue
            else:
                print('Too many attempts. Please check the path you\'ve been entering. \n\n')
                exit()
        break
        
                
    count=1                                                            #Keeping the count to switch to the tab where the Word is open.
    while('Word' not in GetWindowText(GetForegroundWindow())[-4:]):    #The letters 'Word' might differ on your machine. This is for windows.
        p.keyDown('altleft')
        count=count+1
        for i in range(1,count+1):
            p.press('tab')
        sleep(0.3)
        p.keyUp('altleft')
    p.write('Speak Now: ')                                             
    inp=''
    while(inp!= 'Cancel'):
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source, duration=0.2)
            #Listen 
            audio = r.listen(source)
            # Recognise.   Needs active internet connection
            try:
                Text = r.recognize_google(audio)

                for char in Text:
                    if(char.isupper()):
                        p.keyDown('shift')
                        p.press(char)
                        p.keyUp('shift')
                    else:
                        p.press(str(char))
            
            except sr.UnknownValueError:
                p.alert(text="Google Speech Recognition could not understand audio",title="Not Recognized!",button="Okay")
                
        inp=p.confirm(text='Would you like to continue? ', title='CONFIRM', buttons=['OK', 'Cancel'])


def main():
    r = sr.Recognizer()
    SpeechToText(r)


if __name__ == '__main__':
	main()

