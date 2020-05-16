**Random Project No. 1**

Disclaimer!
The script is not bugfree and is not very reliable as of now. This is the first template version. I'll be improving it whenever I'm free.\
This is a personal project of mine and had written it during the lockdown due to corvid-19.
It can be used to have speech based entry to Microsoft Word (or any other editor or text field with simple modifications).

It uses SpeechRecognition library in python and pyautogui for automatic entry of the recognized text into Word.
You can find the link to SpeechRecognition library here : https://github.com/Uberi/speech_recognition#readme 

To use it another application change the string 'Word' in the while condition to the name of your application. Remember it is case sensitive.
This parses phrase by phrase, hence you'll need to select if you want to continue after every phrase.

You can also enable it to listen in the background and have your main thread of your python perform other computation.
Check out the official documentation for a better understanding.

TO RUN THE FILE:
You can run the .py file normally as a python file i.e. Open cmd at the file location and type: 
*python speech_to_text.py*

To run on the notebook file, open jupyter notebook or any notebook editor and run all the cells.
