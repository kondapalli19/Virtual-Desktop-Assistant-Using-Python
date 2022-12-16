Table of Contents
-------------------

* Introduction
* Features
* Installation
* Troubleshooting
* References


Introduction
--------------------

Virtual Desktop Assistant - Jarvis, interactive Assistant who takes input (voice), process and gives output as result.


Features
--------------------

- Opens Google, Youtube, vs code on taking command by micrphone.

- Reads the wikipedia content.

- Telling the current time, checking the weather of the inputted city, read the healines of the today's news and also provides the link.

- play music, play game.

- Tell the schedule (timetable) provided.

- stops after getting shutdown command.


Installation
----------------------

Software Requirements:

        * Python version : Python 3.10.8

        * vs code version : 1.74.0

        * table.txt : File provides the time table

        * Installing python modules : 

                pip install <module_name>

                module_name
                - pyttsx3
                - psutil
                - wikipedia
                - SpeechRecognition
                - requests
                - pywin32
                - bs4

        * Get API from newsapi.org to get the today's news daily.

        * In play music, music_dir variable give the path of the directory which contains only music file.


Troubleshooting
-----------------------

Check microphone of your system is properly working or not fix this issue settings > sound settings > choose your input device.
After that check if Google assistant is taking your voice as command or else fix this issue first.


References
----------------------


* If pip is not working properly.

        https://stackoverflow.com/questions/23708898/pip-is-not-recognized-as-an-internal-or-external-command

        If that doesn't work :
                For Windows, when you install a package, you type:

                        python -m pip install [packagename]


* If speak function does not working properly.

        Check if it's imported or not.

                from win32com.client import Dispatch
                speak=Dispatch("SAPI.spVoice")
                speak.Speak(string)

                or

                from win32com.client import Dispatch

        
        



How to get API for free?


If SpeechRecognition is not working.


If any other problem occured
contact me through email - kkoyal19599@gamil.com




