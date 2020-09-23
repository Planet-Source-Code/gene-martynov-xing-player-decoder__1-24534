Hi everybody!

Would not it be nice to have something that can play and/or decode file by using one control? So here we go. Xaudio.dll can do that.
I want to say that I did not write the DLL itself. Thanks to the guys who wrote it. I just made this DLL work within Visual Basic. 
So you will need the xaudio.dll and xanalyze.dll for this project. Place them into your Windows/System directory.
Some explanation on using xaudio and xanalyzer dll's included as HTML files.


Look at my codes, they are pretty much commented.
Important part of it is:
1. frmMain_Load and Unload events
2. Module modSubs - which is processing all the messages
3. Module modDecoder - which is declaring some variables and functions for decoder.
4. Module modAnalyzer - which is taking care of file data (like bit rate, sampling frequency, mode


As usually, when working with subclassing, running the program step by step (by pressing F8) is not safe. You can use Debug.Print instead.
