# PSWordleSolver
Powershell script that uses selenium and chrome driver to auto solve Wordle

Paths and locations may need to be updated.  The script automatically goes to the wordle page and attempts to solve

disclaimer... my code is a bit messy.  i did this for fun, mainly to see if i could figure it out!

make sure you install the latest selenium module

Install-Module selenium

download chrome driver that matches your browser's version

https://chromedriver.chromium.org/downloads

i put both the webdriver.dll from the selinium module and chromedriver.exe in a folder on my C drive as the "working directory"

Updated 2/16/2022:
The latest iteration no longer requires an external word list.  At each run the word list is extracted from the game itself
Also new is recording each attempt's result, including game number, how many attempts and each word that was used.  By default the result saves to the running directory as results.txt

One common error you might recieve is that of the chrome driver version.  the error looks something like this:

"session not created: This version of ChromeDriver only
supports Chrome version 98
Current browser version is 100.0.4896.75 with binary path C:\Program Files (x86)\Google\Chrome\Application\chrome.exe 
(SessionNotCreated)"

This means the version of the chrome driver installed is for a previous version of chrome.  check your installed chrome version first by going to help > about google chrome.  use the link above to download the appropriate chrome driver and paste it in the same location as the original
