# EasyCustomUriProtocolManager

All custom URI's are registered to the app.vbs file, when app runs, it reads the URIconfigs.ini on what program to run, with what arguments for whatever protocol scheme was just ran. Write your URI's with their designated apps directly on the ini file, without having to manage registry changes or creating batch files to parse arguments.

use `_##_` as a separator to avoid issues with urls that may contain many special characters (csv may be easier but urls contain commas sometimes)

This is the pattern for the config.ini file which serves like a "database"

`YourProtocol_##_C:\Local Folder\program.exe_##_-arg /to --pass =to ^exe`

If you open your browser to `YourProtocol://https://ex.com/path, -/hah!*%"` this is what will actually execute:

`WScript.exe C:\saved\app.vbs "C:\Local Folder\program.exe" -arg /t --pass =to ^exe "https://ex.com/path, -/hah!*%`

(command '-/hah!*%' is used to indicate that all special characters work)

1. After any change made to URIconfigs.ini you must run applyConfigFile.vbs (this registers new Protocols to the windows registry)
2. Paths and Arguments will work in any format (with spaces, special characters etc..)
3. Make sure you use `_##_` as a separator without spaces before or after
4. Dont be stupid and run the example above, im just showing how things are preserved. Also use this at your own risk, read more about custom uri protocols and the dangers that come with it.
Enjoy.
fasterApp.vbs is the same thing as app.vbs but without any safetynet, will only save you 1 millisecond or 2, in performance.


# For Noobs -> What's happening?
download this repo, extract it anywhere you like
run applyConfigFile.vbs without changing anything
open your browser and copy/paste this `vlc://http://techslides.com/demos/sample-videos/small.mp4`
vlc will open on your PC and start playing the video

if that doesnt work, try firefox or the latest chrome
if that doesnt work open CMD and try `app.vbs vlc://http://techslides.com/demos/sample-videos/small.mp4`
if that doesnt work, install vlc bro

PS: Only applyConfigFile.vbs requires admin rights for writing on the registry, but you don't need to run it if you can write the reg changes yourself. just set path to app.vbs "%1"
