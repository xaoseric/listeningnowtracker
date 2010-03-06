
ListeningNowTracker
-------------------

Spotify is an online music player and content store. On the other hand Last.fm is a great 
community based music statistics database (what have you and others played and 
recommendations based on that statistical data).

Fortunately Spotify player supports out-of-the-box linking with Last.fm accounts, so whatever
you listen on Spotify is tracked by your own Last.fm profile. This is nice because you
get recommendations to other bands and artists you might like.

Some people would like to show the currently played music track (artist and song title) in
instant messaging systems also. Spotify supports linking with MSN-Messenger tool, but not 
with Skype tool.

This application links Spotify and Skype together in a way that the title of the currently
playing song is shown as a "mood text" in Skype profile.

Visit http://code.google.com/p/listeningnowtracker/ home page for up-to-date information.


INSTALLATION
------------

(1) Make sure you have Spotify (or Windows MediaPlayer if you want to show those tracks in Skype)

(2) Unzip the archive file to, for example, "c:\program files\ListeningNowTracker\" folder
    You should see following files
      ListeningNowTracker.exe		The main executable doing the magic
      LICENSE.TXT			License text ("MIT" type of opensource license)
      README.TXT			The file you are reading

(3) Start "ListeningNowTracker.exe" application, Skype and Spotify (doesn't matter in which order).

(4) When Skype detects for the first time that "ListeningNowTracker" application wants to use
    Skype application, it throws the "External application wants to use Skype" warning message. 
    You have to open Skype and click "Approve" option to allow this application to use Skype 
    services. 

    You can manage these permissions in Skype Preferences dialogbox if you missed the first 
    approval event (Skype | Tools | Options | Advanced | "Manage other program's access to Skype")


That's all. Start playing music in Spotify and see how your Skype "mood text" is automatically 
updated.

Well, in a normal case that's all. Sometimes there could be, believe it or not, some problems. 
Here are the most typical cases.


TROUBLESHOOTING
---------------

ListeningNowTracker application shows error messages as tooltip text in a tray icon. Move mouse 
cursor on top of the tray icon to see the text. If you don't see the tray icon then press 
small "<" arrow in Windows taskbar to see hidden tray icons.


*** When I play music, the tray icon of "ListeningNowTracker" shows text "ERROR: Could not create 
    coclass".
  
  Usually this means that "Skype4OLE" OLE automation objects are not installed. Normally Skype4OLE is 
  installed with Skype, but in a custom installation of Skype it is possible to skip that installation 
  opion (Extension Manager).

  The solve the problem you should install Skype4OLE library manually
  - http://developer.skype.com/Docs/Skype4COM/
  - Download the latest Skype4OLE library (Feb 2010 it is Skype4COM-1.0.33.zip) 
  - Unzip the package to c:\program files\Skype\ folder (or any other folder. Doesn't have to be Skype folder)
  - Register the Skype4OLE library using the following command line (cmd shell)

        regsvr32.exe "c:\program files\Skype\Skype4OLE.dll"


*** When I play music, the tray icon shows text "WARNING: Skype is not running. Cannot update profile text".

  I guess the message speaks for itself. Start Skype and login to your Skype account because otherwise 
  this tool cannot update the "mood text" property of your Skype profile.


*** When I play music, the tray icon shows text "ERROR: invalid property/method invocation"

  Some kind of OLE communication error with Skype. Check at least following issues
  - Make sure Skype is running and you have logged in Skype (if you have Skype running, but it 
    doesn't have active session then you will see this message)

  - Install or re-install Skype4OLE library (see above Skype4OLE.dll tip). No need to re-install 
    the whole Skype if you really don't want to.

  - Permission problems with current Windows user to use OLE automation services. Guest user may 
    have limited rights to use local services.

  - Hmmm... Something else going on. Please report the problem and supply Skype version, Windows 
    version, Skype4OLE version (use FileExplorer and file properties to see the file version) 
    as reference information.


*** Spotify and Skype are running, but this application doesn't seem to do anything. No error 
    messages and no "Listening now" text in the tray icon of this application or in Skype

  Did you remember to allow "ListeningNowTracker" to use Skype services? See INSTALLATION section
  of this file.
  
  Or maybe latest version of Spotify has changed the event format (see technical background above). Please
  report the Spotify version you are using.

  Or maybe some other application is blocking the "ListeningNow" event? Do you have some other 
  application running which shows the same "Listening now" message?  

  If you have "Microsoft Messenger" or "Microsoft Live Messenger" running (but don't need it because 
  you use Skype) then try to close Messenger to see whether it helps. Don't just close the Messenger 
  dashboard, but close the whole application (select "Quit/Exit" from the popup menu of Messenger 
  tray icon). If that helps then make sure Messenger service is not automatically started at boot time 
  by modifying Messenger startup option.

  If you need both Skype and Messenger, but you'll see the "Listening now" text only in Messenger 
  then I guess Spotify fails to propagate the message to ALL windows registered as having a class 
  "MsnMsgrUIManager" (see technical details below). This is Spotify problem. Sorry, can't help there.

  Or maybe you could throw away "Messenger Live" and use some alternative MSN Messenger tool? 
  See Google for for more info or "http://www.osalt.com/msn-messenger" web page. I don't have anything 
  to do with those alternatives or web pages, so please use your own common sense which way to go.


TECHNICAL BACKGROUND
--------------------

Spotify sends "currently playing" information to MSN-Messenger through a special message,
which is normally listen by MSN-Messenger only. This application listens the same message
and re-routes the event to Skype by modifying "mood text" automatically whenever a new message
is received.

  (1) Spotify starts to play a new song
      - Spotify sends out WM_COPYDATA event to a window registered as "MsnMsgrUIManager" class.
      - The event has 0x547 identifier number and data block uses following format
          \0Music\0status\0format\0song\0artist\0album\0
 
         Null terminated string values with "\0Music\0" prefix tag followed by
            status: 1 = Playing, 0 = Stopped or paused
            format: ? (no idea what this is in Spotify's case)
            song:   Title of the song
            artist: Name of the artist
            album:  Name of the album

  (2) MSN-Messenger listens "WM_COPYDATA, 0x547" event and updates its "Now playing" status info
      whenever this message is received.
          

This applications does the same as MSN-Messenger, but uses Skype API library to update Skype
profile data based on the artist and song title.

Skype communication is implemented through Skype4OLE OLE automation interface. Skype4OLE objects
are supplied by Skype application, not with this application. This app just uses those OLE objects.


FAQ
---

* Does this application require any changes in Spotify or Skype?
  - Not at all. This is independent application between those two and doesn't require any
    custom changes in Spotify or Skype. Only requirement is that Skype was installed with
    "Skype Extension Manager" option (Skype4OLE objects). See installation section for more info.

* Does this application consume some processing resources (CPU) when Spotify or Skype is not
  running even when this application would be running in a background?
  - Not at all. The "MsnMsgrUIManager WM_COPYDATA" technique doesn't require any polling, so
    when this event is not sent then this application does nothing (the process is sleeping)

* Do I have to provide my Spotify or Skype account name or password in the configuration of this app?
  - No. This app doesn't need those. Application updates the mood text of the active Skype
    profile through Skype's own API services, so only Skype knows your password. And it is up to
    Spotify to send the WM_COPYDATA event, so this application doesn't even know that the message
    came from Spotify. In fact, Windows MediaPlayer sends the same event.

* Does this still work with the built-in Last.fm integration of Spotify? 
  - Yes. This app has nothing to do with that integration.


LICENSE
-------

DISCLAIMER: Use this application at your own risk. There is no warranty of any kind.

Feel free to modify or re-distribute this application as long README.TXT and LICENSE.TXT files
are supplied with the application also.

See LICENSE.TXT for more information about the license (opensource, free, re-distributable application).

Copyright(c) 2010 Mika Nieminen. All rights reserved. 

Here by you are given permission to derive your own work from this application, re-use the source code, 
modify it and take parts of it into your own application as long derived work is also opensource application
and you provide both binary and source code format of the work. You are encouraged to contribute back
to the "community" (whatever that is).


TRADEMARKS
----------
- Windows, MSN-Messenger, Messenger Live are trademarks or registered trademarks 
  of "Microsoft Corporation" (www.microsoft.com)
- Spotify is a trademark or registered trademark of "Spotify Ltd" (www.spotify.com)
- Skype is a trademark or registered trademark of "Skype Ltd" (www.skype.com)
- Last.fm is a trademark or registered trademark of "Last.fm Ltd" (www.last.fm)

------------------------------------------------------------------------------------

http://code.google.com/p/listeningnowtracker/

