**ListeningNowTracker application** integrates Spotify track titles and your Skype profile together (WindowsXP/Vista/Win7 platform).

When you are listening music on [Spotify](http://www.spotify.com) (MS MediaPlayer supported also), Spotify can optionally update your [Last.FM](http://last.fm) account to keep track of the music you have played. Spotify can also update MS Messenger account in real time to show _I'm listening now ThisSong made by ThisArtist_ type of text. Unfortunately Spotify doesn't do this for [Skype](http://www.skype.com).

This application sits in between Spotify and Skype. When ever Spotify starts to play a new track, the application updates the "Mood text" property of your Skype profile. This text is shown to other "Skype people" right next to your name.

Please read [README.TXT](http://code.google.com/p/listeningnowtracker/source/browse/trunk/ListeningNowTracker/README.TXT) file for more details.

**How this work?**

When you play a music track in Spotify, it sends out "new track" events to applications registered to listen these messages. This application listens this Windows event.

Input parameter of the event is song title and artist name. ListeningNowTracker application then updates the ["mood text" property](http://support.skype.com/faq/FA594/What-are-mood-messages) of the active Skype profile using Skype OLE API programming interface ([Skype4COM](http://developer.skype.com/Docs/Skype4COM)).

You can control the format of the "Listening now" text through INI file. By default there is _"Listening songTitle by artist from Spotify"_ format mask.


**How to use it?**

Start Spotify and Skype business as usual. If you want to keep track of Spotify tracks in Skype then start this application also (doesn't matter in which order you start these apps). That's it.


**Do I need to change something in Spotify or Skype?**

No. Spotify or your Spotify account doesn't require any changes or configurations. Spotify sends automatically these "new track" events.

Skype and Skype account doesn't have to be changed either, except that for the first time you have to grant ListeningNowApplication a permission to update mood text of Skype profile at runtime. This is something which Skype takes care of automatically when it detects a new application using Skype API services (see [README.TXT](http://code.google.com/p/listeningnowtracker/source/browse/trunk/ListeningNowTracker/README.TXT) file for more info on how to grant and reject these external application permissions in Skype).


**Downloading**

There is both Windows binary executable and source codes available. See "Downloads" page for more info.