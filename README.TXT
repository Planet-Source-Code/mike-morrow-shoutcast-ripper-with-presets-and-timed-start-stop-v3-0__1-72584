GetShoutcast - Download mp3 streaming

With GetShoutcast you can save to disk the mp3 data stream 
from the on-line radio while you listen to the audio though 
an internal instance of Windows Media Player.  

-------------------------------------------------------------

Comments added by Mike Morrow  on  Nov 7,2009

.aac support is in the program!  I now support both formats
of Shoutcast streams as of today.  Who knows when more will
appear...

*** N O T E *** You must get this download to be able to
play .aac files in WMP (used by this program).

http://www.orban.com/plugin/

Download and install the codec, here, to be able to play and
hear .aac file streams with this program and with WMP. There
may be other codecs, I expect but this one was the first in
Google, seemed professional and it worked.

-------------------------------------------------------------

Comments added by Mike Morrow  on  Nov 6, 2009

Extensive changes in the last two weeks.  A new form with code
copied and highly modified from another bit of code on PSC
submitted by Mario Grimm which had no error checking.  It
could hang on any Shoutcast server interchange.  Admittedly
the are rare but it is still nice to detect and stop waiting.

Removed is the Winamp server.  It was undependable.  I wish it
worked though then one could play without recording.  I may
try to revisit that at some time. 

Added to Mario Grimm's code was an enabled timer wait to lower
the CPU spin utilization while waiting on the Shoutcast server
to respond.  Admittedly, it is a small thing unless the server
never answers or the answer is corrupt and, then, the program
would remain 100% busy until crashed by the user, waiting on a
valid resonse which was never to arrive.

The spin termination is by way of a counter that aborts the
wait after 400 waits of at least 20 milliseconds.  This sets a
maximum wait of about 8 seconds for the Shoutcast information
to return.  After that, I declare the data is here, notify the
user and abort trying.  The user may resend the request, if
desired, and after finding out if the internet connection has
a problem.

Also added was multiple stream support from one station.  One
station has 23 streams under one banner.  Others have 2 to 4.
Most have multiple listings if they have multiple streams.  I
detect this along with the user count for each of these
streams and allow the user to select the desired stream.  It
may matter if the streams are on disparate URLs (IP addresses)
and the heavier uses may be intermittent.  It may not matter
in actual use!  Just something that was there in the data so I
wrote code to use it.  Probalby just Mikey Overkill in action!

Removed was the beautiful but irritating and useless colorful
banner and the plain old Windows standard banner is back.

Color coding has been implemented to try to make boxes and
buttons relate to each other.  Also buttons on one form do a
similar function on another form with the same color.  This is
almost an accidental afterthought and probably needs some
attention.  But I like the color.  On a form, colors are
coordinated.

Still to come, the continuous, automatic recording of multiple
streams in sequence.  It is not far away.  Within the week,
that should be running and that is all I propose for this
project for now.  It will then do all that I want.  Probably
be bug fixing for a few weeks or more.  Sometimes they take
years to find!

Submit change, bug and suggestion reports to Mike Morrow
either via the web interface on PSC or directly with an
obvious email subject (to avoid the SPAM filter) to
mailto:mikey@ILikeTheInternet.com

-------------------------------------------------------------

Comments added by Mike Morrow  on  Oct 26, 2009

The external WMP solution works fine except the time slider
gets lost almost immediately.  No big deal, really. Still in
the queue to test the internal one.  It may work just as well
but I will have to add buttons for it.  The work never
ends....  ;-)

There is room for 20 stations in the memory form.  I need to
convert that over to listview.  The work never ends.... ;-)

The automatic, timed recording of one entry is now working
properly. Coming is the automatic recording of all stations
set up for that.  There are other stream rippers that claim
timed record but most don't work.  The form is there but the
code isn't.  I also check for overlaps and report them so you
can fix them.

Also coming, full integration of Shoutcast stations right in
the program.  That, instead of having to use Winamp to find
the station and then copy if across.

Also coming is support for AACP stations.  There are 7,800 of
them listed in the full Shoutcast directory.  They do have
lots of duplicates so it is not actually that many but it does
mean that Shoutcast rippers will have to support .aac soon.

But the good news is that it is easy.

So, lots more to come until I get totally bored with it and
someone else takes it up and carries on.

-------------------------------------------------------------

Comments added by Mike Morrow  on  Oct 24, 2009

Well, the playback problem has been cut down to size.  There
may be a further improvement coming by internalizing the WMP
but that will take some further research.  Then the program
can be a Shoutcast recorder and player.  Trying to get there,
yaw'll...

The best way to monitor the recording is to let it start up
and record it and then after a few seconds, use the menu item
to startup WMP playing the file.  The time slider in WMP will
not be accurate but the sound will be.  So go to the File menu
and play it with WMP and enjoy.  Of course, you will have to
point to WMP on your system, first, with the first menu item.
Then the second menu item will start WMP pointed to the file
and it will play the file a little behind real time.  The
Winamp connection is still not working after more testing,
double and quadruple buffering finally up to 10 buffers.  All
to be stripped out some time soon.

Comments added by Mike Morrow  on  Oct 21, 2009

Don't worry about your English and don't apologize.  We
appreciate you writing in English so we don't have to learn
your language.

And, thanks for the code base.  I fixed a couple of very minor
concerns and added the first coat of varnish on the old girl
to make her record a single item.  I may go back and make it
record multiple streams in a row.  This would be very nice for
me to record KCSM.org on Friday night for the 3 hours of blues
and then record one of the NPR stations to get Prairie Home
Companion on Saturday night.  So, I am doing this for the both
of us.

Again, thanks for the code, your English is just fine with us
and I hope that you don't mind me tweaking your stuff.  I left
your name in there and in this readme file.

Also pending is to find out why the current Winamp does such a
bad job of playing the stream from this program.  It rarely
works right.

That one will take some learning on my part!!  Again, thanks
for the jump start on the code.

mailto:mikey@ILikeTheInternet.com

-------------------------------------------------------------

Original comments by Fernando Aldea G. before Feb 17, 2005

With GetShoutcast you can save to disk the mp3 data stream
from the on-line radio while you listen audio though winamp
player (for example). This work like as "sniffer".

-You can to use Winamp Player (or no)
-You can create mp3 file by song.

Feb 17, 2005:  Bugs Fixed (the "\" in the out path and path in
url)


Programmed by Fernando Aldea G.   (Valparaiso, Chile). 
mailto:fernando_aldea@terra.cl

Sorry about my English  :)


