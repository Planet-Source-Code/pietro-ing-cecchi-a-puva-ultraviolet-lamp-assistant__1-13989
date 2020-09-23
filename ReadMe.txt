Ultraviolet (PUVA) Tanning Lamp Assistant

README of 'Ultraviolet (PUVA) Tanning Lamp Assistant', by Pietro Cecchi
          email: pietrocecchi@inwind.it, cecchi@infinito.it
          http://members.theglobe.com/cecchi
          31 december 2000

INTRODUCTION:

This is the 'Ultraviolet (PUVA) Tanning Lamp Assistant' program 
posted on the www.planet-source-code.com/vb by Pietro Cecchi.
This is a complete, working application intended to assist the exposure
of human body to an ultraviolet lamp.
In fact, the program can be configured for up to four timed phases 
(front, left, right, rear sides) and will speak messages at the
start of each phase and at the end of sequence (exposure).
The voice assistance (you need your speakers on) is basic, because
during these exposures (that can last up to 60 minutes and more, 
depending on the lamp emitting power and nature), it is really
possible to fall asleep!!! The assistant will awake you anyhow,
avoiding to you a bad skin burning.

Normally, a commercial PUVA lamp has only a timer (one hour), and 
in the case of multiple exposures (the various sides), each time 
the lamp has to be timed for the 'phase duration' (example 7 to 
15 minutes, see the instruction booklet that comes with the lamp).
PUVA lamp exposures has very benefic effects on your skin, and moreover
gives that special look that someone likes so much.

NOTE: I have a Philips 120w PUVA lamp, that mounts the same tubes
(but smaller, 20 watt instead of 100 watt) than the ones mounted 
in the PUVA units (big like telephon cabs) used in the Hospitals. 
I attached the description of my lamp in the file LampDesc.doc 
(take a look at it).

How I use the program with the lamp?

I start a 'midi' player (so I can control the volume of music) as
music background, then I launch the lamp assistant program (never
forget to dress the special UV goggles that come with the lamp,
or you will become blind in few exposures!!!). This is a better way 
to spend a long hour of your time. Furthermore, it is faster, 
because there are no stops between the phases, and can be considered
equally safe, if you have no hearing problems.

Now, let's go back to the program. See the screen shot.
The few commands (the fewer possible) in the big screen (with the 
goggles you can't see much, really) allows to you to:
-set the phases you want to run (from 1 to 4, even 0 is allowed), in 
 the order you click over them
-set the durations of the phases, in minutes
-save the configuration
-load a previously saved configuration
-run a beautiful audio demo, lasting about 8 minutes (very original 
 code structure)
-start, pause, abort the sequence (pause is very useful, don't ask me
 why..., and the abort button is protected against accidental clicks, 
 in fact it is timed, and needs to be reclicked to accept the command)
-acknowledge the natural end of sequence or the user abortion of it
 
To test the program run short phases durations (also 0 minutes, or 1).

I recommend this code for any level, from beginners to 
superprofessional fellows who kindly read this popular web-living
VB magazine written by us amateurs. I think this code is good for
teachers too, who teach VB (or not) in the schools: the pupils play 
with it, and learn something new and useful.

WHAT YOU CAN LEARN FROM THIS APPLICATION:
-how to approach the project of a sequencer
-how to use relatively new controls (Office) for a better performance 
 (CommanButton with ForeColor, SpinButton, Image with Picture Tiling 
 property for the background of the form, CheckBox in the Graphical 
 Style: all controls of Microsoft Forms 2.0 Object Library)
-how to structure a living audio demo
-how to use the TextToSpeech (VTEXT.DLL control) for vocal interfaces
-how to build and let work the fancy phases indicators

ADDITIONAL NOTES:
-the background (very ultraviolet!) is a gradient of mine and is used
 in my web site too (visit it, is not a VB site)
-the Image control used to show the background, can also zoom a picture,
 automatically and in a very fast way, id est you can make a fast image 
 viewer with it, in a glance.
-the phases progress indicators are a must for exigent programmers, 
 remember the Picture box and the routine (UpdateProgress) used to 
 build them


A LAST VERY IMPORTANT POINT:
If we have to expose a body for a certain advised maximum duration per 
day, e.g. 10 minutes, we have the following possibilities:
-expose 10 minutes one side (e.g. rear) and 10 minutes the opposite one
(e.g. front)
-or expose 5 minutes all sides (id est rear, left, right, front in 
whatever order)
Please remember always this, otherwise you may overexpose yourself
seriously.
So, if you decide to expose three sides, you have to expose them for 5 
minutes, in order not to overexpose the side in the middle.
As a rule, if the sides are one or two (opposite), just apply the maximum 
time, and in the cases of three or four sides or even two but adiacent, 
divide by 2 the maximum time, id est expose each one side for 5 minutes.

 

THANKS FOR READING:

I hope you enjoy this program. Should this be the case, please take 
few minutes to rate it. 

Best programming to you friends, 
   Pietro Cecchi 


With the occasion, I wish you a really Happy New Year!


POSTED ON www.planet-source-code.com the 31 december 2000, by PIETRO CECCHI
COMPATIBILITY: This program has been written in VB6
LEVEL: I would say: 'any', even if I selected 'intermediate'
TYPE: complete application