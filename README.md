﻿<div align="center">

## Multithreading with multi\-instancing and inter\-thread communication support

<img src="PIC200172705202310.gif">
</div>

### Description

This article teaches how to multithread safely and effectively using pure VB. No need to use C/C++ or the infamous CreateThread API notorious for its instability with VB 6. Simple and effective code complete with a HTML article documenting the various aspects of Multithreading are included to help you get started right away ! What's more, multiple-app instances are also now supported ! So say goodbye to Timers, crashes, threading API's and "freezing" forms. And if you find this code useful, remember that your vote will be GREATLY APPRECIATED ! (NOTE: When you create your own multithreaded app, set the threading model to "Thread per object")

NO EXPERIENCE OF C/C++ IS REQUIRED . SO NO NEED TO GO RUNNING ABOUT FOR A VC++ COMPILER. ENTIRELY BASED ON MICROSOFT'S ACTIVEX TECHNOLOGY FAMOUS FOR ITS SCALABILITY AND STABILITY. AND FOR THE SCEPTICS, THIS DEMO CAN CREATE 102 THREADS FOR A MEAGRE RAM MEMORY REQUIREMENT OF 2.5MB ALTOGETHER (A STANDARD BLANK FORM EXE TAKES AROUNG 2.0 MB OF MEMORY IN WIN98 FE)

IF YOU FIND THIS ARTICLE USEFUL PLEASE VOTE FOR ME !

[REQUIRES VB 5(Sp2) / VB 6)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-06-25 20:10:02
**By**             |[Srideep Prasad](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/srideep-prasad.md)
**Level**          |Advanced
**User Rating**    |4.8 (913 globes from 190 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Multithrea985736252002\.zip](https://github.com/Planet-Source-Code/srideep-prasad-multithreading-with-multi-instancing-and-inter-thread-communication-support__1-24672/archive/master.zip)





### Source Code

<html>
<head>
<title>Multithreading</title>
</head>
<body>
<p><b><font size="6">Multithreading - Understanding the pros and cons</font><br>
</b> One of the greatest problems of the earlier Win16 environment was that an application could do only one thing at one time (that is "single thread"). However with the advent of Windows NT 3.5x, this changed. In 1995, with the release of Windows 95 this ultra powerful technique came to be used in the common PC.<br>
 So what is the use of Multithreading ? - Consider Microsoft Word 97 or higher. It checks spelling while you type ! It does it by multithreading - i.e running two "threads" (in layman's language a "thread" is nothing but a piece of code, a sub or function running simultaneously with the main program). In VB 5, a new function AddressOf was introduced that enabled VB programmers to get the address of any public function in a standard module. This enabled developers to use the CreateThread API to create raw Win32 threads. Though this was effective in VB 5.0, with VB 6.0 it crashes miserably !<br>
Even at planet-source-code.com, I came across a multithreading demo using the CreateThread API. Though THE PROGRAM works with VB 6, it is VERY unstable. Also For ... Next loops, Msgbox..., Open .. etc statements do NOT work in the multithreaded procedures!<br>
</p>
<p>
Does this mean we cannot multithread safely !? Does it mean that we have to worry about
"exception errors" and GPF's popping up any time ?
</p>
<p>
The Answer is a BIG NO !<br>
Multithreading is VERY easy once you master the concepts... So just have a look at the sample code.. You will understand just how easy it is to perform true and safe multithreading in VB !<br>
If you gained any information, or if this article is useful to you, a vote of yours will be appreciated. If you found it useless.... just DELETE it !<br>
</p>
<p><font size="4">Multithreading In VB 6 - The Safe Way<br>
</font>
</p>
<p>The trick to effective and safe multithreading in VB is to use the ActiveX EXE project type (set to standalone EXE). The trick here is to create a new object on a new thread by callin the CreateObject() function and to create the Form that you want to be multithreaded from within this object. As a reult the form is created on a new thread, both of them can run almost independently of the other ! The only problem is managing the code re-enterancy - VB calls the Sub main() procedure every time a new object is created - we must find whether the main window is shown or not - if not we must initialize it. This method is actually very easy ! Just check out the sample code and I am sure that you will be Multithreading right
away
</p>
<p>What's more - this code now even demonstrates how to communicate between
threads !
</p>
<p> And if you found this code useful - be sure to vote for me ! After all, coding is a tough job, and so is writing a tutorial
!<br>
</p>
<p><b>IMPORTANT: You can now download a new generic multithreader component at
the following link </b><a href="http://planet-source-code.com/vb/default.asp?lngCId=26900&lngWId=1">http://planet-source-code.com/vb/default.asp?lngCId=26900&lngWId=1</a> 
. <b>This component allows you to multithread any sub or function in a standard
EXE. No ActiveX EXEs needed (PS:Thanks for all your votes And I am happy to know
that my articles are of some use to you !)</b><br>
<br>
</p>
</body>
</html>

