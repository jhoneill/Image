# Image
This repo is a module I have used for a long time named image, which I'm trying to prep for sharing, but is not yet in a fit state.

There are 4 pieces. 
1. Code orginally written by James Brundage https://github.com/StartAutomating and published by Microsoft in a module name PSImageTools
as part of the PowerShell Pack (https://devblogs.microsoft.com/powershell/introducing-the-windows-7-resource-kit-powershell-pack/). 
This uses the WIA COM objects, to scale, crop, rotate and flip imanges, convert to JPG or BMP, and add overlays. 
I added more things to work with WIA Objects, including **Reading and Writing EXIF data**
2. **Read-Exif** Exif is not the only metadata stored in an image file, there is also IPTC data - which may be inside photoshop data,  XAP Data, etc.
there are tools to get this data - in particular Windows File properties which will display IPTC title, caption etc or Exif versions of them, 
but I wanted to be able to see the XAP which Adobe light room adds, and to Parse the MakerNote fields form my cameras. 
I'm a Pentax shooter so there are a few bits of information from other camreas but **Read-Exif** will read generic data from all cameras 
and make sense of Pentax specific data. It reads JPG, TIF and RAW files (which are specialized TIF files and can be treated as TIFs for reading metadata). 
Writing metadata can be doen with the WIA objects (although this seems delete XAP and Photoshop data), and is better done with tools like exiftool. 
Exiftool https://exiftool.org/ will get this information but I wanted to get information which it doesn't readily display, and wanted it in 
the form of PowerShell objects - so went to my drawing board and created Wheel 2.0 from an blank sheet of paper! 
3. **Lightroom** tools. Underpinning Adobe lightroom is a SQLite database. I can read AND write to this database with my GetSQL 
module https://www.powershellgallery.com/packages/GetSQL and first wrote about doing so here 
https://jhoneill.github.io/powershell/photography/2012/08/09/Lightroom-data.html - and apart from Adobe compressing some fields in a way that 
breaks compatibility I've been able to keep using the same queries and adding to the functionality ever since. Writing to the lightroom catalog 
is not something adobe support so ensuring the datbase (.LRCAT file) is backed up is particularly important. But it means I can set 
colour labels, flags (pick / reject), keywords and collection membership external, and get a list of files which meet certain 
criteria (including exported and converted to black and white) 
4. Sundry tools for taking other data and converting it to use for tagging images - GPS information for example. These are in most need of cleaning up. 

