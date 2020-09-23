+====================================================================================================+
* OpenGL in Visual Basic is a largely unexplored area. I did a search on PSC for OpenGL and came up  *
* with a mere handful of results. Some of the submissions were quite good, however I felt that they  *
*  were a little too complex for the beginner in OpenGL. This submission is an altered version of a  *
* port to VB from C++ (Original tutorial + VB port available on http://nehe.gamedev.net). This demo  *
* initializes OpenGL in fullscreen mode and displays a mipmapped textured cube which you can rotate. *
*   The main feature I'm trying to demonstrate here is loading a TGA file from disk and using it as  *
*    an OpenGL texture. You will need hardware acceleration of some form to view this demo at any    *
*        form of acceptable frame rate, but I consider this an acceptable requirement.               *
+====================================================================================================+


Thank you for downloading my submission to planet-source-code.com! Hopefully you will
find this code quite useful. I have taken time to document the source code so that
people who aren't very good at reading code can get a general idea of what's going
on. As a rule, the code in modOpenGL you do not need concern yourself with as no
modification is necessary at any point. Just include the module with any project you
want to make OpenGL enabled and you will be able to initialize OpenGL in either
fullscreen or windowed mode and shut down OpenGL. modMain.bas just contains a couple
of application specific global declarations.

cTexture.cls is my own work and is really the focal point of this demonstration 
application. cTexture is a class designed specifically for OpenGL and provides a 
generic method for opening multiple image types and using them as OpenGL textures. 
Again, the actual code contained within is not important unless you are thinking of 
adapting my code or plain ripping it off °¿° :)
                                          -
At the moment the following file formats are supported:+
+======================================================+
* 24BPP bitmap
* 32BPP bitmap
* 24BPP TGA
* 24BPP TGA + Alpha
* 24BPP RAW (RGBA bottom to top format) + Alpha


   ææææææææææææææææææææææææææææææææææææ .-= LockJaw3D =-. æææææææææææææææææææææææææææææææææææææ
  |                                                                                            |
  | cTexture is a part of LockJaw3D, an open source 3D game engine being developed in pure VB. |
  |  I am the founder of the project and so if you wish to learn more about LockJaw3D or you   |
  |  want to help in some way then please feel free to contact me at thegilb@hotmail.com. A    |
  |    website is in the pipeline too - when it's up I'll post the addy on PSC in my user      |
  |                             details or in my next submission :)                            |
  |                                                                                            |
  +~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~+


This at first glances might well seem complex due to the sheer amount of code. However,
the vast majority of the code here is reusable and it doesn't get much more complex
than this. It's as simple as telling OpenGL what to render, then OpenGL renders it.


 INSTALLATION
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
You should find in this zip file a copy of vbogl.tlb put this file in your
Windows\System directory. This is the type library VB uses to access OpenGL. It's
basically the same as typing out all the OpenGL functions like you would functions
in the Windows API, it's just more convenient and neater to reference this library
from VB. NOTE: If you make an OpenGL project, remember to reference this library!
To do so go to Project->References and search for "VB OpenGL API 1.2 (ANSI)". If
you can't find it there and you're sure it's in your Windows\System directory then
you need to register the type library. Luckily VB will do this for us automatically
when we first reference the library so click on the browse button and locate
vbogl.tlb. All done!

Compilation should go through without a hiccup. If you get nothing but a blank screen
when you run the application or you have other problems, please make sure you have
the latest OpenGL drivers for your 3D card by visiting www.glsetup.com. If you're
still having problems then make sure that crate.tga is in "\Data".

KEYS:+
+====+
* Up Arrow 	- 	Increase y rotation speed of the cube
* Down Arrow 	- 	Decrease y rotation speed of the cube
* Left Arrow 	- 	Decrease x rotation speed of the cube
* Right Arrow 	- 	Increase x rotation speed of the cube
* F 		- 	Toggle filtering (None->Linear->Mipmapping cycle)
* L 		- 	Toggle lighting
* Page Up 	- 	Zoom in
* Page Down 	- 	Zoom out
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Next submission planned is some form of model loading and morphing code.

Oh yeah I nearly forgot! Don't forget to vote for me if you think this code is
worth it! Comments and constructive criticism welcome :) This is my first
submission to PSC which involves 3D - am I headed in the right direction? Talk
to me because I'm writing for you to try to get you to vote for me! What do you
want? Speed optimisations like frustrum culling and hidden surface removal? Or
things like model loading and animation? Tutorials or demo's with comments or
both? Whatever it is you want, email me and let me know and I will write it for
you!

-CG