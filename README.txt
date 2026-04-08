======================================================================
•▀			AutomateXL Setup Guide
======================================================================
Written by: André Hacker
----------------------------------------------------------------------
Latest Revision: 4/8/2026
----------------------------------------------------------------------
Version: 1.0.0
----------------------------------------------------------------------
Developer(s): André Hacker
----------------------------------------------------------------------
Contact:

Email: andreissoftware@gmail.com
Social: @a_t_hacker 
Web: https://github.com/a-t-hacker

(Don't hesitate to reach out if you're having any issues!)

Thank you for your support! <3

/====================================================================================================================\
AutomateXL is an Excel (VBA), xlAppScript hosted application for screen mapping click and/or key events.
/====================================================================================================================/

License Information:

Copyright (C) 2022-present, André Hacker.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, 
THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES 
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) 
HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=====================================================================================================================\
INSTALL GUIDE:
=====================================================================================================================/

Default Install Guide (Recommended):

1. Place the first "AutomateXL v1.0.0 Download" folder from the .zip file on your desktop

2. Open the "setup.xlsm" file. You should be prompted w/ a message stating whether or not you accept running macros near the top. Click to accept. You should then be prompted if the connection was completed or not.

***If the connection fails during step 2, you may need to manually create the following folder locations below:

C:\Users\UserEnvironment\.xlas

C:\Users\UserEnvironment\.xlas ----> \autokit\ ----> \automatexl\

Add these 5 five below to the "automatexl" folder:
- \app
- \debug
- \mtsett
- \scripts
- \shell\win

****************************************************************************************************************
3. Go to your "C:" drive, find your "Users" folder, select the current user's home folder & find the directory titled ".xlas"

4. Within the ".xlas" folder click the folder titled "autokit", then the folder titled "automatexl"

5. Move the "AutomateXL.xlsm" file to the "app" folder inside your "automatexl" directory

6. Move the "show_automatexl.vbs" & "show_automatexl.cmd" files to the "win" folder inside the "shell" directory

7. Find & click the "shell" folder, click the "win" folder, & edit this path inside "show_automatexl.vbs": ("C:\Users\EDITHERE\.xlas\autokit\automatexl\app\AutomateXL.xlsm")

8. All done!

================================================================================================================