Attribute VB_Name = "xlAppScript_logs"
'/\_____________________________________________________________________________________________________________________________
'//
'//     Developer(s): André
'//     Contact: xlappscript@gmail.com | andreissoftware@gmail.com
'//     Web: xlappscript.dev
'/\_____________________________________________________________________________________________________________________________
'//
Sub lex_ChangeLog()
'//=========================================================================================================================
'//
'//         lex CHANGE LOG
'/\_________________________________________________________________________________________________________________________
'
'
' Version 1.2.1
'
' [ Date 4/8/2026 ]
'
' (1): Updated copyright, contact, and web information
'
' Version 1.2.0
'
' [ Date: 1/31/2023 ]
'
' (1): Added "clear" statement for removing a named range value by it's index
'
' (2): Added "while" instruction for looping until a set condition is not met
'
' (3): Added "group" structure & "call" instruction for organizing code into recallable blocks
' ***Groups are similar to subprocedures as a "group" will not return a value
'
' Example:
'
' group(groupName){@xlasCode;}endgroup;   <--- Create group
'
' call(groupName);    <--- Parse group
'
' (4): Added "get" statement for retrieving a value from a named range by it's index
'
'
'
' [ Date: 1/30/2023 ]
'
' (1): Complete refactor of lexer & parser architecture
'
' (2): "for-next" support
'
' (3): Use of collections for parsing through tokens as opposed to stringing tokens through functions
'
' (4): Added special character support for braces "{ }" & brackets "[ ]"
'
'  Left Brace = [lbr]      Left Bracket = [lbrk]
'  Right Brace = [rbr]     Right Bracket = [rbrk]
'
'
'[ Date: 1/29/2023 ]
'
' (1): lexer/parser readability changes
'
' (2): Changed "lexKey" function call to "xlas"
'
' (3): Added "[dblquote]" & "[sinquote]" special characters for returning double or single quotations
'
'
'[ Date: 12/30/2022 ]
'
' (1): Changed instruction pointer separator ",#!" --> [,]"
'
'
'
'[ Date: 12/24/2022 ]
'
' (1): "let" instruction has been changed to "set" instruction ("let" properties will remain the same for "set").
'
' ***When setting a variable's value to another variable's value you can use the "set" article.
' ***Another method is to use the "=" operator between the variables.
'
' Example(s):
'
' set @var1=@var2; <--- @var1's value will equal @var2's assigned value
'
' @var1=@var2; <--- @var1's value will equal @var2's assigned value
'
'
' (2): Fixed bug in "kinSet" function where setting a variable to another variable using an article like "sum()" would
' cause the parsed variable to equal itself.
'
' (3): Fixed bug in "kinSet" function that disallowed multiple variables from parsing their values when used on the same line
'
' ***Performing an operation on a variable assigned to itself will cause the assigned variable value to take precedence.
'
' Example(s):
'
' @var5=sum(@var1 + @var2 + @var3 + @var4 + @var5); <--- This will equal @var5's assigned value
' @var6=sum(@var1 + @var2 + @var3 + @var4 + @var5); <--- This will equal (@var1 + @var2 + @var3 + @var4 + @var5)
'
'
' (4): Changed "[feed]" special chracter string to "[lf]" for "Line Feed"
'
'
'
' Version 1.1.6
'
'
'[ Date: 8/11/2022 ]
'
' (1): Fixed issue where lexer would loop infinitely when using multiple instructions
'
'
' Version 1.1.5
'
'
' [ Date: 8/8/2022 ]
'
' (1): Fixed issue w/ goto statement causing error due to incorrect function call
'
'
' [ Date: 8/2/2022 ]
'
' (1): Updated WinForm #'s for eTweetXL
'
'
' Version 1.1.4
'
' [ Date: 6/28/2022 ]
'
' (1): Changed "fnd" function prefixes to "get"
'
' (2): Added "setWindow" function for setting current & last WinForm #'s when activating an application Window
'
' (3): Updated "runScript" function to set "xlasErrRef" block address to it's assigned cell from setup
'
' (4): Additions (AutomateXL) & changes to application "WinForm" #'s
'
'
' Version 1.1.3
'
' [ Date: 6/21/2022 ]
'
' (1): Fixed issue where "goto" statement would cause an error b/c of invalid ".Value2" as ".Value22"
'
'
' Version 1.1.2
'
'
' [ Date: 6/6/2022 ]
'
' (1): Removed black & red text changes to current Run Tool from the main "lex" module
'
' (2): Minor label name changes
'
' (3): Changed special character identifiers for ($), (;), ((), ()), & (#):
'
' $ --> [doll] | ; --> [semi] | ( --> [lp] | ) --> [rp] | # --> [hash]
'
'
' Version 1.1.1
'
' [ Date: 5/24/2022 ]
'
' (1): Made minor adjustments to WinForm #'s in "fndWindow" & "fndRunTool" functions
'
'
'
' Version 1.1.0
'
' [ Date: 5/11/2022 ]
'
' (1): Code optimization, function renaming, etc.
'
'
' Version 1.0.9
'
'
' [ Date: 4/23/2022 ]
'
' (1): Removed "enableWbUpdates" & "disableWbUpdates" functions from running automatically, and moved them to the "xbas" lib
' as a library dependent switch.
'
'
' (1): Adjusted "kinSet" & "kinExpand" functions to accept alternative cases w/ variables such as parsing a variable call identifier
' (@) or (=) through a file or when needing to set a variable value from another variables value.
'
' Variable Preservation Sequence = %@var
'
' Example: %@var=read(-all @filePath); <--- if the text in this file contains an "=" this will still return all text to the assigned variable
'
'
' Version 1.0.8
'
'
' [ Date: 4/20/2022 ]
'
' (1): Added function for removing the "%" alternative variable escape character.
'
' When "%" is used in front of a call to a variable like "%@var" if the variable contains "@" within a string
' it will parse through the "@" identifier as if it's a normal string instead of interpreting the string containing @
' as a variable.
'
'
' [ Date: 4/18/2022 ]
'
'
' (1): Added "xlasGlobalControl" block for triggering a control that allows Local variable usage in external environments
'
'
'
'
' [ Date: 4/17/2022 ]
'
'
' (1): "libcall" bug fixes when using after another instruction
'
' (2): Changed lexer counter variable names to contain "Ptr" (as in "Pointer")
'
' (3): Added "[(]" & "[)]" special character escape sequences for left and right parentheses
'
'
' [ Date: 4/16/2022 ]
'
'
' (1): Added "[space]" special escape character
'
' (2): Reworked "<env>" & "<blk>" commands to accept spaces
'
'
' Version 1.0.7
'
'
' [ Date: 4/2/2022 ]
'
' (1): Fixed recursive loop caused by "end" instruction when used w/ certain variations of "if-else" instruction
'
' (2): Fixed "if-else" instruction not parsing parentheses in certain instances
'
'
' [ Date: 3/31/2022 ]
'
' (1): Made adjustments to escape character function & added return function for retrieving values before searching through
' a library.
' ***New sequence is [$] <--- where the special character is enclosed in brackets
'
'
'
' Version 1.0.6
'
'
' [ Date: 3/17/2022 ]
'
' (1): Added "xlasLocalContain"
'
'
' [ Date 3/16/2022 ]
'
'
' (1): Added Runtime block addresses "xlasBlkAddr79" - "xlasBlkAddr99" (21 total)
'
'
' [ Date: 3/13/2022 ]
'
'
' (1): Added "xlasLocalStatic" address for locking libraries to the current runtime session.
'
' Version 1.0.5
'
' [ Date: 3/9/2022 ]
'
' (1): Added "let" instruction for assigning variable values on the fly after declaration.
'
'
'
' [ Date: 3/8/2022 ]
'
' (1): Added "libcall" instruction to allow directing an article string/block to a desired library. This will place the
' desired library at the top of the library stack, which will help cut down on article querying time & avoid collisions
' where articles from different libraries share similar names.
' (in instances where many libraries are used, this will allow those lines/blocks of code to parse quicker).
'
'
' [ Date: 3/7/2022 ]
'
' (1): Added an escape key function & sequence for retaining special character(s) during parsing
'
' Example: --->  [^$] (for escaping the '$' symbol)
'
'
'
'
' Version 1.0.4
'
'
'
' [ Date 3/2/2022 ]
'
' (1): Added "Alternative Expansion" functionality & characters "{", "}", & "~" which allow articles to ignore initial parsing.
'
' Replace enclosers/enders (() & ()) w/ ({) & (}), & (;) w/ (~)
'
' w/ Alernative Expansion = rng{A1}.value{100}.bgcolor{gainsboro}.fcolor{cornflowerblue}~$
' w/o Alternative Expansion = rng(A1).value(100).bgcolor(gainsboro).fcolor(cornflowerblue);$
'
' (2): Added check for "<blk>" article.
'
' The "<blk>" article's purpose is for manually setting the runtime block (worksheet)
'
'
' (3): Seperated instructional components from lexer (goto, if, do, etc.) into respective functions.
'
'
' (4): Fixed issue's w/ placing an if-statement within a do-loop & vice versa
'
'
' [ Date 3/1/2022 ]
'
'
' (1): Added "-and", "-or", "-nor", & "-xor" to accepted if-statement operators
'
'
' [ Date 2/28/2022 ]
'
'
' (1): Adjusted the "fndEnvironment" function to now find the applications "runtime block" (Worksheet hosting xlas memory/states).
' ***Changed it in a way to allow editing the name of the worksheet w/o it affecting the current runtime block
'
' (2): Created 3 functions to deal w/ cleaning up an article during parsing (modArtP, modArtQ, modArtS)
' ***Removes parentheses, quotations, & spaces
'
' (3): Seperated loops & conditionals from "lexKey" function & instead put them into their own respective functions
'
' (4): Adjusted do-loops to support nesting if statements
'
'
' [ Date 2/27/2022 ]
'
' (1): Revised "if" statement lexing to check for boolean & seperate initially before setting variable values due to retention
' of the first article following the conditional check during variable parsing.
'
' ***Revised "-eq" boolean operator to except strings
'
' (2): Moved "findbasSaveFormat" function from "lex" module to "xbas" library module (now titled "basSaveFormat")
'
'
' Version 1.0.3
'
' [ Date 2/26/2022 ]
'
' (1): Changed environment variable prefix back to "env" from "e"
'
'
' [ Date 2/24/2022 ]
'
' (1): Added "xlasUpdateEnable" range/memory location for allowing control of the "enableWbUpdates" & "disableWbUpdates" functions
' (Coincides w/ new "--e" & "++e" switches for "xbas v1.0.3" library)
'
'
' [ Date 2/7/2022 ]
'
' (1): Added additional window assignment compatibility for eTweetXL post boxes.
'
' form(31) = Tweet Setup Post Box | form(41) = Tweet Queue Post Box
'
'
'
' [ Date 2/4/2022 ]
'
' (1): Edited "kinSet" function to double check if variable was found and recheck if not before expanding.
'
'  ***Would cause issues when trying to increment a variable such as @x++
'
'
' Version 1.0.2
'
'
' [ Date: 1/31/2022 ]
'
' (1): Added "xlasLink" to "connectWb" function to help confirm the application's connected when opening/starting
'
' (2): Shortened environment variable prefixes from "env" to "e"
'
' (3): Moved connect & disconnect functions to seperate setup module
'
'
'
' Version 1.0.1
'
'
' [ Date: 1/6/2022 ]
'
' (1): Adjusted lexKey to only remove the leading & ending "()" from a declared "kin" variable. Initially would remove all parentheses
' from the article, but this would cause problems in a scenerio where we still needed the parentheses to identify a command.
'
'
'
'
'
' [ Date: 1/5/2022 ]
'
' (1): Created kinExpand function to help w/ expanding (setting the value of a variable) before, & @ runtime.
' ***Working to solve an issue where if a variable is declared w/ a value & needed again for an article such as repl() or input(),
' the variable will still retain the value it was declared as instead of expanding (will work fine when set to nothing, or
' when incrementing the var, or changing it's value to another).
'
'
'
' [ Date: 1/4/2022 ]
'
' (1): Changed name of memory locations to prefix "xlas" as to create lesser ambiquity when connecting to alternative workbooks
'
' (2): Changed "setupWb" function to "connectWb" & created a paired "disconnectWb" to help w/ adding & removing a runtime environment
'
'
' [ Date: 1/3/2022 ]
'
' (1): Added enableWbUpdates function to pair w/ disableWbUpdates to help speed up runtime
' (previously disableWbUpdates left sheet calculations on)
'
' (2): Removed call to syntax error message
'
' (3): Removed enable & disable FlowStrip functions from lex module
'
'
'
' [ Date: 1/2/2022 ]
'
' (1): Added change log, & license information & version #
'
' (2): Created "fndEnvironment" function to find the runtime environment while running xlAppScript
'
' (3): Created "fndJunk" function to remove unwanted data before getting parsed further
'
' (4): Added labeling for most functions & their purpose
'
' (5): Moved script memory locations farther over in the workbook (from MA1-MAA1, MB1-MAB1, etc.)
'
'
End Sub
Private Sub xbas_ChangeLog()
'//=========================================================================================================================
'//
'//         xbas CHANGE LOG
'/\_________________________________________________________________________________________________________________________
'
' Version 1.1.6
'
' [ Date 4/8/2026 ]
'
' (1): Updated copyright, contact, and web information
'
' (2): Changed "Autokit applications" titling to "xlAppScript hosted applications"
'
' Version 1.1.5
'
' [ Date 1/30/2023 ]
'
' (1): Library updated to reflect lexer/parser refactoring changes
'
' (2): Moved "Change Log" to "xlAppScript_logs" module
'
'[ Date: 12/30/2022 ]
'
' (1): Changed variable separator ",#!" --> [,]"
'
'
' Version 1.1.4
'
' [ Date: 12/18/2022 ]
'
' (1): Added "hash()" article for hashing & verifying hashed text strings
'
' Article switches:
'
' -binary1 = create/verify using a basic DIY binary hashing algorithm (unsafe!)
'
'
' [ Date: 12/16/2022 ]
'
' (1): Added "conv32()" article for converting 32-bit signed (-2,147,483,648 <-> 2,147,483,647) integer values to binary (2's complement).
'
' Article switches:
'
' -8bit = set return value as 8-bit binary
' -16bit = set return value as 16-bit binary
' -32bit = set return value as 32-bit binary
'
'
' ***If no switch specified conversion will be returned as 16-bit binary
'
' (Keep in mind returning 32-bit will not print all points of precision to the worksheet!)
'
'
' Version 1.1.3
'
' [ Date: 8/10/2022 ]
'
' (1): Added "shell32()" article for envoking shell32 API library functions
'
' Parameters:
'
' .execute = perform operation on file
' .namespace = get folder information
' .information = get system information
'
' (2): Changed "wastetime()" base article to "delayevent()"
'
' [ Date: 8/11/2022 ]
'
' (1): Added "-xlas" switch to "run()" article for parsing, & exectuing a .xlas script file
'
'
'
' Version 1.1.2
'
' [ Date: 7/26/2022 ]
'
' (1): Fixed issue w/ "lptrim()" & "rptrim()" articles triggering the "ptrim()" article b/c of similiarities in name
'
' (2): Changed "wb()" parameters...
'
'  wb().add --> wb().addsheet
'  wb().addafter --> wb().addsheetafter
'  wb().addbefore --> wb().addsheetbefore
'  wb().new --> wb().newbook
'  wb().nwin --> wb().newwin
'
'
' [ Date: 6/28/2022 ]
'
' (1): Changes to various lexer functions were reflected
'
'
' Version 1.1.1
'
' [ Date: 6/7/2022 ]
'
' (1): Added "sum()" article for performing mathematic operations
'
' Example(s):
'
' @x=sum(1+1); <-- Equals '2'
' @x=sum(1*10); <-- Equals '10'
' @x=sum(((1+1)/1)*5); <-- Equals '10'
'
' Version 1.1.0
'
' [ Date: 5/12/2022 ]
'
' (1): Associating names/syntax for xlAppScript parameters are now ".param" & modifiers "mod."
'
' Changes to variables/titling were made to reflect.
'
' (2): Added "get.", "post.", & "set." modifiers to "me()" article for retrieving, updating, & modifying X & Y coordinates for the currently
' active WinForm.
'
' ---> "-x", "-y", & "-pos" switches for assigning arguments to a specific position or both using "-pos"
'
'
' Version 1.0.9
'
' [ Date: 5/4/2022 ]
'
' (1): Added ".value2" parameter to "rng()", "sel()" & "cell()" articles
'
' This method is slightly more efficient than using ".value" for dealing w/ cell values (not including 'Date' cell types)
'
' (2): Adjusted parsing properties in "rng().value", "sel().value" & "cell().value" & ".value2" variations to accept a null
' value for clearing a cell.
'
'
' Version 1.0.8
'
' [ Date: 4/24/2022 ]
'
'      /``.
'     []
'   _[__]_
'  [______]
' [________] Happy Birthday André! :)
'
'
'
'
' (1): Fixed "q().stop" article not parsing the application extracted to stop
'
'
' [ Date: 4/23/2022 ]
'
' (1): added "--time" switch for getting the current time as a string
'
'
'
' [ Date: 4/21/2022 ]
'
'
' (1): Added extended syntax functionality for "cell()", "rng()", & "sel()" article parameters
'
' These article parameters can now be typed like:
'
' cell(1,1).value @var; or cell(1,1).value(@var);
'
' rng(a1).bgcolor @var1 .fcolor @var2; or rng(a1).bgcolor(@var1).fcolor(@var2);
'
' (2): fixed issues w/ "rng().read" sequence not returning range values
'
' (3): Prefixed library functions w/ "bas"
'
' (4): Added "--enableerr" & "--disableerr" flags for setting the library error level
'
' (5): Added ".move" parameter for "q()" article to query & move files/folders
'
'
'
' Version: 1.0.7
'
'
' [ Date: 4/18/2022 ]
'
' (1): Added "--enableglobal" & "--disableglobal" flags for triggering Global Control variables active/inactive
'
'
' [ Date: 4/17/2022 ]
'
'
' (1): Added "-loose" & "-strict" switch to "q()" article for finding binary or text string match to file/folder path provided if an
' exact not found.
'
' (2): Added "-file" & "-fldr" switches to "q()" article for assiging query search as a file or folder
'
'
' [ Date: 4/15/2022 ]
'
' (1): Added "click()" article for assigning mouse clicks/positioning
'
' Example(s):
'
' click(-double 500,500); <--- this will double click at positon 500,500 on the screen
' click(-leftdown 500,500); <--- this will left click and hold at positon 500,500 on the screen
'
'
' Version: 1.0.6
'
' [ Date: 4/3/2022 ]
'
' (1): Added "read()" article for capturing file text
'
' Article switches:
'
' -all = read all file text
' -nlall = read all file text & seperate each line
' -string = read first occurence of file text w/ this string only
' -xstring = read all file text until first occurence of this string
' -nlstring = read all file text until first occurence of this string & seperate each line
' -line = read this line of file text only
' -xline = read all file text until this line
' -nlline = read all file text until this line & seperate each line
'
' ***Example:
'
' read(-all @filePath); <--- This will read all file text w/o seperating each line
' read(-string @findStr, @filePath); <--- This will read all file text up until a string is found & return that line
' read(-nlstring @findStr, @filePath); <--- This will read all file text up until a string is found & return everything up until then
' w/ each line seperated
' read(-line 5, @filePath); <--- This will read line 5 of the file text
'
'
' [ Date: 4/2/2022 ]
'
' (1): Modified "msg()" article to return input selection value to a variable
'
' (2): Added "incr()" article to allow holding an incremented number based on an assigned value & operation.
'
' ***Example: @var = incr(+1); = 2 | @var = incr(-1) = -2 (These values will stay the same if looped through)
'
'
'
' [ Date: 4/1/2022 ]
'
' (1): Fixed issue w/ "rng().read" command causing an error
'
' (2): Made adjustments to "input()" article so its possible to return user input to a variable
'
' (3): Added ".sel" parameter to "rng()" & "cell()" articles
'
'
'
'
' Version: 1.0.5
'
' [ Date: 3/17/2022 ]
'
' (1): Added "--enablecontain" & "--disablecontain" flags for setting "Local Contain" on/off
' ***Local Contain allows previously used runtime memory addresses to retain their data
'
' [ Date: 3/13/2022 ]
'
' (1): Added "--enablestatic" & "--disablestatic" flags for setting "Local Static" on/off
' ***Local Static allows libraries to stay locked to the current runtime environment session.
'
' [ Date: 3/11/2022 ]
'
' (1): Added base articles "build()", "printer()", & "name()" & respective enhanced
' "app.build()", "app.printer()", & "app.name()" articles
'
' (2): Added "cell()" article for analyzing/modifying cells
'
'
' [ Date: 3/9/2022 ]
'
' (1): Added "fil()" & "dir()" articles & reworked them to include modifiers "mk." & "del."
'
'
' [ Date: 3/8/2022 ]
'
' (1): Added "rng().name" & "sel().name" articles for setting cell names
' ***You can alternatively clear a cell name by leaving it blank
'
' (2): Added "wb().delname" article for deleting a specific cell name
'
' (3): Added "-me" switch to "wb()" article for expanding the workbook name at runtime
'
'
'
' Version: 1.0.4
'
'
' [ Date: 3/2/2022 ]
'
' (1): Added "ptrim()" article to deal w/ removing the starting & ending parentheses of a string.
'
' (2): Added "wastetime()" article to halt the parser while still allowing for user input as oppossed to the "wait()"
' article which halts the parser but also freezes the environment.
'
'***Even though the environment becomes frozen, the "wait()" article is more precise than the wastetime() article for time
'
' wastetime(100) = approx. 1 second | wastetime(1000) = approx. 10 seconds
'
' (3): Updated "winform()" article to include "-last" switch for setting the last WinForm as the current
'
'
' [ Date: 3/1/2022 ]
'
' (1): Added "echo()" & "host()" articles.
'
' echo() = output string using cmd host | host() = output string using vba host (msgbox)
'
' ***echo() supports different window focuses by supplying a value from (0-6) ---> echo(2)(@strToShow); <--- maximized view
'
' (2): Added "conv()" & "xtrim()" articles.
'
' conv(@str, -upper); = convert char or string to a desired case (uppercase, lowercase, etc.)
'
' xtrim(@str, ":"); = remove first & last characters from string by a desired character
'
' (3): Added "lptrim", "rptrim()" & "qtrim()" articles (due to trim() article removing parentheses & quotations during parsing).
'
' lptrim(@str); = remove first & last left facing parentheses from a string
'
' rptrim(@str); = remove first & last right facing parentheses from a string
'
' qtrim(@str); = remove first & last quotes from a string
'
' (4): Added "strrev()" article to reverse strings
'
' strrev(@str); <--- string will be backwards
'
'
' [ Date: 2/28/2022 ]
'
' (1): Updated library to utilize article cleaning function within lexer
'
' (2): Included additional updates made to lexer as well as addition of the "runtime block"
'
'
'
'
' Version: 1.0.3
'
' [ Date: 2/26/2022 ]
'
' (1): Fixed issue w/ "wb().hd" & "wb().sh" articles not parsing w/ correct syntax
'
' [ Date: 2/24/2022 ]
'
' (1): Added "++e"  & "--e" switches for control over enabling/disabling workbook updates during runtime
'
'
' [ Date: 2/23/2022 ]
'
' (1): Removed "colors.txt" file & instead created a function that essentially acts the same way where a color is searched
' for within a list of color name/hex/rgb values.
'
'
' [ Date: 2/10/2022 ]
'
' (1): Fixed an issue w/ key() article leaving leading & ending quotations on the supplied keystroke when parsed.
'
'
' [ Date: 2/8/2022 ]
'
' (1): Changed "WINDOW FORM ARTICLES" labeling to "WINFORM ARTICLES"
'
' (2): Changed "form()" article to "winform()" for readability.
'
'
' [ Date: 2/5/2022 ]
'
' (1): Added "form()" article to "xbas" library for manually setting an application Window
'
'
' Version: 1.0.2
'
'
' [ Date: 1/31/2022 ]
'
' (1) Fixed issue w/ key() article leaving "key" in output
'
'
'
' Version: 1.0.1
'
'
' [ Date: 1/6/2022 ]
'
' (1) Added "ins()" article to find a char/string within another string/variable
'
' ***Will return "TRUE" or "FALSE"  to the assigned variable based on if the char/string searched for was found or not
'
' Syntax: @var = ins(@startPosition, @strToSearch, @strToFind, @compType)
'
' (2) Added "app.run(") or simply the "run()" article to allow for running a module within an opened workbook
'
' ***Currently only supports a single listed (,) argument
'
' Syntax: app.run(moduleName.subName, (arg)); also run(moduleName.subName, arg);
'
' (3) Added pc power articles as well as a copy-paste articles
'
' ***pc.shutdown & pc.reboot articles accept the "-e" switch for auto logging in & bringing up the
' previous session on start-up, & "-t" for setting a timer before shutdown.
'
' pc.copy&() = copy & paste file or folder
' pc.copy&!() = copy, paste, & delete copied file or folder
' pc.shutdown() = shutdown pc
' pc.off() = logoff
' pc.rest() = set pc to rest mode
' pc.reboot() = restart pc
' pc.clr = clear shutdown queue
'
'
'
' [ Date: 1/5/2022 ]
'
' [ Edited: 5/4/2022 ]
'
' (1) Added "repl()" article to replace a value within a string
'
' Syntax: @var = repl(@strToReplace, @strToFind, @strToReplace, @compType)
'
' ***If using 3 arguments like: @var = repl(@strToReplace, @strToFind, @strToReplace) the default comparison method
' will be binary.
'
'
' (2) Added "dfil()" & "ddir()" articles to delete files/folders
'
' ***ddir() will only delete a directory if it's completely empty, so in that instance you could
' use the del.dir() article instead to remove everything.
'
'
'
' [ Date: 1/4/2022 ]
'
' (1): Changed all replace & string check commands for articles to ignore case
'
' ***User can type sh( or SH(, q(, or Q(, etc. & that will be accepted as the same article
'
'
' [ Date: 1/3/2022 ]
'
' (1): Added "q()" article which allows the user to query search either a file or folder (depending on the (.)extension.
' W/ this command you can check for the existance, open, delete, or taskkill a file/folder.
'
' ***q() command is able to search through a total of 3 sets of directories starting from a local drive & base folder.
'
' User only needs to include the drive & base folder.
'
' Examples of drive & base folder:
' C:\Users\ <----
' C:\Windows\ <----
'
' Syntax: q(C:\Users\@fileToQuery).exists (this will prompt whether a file exists or not & it's location)
'
'
' [ Date: 1/2/2022 ]
'
' (1): Added change log, license information & library requirements. Edited library description.
'
' (2): Adjusted "key()" article so it could be split into 7 locations & variables based on a numbered reference (0-6)
' This helped stop collisions w/ VBA, & VBS when the VBA parser ran quicker than the variable was released from the previous run
'
' (You'd likely only come across this issue when trying to run consecutive key() articles w/o using a wait() offset in-between).
'
' ***Numbered references will also be attributed to a corresponding VBA shell mode (0-6) when ran
'
' (3): Shell mode can now be set for sh() article (shell modes (0-6) will correspond w/ the same VBA shell modes (0-6)).
'
' (4): When opening, activating, & "saving as" a workbook, the application environment will be linked to that workbook.
' (This helped w/ navigating back to the original application environment when performing those actions due to the newly
' opened, activated, or saved workbook now being the one activated)
'
' ***Linking is simply just relaying to the same cell (memory location), name ("xlasEnvironment"), & value (current runtime environment) to the currently
' activated workbook.
'
End Sub
Sub xtwt_ChangeLog()
'//=========================================================================================================================
'//
'//         xtwt CHANGE LOG
'/\_________________________________________________________________________________________________________________________
'
'
' Version: 1.0.8
'
' [ Date 1/30/2023 ]
'
' (1): Library updated to reflect lexer/parser refactoring changes
'
' (2): Moved "Change Log" to "xlAppScript_logs" module
'
'
' [ Date 12/4/2022 ]
'
' (1): Minor bug fixes
'
'
' Version: 1.0.7
'
' [ Date 8/1/2022 ]
'
' (1): Added "split.post()" article for splitting a post w/ more than 280 characters into a thread
'
'
' Version: 1.0.6
'
' [ Date 7/26/2022 ]
'
' (1): Changed "sh." modifier to "show." & "hd." modifier to "hide."
'
' sh.home --> show.home
' sh.setup --> show.setup
' sh.post --> show.post
' sh.queue --> show.queue
' sh.apisetup --> show.apisetup
' sh.me --> show.me
'
' hd.home --> hide.home
' hd.setup --> hide.setup
' hd.post --> hide.post
' hd.queue --> hide.queue
' hd.apisetup --> hide.apisetup
' hd.me --> hide.me
'
'
' [ Date 6/28/2022 ]
'
' (1): Changes to various lexer functions were reflected
'
' (2): Added "trim.post()" article for trimming a post fit the Twitter character limit of 280 (starts from the end
'
' (3): Added "load.post()" article for creating a draft from a designated text file
'
' Accepts 2 arguments: arg1 = Name of draft, arg2 = File path
'
' Example: load.post(xName, xPath); <---
'
'
'
' Version: 1.0.5
'
'
' [ Date 6/6/2022 ]
'
' (1): Added "app.hide()" article for hiding the entire application
'
' (2): Fixed an issue where "ADA" articles couldn't be found b/c of there placement in the library
'
'
' [ Date 5/12/2022 ]
'
' (1): Associating names/syntax for xlAppScript parameters are now ".param" & modifiers "mod."
'
' Changes to variables/titling were made to reflect. Various application functions were also renamed.
'
'
' [ Date 5/9/2022 ]
'
' (1): Fixed issue w/ "errLvl" variable causing error b/c of mismatch in data types
'
'
' [ Date 5/5/2022 ]
'
' (1): Various library optimizations (mainly dealing w/ variables)
'
'
' Version: 1.0.4
'
'
'
' [ Date 3/27/2022 ]
'
' (1): Fixed an issue w/ flags not parsing b/c of a missing call to get the Runtime Environment, & Block
'
'
' [ Date 3/4/2022 ]
'
' (1): Changed "-rtime" switch to "-runtime" for clarity
'
'
'
'
' Version: 1.0.3
'
'
' [ Date 2/28/2022 ]
'
' (1): Updated library to utilize article cleaning function within lexer
'
' (2): Included additional updates made to lexer as well as addition of the "runtime block"
'
' Version: 1.0.2
'
' (1): Removed "app.dptrig" article. Initially was needed for performing actions after a profile's data was loaded.
'
' No longer needed due to recent bug fixes/changes.
'
' (2): Added "app.dfilter" article to set draft filter to single/threaded posts.
'
' Example: dfilter(0) = single | dfilter(1) = threaded
'
'
'
' [ Date 2/10/2022 ]
'
' (1): Changed "app.ptrig()" article to "app.dptrig()" for clarity. It still performs the same action of manually setting
' a pre/return code when pulling in specific application data.
'
' dp = Data Pull
'
'
' [ Date 2/8/2022 ]
'
' (1): Changed labeling for "SHOW WINDOW" & "HIDE WINDOW" articles to broader "WINFORM ARTICLES"
'
' (2): Removed "-re" switch from "load.linker()" article. Instead will default to reload if left empty.
'
' Example: load.linker() <---
'
' (3): Added "winform()" & "me()" articles back to xtwt library.
'
' (4): Added "-true" & "-false" switches for boolean operations and parameter values.
'
' (5): Added (,) switch to post() article. Second parameter decides if the post is saved or not using boolean a check.
'
' Example: post(insert text for your post here, -true)   <--- this would save your post to the current focused draft w/
' the text in the first parameter.
'
'
'
' [ Date 2/7/2022 ]
'
' (1): Added "app.dynoffset()" article for activating/deactivating the "Dynamic Offset" option.
'
' Examples: dynoffset(0) = Inactive | dynoffset(1) = Active
'
'
' (2): Added "app.media.show()" article for viewing currently instanced media from either the post or queue window.
'
' ***can be shortened to ---> show.media
'
' (3): Added "app.freeze()" article for pausing/unpausing remaining application automations.
'
'***If used after starting a run, the next automation(s) afterwards will be halted until unfrozen.
' The user will need to trigger a start again after the applications been unfrozen to resume the current run.
'
' Examples: freeze(0) = Unpaused | freeze(1) = Paused
'
'
'
' [ Date 2/6/2022 ]
'
' (1): "profile()" & "user()" articles can now be used across "Profile Setup" & "Tweet Setup" windows
'
' (2): Added "del.profile()" & "del.user()" as well as corresponding "mk.profile()" & "mk.user()" articles to deal w/
' creating & removing profiles/users from an archive.
'
' ***del. parameter supports (*) wildcard for removing all items
' ***Both del./mk. parameters support (,) character for listing items
'
' (3): Changed "SHORT COMMANDS" to "DIRECT ACTION" short for "APPLICATION DIRECT ACTION" or "ADA" to organize articles by "app." use/prefix.
' This may slightly speed up parsing but mainly changed these articles for readabilities sake.
'
' Added "-last" & "-re" switches to "load.linker()" article.
'
' *** -last = reload last connection
'
'
'
'
'
' [ Date 2/5/2022 ]
'
' (1): Created "add.thread()" & "rmv.thread() articles to deal w/ adding & removing threads (supports "*" wildcard for removing all threads)
'
' (2): Added "clr.post" & "clr.linker" for clearing post box & linker
'
' (3): Added "clr.setup"
'
'
' (4): Added "set.ptrig()" article for added control when switching through profiles & needing different users.
'
' ***The profile trigger helps stop instances of code from importing data multiple times during changes to certain window (namely profile/user changes)
'
' set.ptrig(0) = Inactive | set.ptrig(1) = Active
'
'
'
' (5): Updated corresponding setup commands for clarity:
'
' savepost = save.post | savelinker = save.linker | reload = re.linker | load() = load.linker()
'
'
'
' (6): Added "set.sendapi()" article for assigning posts for default or api send
'
' Example(s): set.sendapi(0) = default | set.sendapi(1) = send w/ api
'
'
'
' [ Date 2/4/2022 ]
'
' (1): Removed "form()" article from xtwt library to "xbas" library as it's become a much broader command
'
'
'
' Version: 1.0.1
'
' [ Date: 1/2/2022 ]
'
'(1): Added change log, library requirements, & license information. Edited library description.
'
'(2): Added "LoadLess" functionality which if set will ignore certain loading features the application would
'normally perform when pulling in data to a UserForm window (not capatible w/ eTweetXL versions prior to v1.4.1)
'
'Set "LoadLess" w/ "--l" switch.
'
'Set back to normal w/ "++l" switch.
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
End Sub


