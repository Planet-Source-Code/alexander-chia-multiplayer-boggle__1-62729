Multiplayer Boggle
==================
Chia Yan Sheng Alexander
katana773@yahoo.com

An attempt to create a multiplayer version of the popular game Boggle to play with my friends online.

Features include:
-Ability to customize almost every aspect of the round - eg Time of each round, interval timing, minimum word length, etc.
-A plethora of awards (eg Champion medal, Caterpillar word award, Treasure hunter award, Detective award, etc.) to be awarded to outstanding players every round!
-Custom coded Depth First Search to find if word entered exists in board
-Top 50 words of each round displayed after each round
-In-game Dictionary to display dictionary entry of any of the top 50 words. Just click on the word in the listbox and an explanation will pop up.
-Interesting End-round statistics shown (eg % of total words found, typing speed, words found by each player, etc.)
-Dynamic in-game chat
-A short help tab to explain to beginners how to play this game by Hasbro, as well as to display the awards that could be won
-Gigantic internationally recognized lexicon of words (Enable2k) for nearly every valid word in the dictionary to be recognized by the game.
-Up to 1000 players supported! (If you can find that many)

Useful for programmers attempting to learn tcp network code.
Coded in 3 days in VB 6.0.
Please leave constructive comments and report any bugs encountered. Suggestions are also very welcome.
Thanks!

=================================================================================================
Files each player is required to have:
-BClient.exe
-data\1.wav
-data\2.wav
-data\Def1.dat
-data\Def2.dat

Files the server needs in order to run:
-BServer.exe
-BSolver.exe
-Word.lst

Instructions to run the game:
1.Compile bSolver,bServer, and bClient.
2.***Put BSolver.exe in the same directory as BServer.exe (It generates an output file for bServer.exe)
3.Run BServer.exe and click on 'Start Server'.
4.Run BSolver.exe and click 'Connect'. (This is needed to show the top 50 words of each round and speed up the solving process.)
5.Give each player a copy of the Client runtime files: BClient.exe and the folder data.
6.Have each player run BClient.exe and connect to the server.
7.One player has to type in the command /start in the chat window and the game will start!
8.Enjoy.

P.S. Type in /help in chat to view ingame help
P.S. I usually use /start 2 (long game settings) for my games with friends because the end-game stats are just so incredibly fun to view!

Have fun.
