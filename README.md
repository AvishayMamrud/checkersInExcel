# checkersInExcel
 - open excel file with .xlsb extension (any macro compatible extension is fine)
 - make sure that Macro Commands are enabled in your excel (instructions at https://www.youtube.com/watch?v=vtTa8y6i_rM . **not my video**)
 - press ALT + F11 to reach the vba 'section'
 - select the first and only sheet at the top of the project panel (unless there are other workbook open. then, you should look for your new workbook by name and select the sheet.
 - copy the content of CheckersInExcel.txt and simply paste it there (the screen you found in the last step)
 - save, close and open the workbook
 - go to view -> macro command -> sheet1.newGame and then press 'run'
 - start playing. white starts. (you can toggle turn with the command sheet1.toggleUserTurn)
 - enjoy!


 - in the future i hope to make it posible to play in multi-player mode as a shared workbook.
 - there are a few options i did not enforce. e.g. the player can 'eat' only one piece each turn

