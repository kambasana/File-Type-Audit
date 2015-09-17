# filetypeaudit
A File type audit tool that allows you to view the file types of folder without manually counting them. This allows you to analyse a large directory and review the results for investigation or for source code review.
Standard IIS install required or environment that can run ASP.net
index.asp is set for 1024 by 768 screen res.

index2.asp is set for 1280 by 1024 screen res.

If you need to use the higher res window, replace 
the index.asp with index2.asp and rename to index.asp.

Place all code in the Code folder and run the index.asp

If you prefer to have a straight view and not use the window system, you can run
castcode from the main folder. Please note that you must have run the castcount.asp first
and not the showExtension.asp as Castcount pipes the session data to showextension.