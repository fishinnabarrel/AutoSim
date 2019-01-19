## AutoSim 1.3

### Introduction:
AutoSim is an [AutoIt](https://www.autoitscript.com/site/) script that automates the tedious button mashing required to simulate multiple seasons with [Diamond Mind Baseball](http://www.diamond-mind.com/) and import the results into the DMB Encyclopedia.

### Getting started:
AutoSim will run from any disk location and does not need to be installed.  The program looks for the DMB11 and DMB11 Encyclopedia executable files in their default installation folders -- "C:\dmb11" and "C:\dmbenc11".  If you have DMB or the encyclopedia installed somewhere other than the defaults you can easily change the program folders through the **Options** menu on the AutoSim main menu.

Before starting AutoSim make sure that you have made any desired league or organization adjustments from within DMB and that you have created a new encyclopedia database to store your results in the DMB Encyclopedia.


### Running AutoSim:
AutoSim will determine the active database and encyclopedia and display them.  If you would like to change one of these values you must change the active database or the active encyclopedia from within the respective program.

Open the **Options** menu and select **Refresh active database** to update the display once you have made the change(s).  Before beginning you can change the number of seasons that you would like to simulate and the season that you would like to begin counting from by entering new values in the appropriate boxes.  

Click the **Start** button to begin processing a batch of seasons.  The program performs the following activities each cycle: restart the season, simulate one complete season, and then import the season into the encyclopedia. 
A recap message will be displayed once the batch is finished. The AutoSim program will close all open DMB and Encyclopedia windows before beginning and again after completing a batch.

If you have the options to save boxscores and/or Game-by-game statistics in your season's league or organization options it will significantly slow things down. 

Because AutoSim is basically a better button pusher, once processing has begun, a batch cannot be paused or stopped until it has completed.  If you need to exit AutoSim before this, press **Ctrl + Alt + x**.  This will close AutoSim, but leave DMB and the Encyclopedia running without any further automation.  

It is possible to interact with DMB and the Encyclopedia while a batch is running, but it is not recommended as it may cause the automation to stall or stop.  If this happens it is sometimes possible to give it a nudge by manually performing the step that was missed.  Note that closing DMB or the Encyclopedia will cause the automation to stop completely, even if you open them again.


### Summary of automation:
- Close all currently open copies of DMB and DMB Encyclopedia.
- Open one copy of DMB and one copy of DMB Encyclopedia.
- Restart season.
- Simulate season.
- Import season into Encyclopedia. 
- Repeat (Restart, Simulate, Import)
- Display recap message.
- Close all currently open copies of DMB and DMB Encyclopedia.


David Pyke
January 18, 2019
