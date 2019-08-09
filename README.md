# NOAAUpdater
An automation utility designed to scrap online data to update a local Excel file and an online SQL server.

### Information
NOAAUpdater is a program designed to automatically retrieve weather data from the NOAA website and update both NOAA_Weather.xlsx and the NOAA_Weather table on the **//REDACTED//** database.

To initiate the program, make sure NOAA_Weather.xlsx is not currently open, and run the NOAAUpdater.exe executable. Follow the prompts that appear on screen.

The program takes approximately 30 - 60 seconds per missing month of data to run depending on your computer, and it has no loading bar. Do not run the program again before the program completes its operation. You will be notified when the program finishes.

It is recommended to run this program every day to prevent data gaps.

### Warnings
Do not open NOAA_Weather.xlsx while this program is executing. The weather.NOAA_Weather table on the SQL database may be unable to take SQL queries while the utility is running.

### Notice of Publication
This utility is designed to be used by Con Edison only. Certain information has been redacted from the source code and associated files to provide corporate security. The program as-is is not runnable without replacing "//REDACTED//" with valid entries within the source file.
