
# aedifion Excel plug-in

Welcome to the aedifion Excel plug-in!

The aedifion Excel plug-in allows you to one-click-import time series data stored at the aedifion data platform directly into your Microsoft Excel.

**Prerequisites:** 

A windows computer running a newer version of Microsoft Excel. Installations of Excel on e.g. Apple computers currently do not support the installation of Excel plug ins.
Further the installation of the aedifion Excel plugin does not require admin rights on your computer. Please inform your IT-admin about the installation, anyway.

**Installation:** 

Execute the file setup.exe. The plug-in will be directly installed to your Excel plug-ins. When the installation is done, open an empty workbook in your Excel. You can find the aedifion Excel plug-in as a tab next to "file", "home", "view" etc. as "aedifion". Click on the aedifion tab and go to "Settings" ("Einstellungen" in German). Please tick the check-box “Developer Options” ("Entwickleroptionen" in German) and insert the aedifion API URL: https://api.aedifion.io

**How to use:** 

You can find the aedifion Excel plug-in as an extra tab on the layer where you can find file, how, view etc.. Choose the aedifion tab and you are right inside the aedifion plug-in. 
The preselected language is German. In case you want to switch to English go to "Einstellungen" and select English.
Login in with your aedifion user credentials at "Login" ("Anmeldung" in German) and you are set to go. In order to request time series choose the "Read Data" ("Daten lesen" in German) option. Choose your "Building" ("Gebäude" in German) first so that the list of data points from this building will be loaded. Pick one or more data points you want to get data from and configure the other options according to your wishes. Press start to load the data with the configured options. As soon as the data is loaded into Excel the option "Diagram" is available to plot the received data right away.
So far as a brief introduction in the aedifion Excel plug-in. Feel free to play around and experience the tool as much as you like.

**Trouble shooting:**

401 error on login - Please check your login credentials. Another source of error could be the server URL queried. Please go to "Settings" and check "Developer Options". The server URL will open up. Please check, if the URL is correct.

time-out error - Please check your network settings and report this error to support@aedifion.com

Remote name not found - Please go to "Settings" and check "Developer Options". The server URL will open up. Please check, if the URL is correct.
