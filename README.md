# aedifion Excel Add-in

The aedifion Excel Add-in allows you to one-click-import timeseries data stored at the aedifion data platform directly into your Microsoft Excel.

## Use case
Excel is a wide spread tool among engineers. Either to create plots, examine timeseries data, or to execute scripts e.g. to determine KPIs. Thus the aedifion Excel Add-in integrates to existing processes and scripts, allowing for performance enhancements and new services in Excel.

## Prerequisites

* **Windows** operating system running a newer version of **Microsoft Excel**.

* **Apple** and its **iOS** operating system **do not enable installation of Excel Add-ins**. Thus the aedifion Excel Add-in cannot be installed on Apple devices.

## Installation

* **Download all files as .zip folder** via *Clone or download* -> *Download ZIP*. If possible, use program *7-Zip* to unzip the files. **Do not git clone** the repository, this will result in installation errors.
* **Admin rights** on your computer are **not required**. Please inform your IT-admin about the installation anyway.
* Execute the file **setup.exe**.
* Publisher check appears. Confirm with a click on Install. The Add-in will be installed to your Excel.
	![](/images/allow_installation.PNG)
	* In case of an error check our **FAQs**.

## How to use
* Open an empty workbook in your Excel.
* Click the ***aedifion*** **tab** next to *file*, *home*, *view* etc. 
	* **Change language** (default *German*): *Settings (Einstellungen)* -> *Language (Sprache)*
	* **Change API server** (default https://api.aedifion.io): *Settings (Einstellungen)* -> click the check-box *Developer Options* -> insert the *API server address* of your project. You received the API server address as part of the *technical clearing and accounts* email.
	* **Login**: Click *Login (Anmelden)* -> insert your user credentials
	* **Data services**: The click-able icons.
		* **Read Data (Daten lesen)**: This option allows to import timeseries data into Excel. You can choose from multiple datapoints of different buildings (aka. projects), several interpolation methods, diagram only, and flexible time settings. Press *Start* as soon as you want to query your configuration.
		* **Diagram (Diagramm)**: This option is enabled as soon as timeseries data is imported. It plots the selected timeseries data to a diagram.
		* **Add Curve (Kurve hinzufügen)**: This option allows to add a further timeseries to an existing plot. It also allows to plot timeseries of different buildings (aka. projects) into the same plot.
		* **Data Points (Datenpunkte)**: This option allows to import a list of all datapoints of a project into Excel. 

# FAQs
This section collects general Frequently Asked Questions. If your question is not answered here, please contact support@aedifion.com .

## Installation

### Certificate issue

![](/images/certificate_issue.PNG)

Depending on the security settings of your computer, the installation of Excel Add-ins from internet sources is blocked.

* *Option 1*: Right click on the file *setup.exe*. Click the check-box *Unblock* (see screen-shot below). 

	![](/images/certificate_issue_solved.PNG)

	In case this option is not available on your computer, use *option 2*.

* *Option 2*: Adjust windows registry for [*ClickOnce trust promt*](https://docs.microsoft.com/en-us/visualstudio/deployment/how-to-configure-the-clickonce-trust-prompt-behavior?view=vs-2017).
		We advice you to use tools like [Trust Promt](https://www.smartlux.com/software/trust-prompt-tool/) to do so. If you do not want to use Trust Prompt, try *option 3*.
	* Download and execute (admin rights required) Trust Prompt.

	![](/images/Trust_Promt.PNG)
	* Click *Read from registry* and remember the settings. Create a new entry, if there is none available.
	* During installation of the aedifion Excel Add-in set *Internet* to *Enable* and click *Write to registry*.
	* After installation reset the settings to their initial state and click *Write to registry*.
* *Option 3*: Adjust windows registry for [*ClickOnce*](https://docs.microsoft.com/en-us/visualstudio/deployment/how-to-configure-the-clickonce-trust-prompt-behavior?view=vs-2017) manually.
	* Click on *Windows start* and execute *regedit.exe*.
	* 64-bit systems navigate to *\HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel*
	* 32-bit systems navigate to *\HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel*
	* Create *string*: *Internet*
	* Double click on the string and set its value to *Enable*.
	* Verify with *OK* and rerun installation. (In case the 64-bit fix does not work, try the 32-bit option.)

### Hash issue

![](/images/hash_issue.png)

This error occurs e.g. if the add-in is downloaded via git clone. Please download the add-in installation files as a .zip folder and retry the installation.

## Authentication
### 401 error on login
Please check your login credentials e.g. by logging in to the aedifion front-end with your credentials. If your credentials work on the front-end, check the API server address via *Settings* -> *Developer Options*. The server URL will open up. Please check, if the URL is the one you received in the _technical clearing and accounts_ email.

### Remote name not found
Check the API server address via *Settings* -> *Developer Options*. The server URL will open up. Please check, if the URL is the one you received in the _technical clearing and accounts_ email.

### Time-out error
Please check your network settings and make sure your computer has sufficient internet connection. If the error occurs during good internet connectivity, please contact support@aedifion.com .
