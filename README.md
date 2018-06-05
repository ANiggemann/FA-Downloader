# Moved to GitLab: [FA-Downloader](https://gitlab.com/ANiggemann/ANI-Camera-Remote)


# FA-Downloader
FA-Downloader is a VBScript to automatically download images via WiFi from a Toshiba FlashAir SD-Card inside a digital camera. Each new image is transfered within 2 to 5 seconds to the computer.

FA-Downloader Features
-
* Supports all 4 Versions of the Toshiba FlashAir SD-Card and the Transcend WIFI Card
* Automatic detection of the card type
* Transfers automatically all files of certain file types from the card to the local disc of the computer
* If there configuration contains no file type list all files will be transfered
* By default JPG is set, so only JPG files will be transfered
* Traverses all images directories below the DCIM directory
* Already transfered files will not be transfered again. Detection based on filename
* In case of disconnections the transfer will automatically resume
* Ignors the TSB directory 
* No visible window, background process
* Runs indefinitly, can be stopped by a signal file in the TEMP directory
* Message at start and end of the script
* Can be configured to start a CMD file (batch) after each transfer
* Log configuration in differenz verbosity levels

Installation
-
The File FA_Downloader.vbs can be stored in an arbitrary directory and executed there.  Command line parameters see below.

Configuration
-
In the header of the VBScript file you can find a section that reads "Configuration". Here the 4 most important settings can be set.
* CARD = "flashair"
  * This is the name or the IP of the card
* localfolder = "C:\PHOTOS\FLASHAIR\"
  * This is the target drectory for all transfers.  The trailing backslash is mandatory
* filetypes = "JPG"
  * List of file types, separated by comma
* loglevel = 0
  * Verbosity of the log
    * 0 = no log
    * 1 = transfers only
    * 2 = extended log
    * The log can be found in the file FA_Download.log in the TEMP directory

Command line parameters
-
The first 3 configuration settings can be overriden by command line parameters. Example:
* FA_Downloader.vbs "my_flashair" "X:\images\Import\" "JPG,CR2"
* In this example JPGs and CR2 will be transfered.

Stop the script
-
FA_Downloader.vbs stores a certain file in the TEMP directory.  As soon as this file is deleted, the script will end. Already running transfers will not be interrupted. The filename reads "FA_Downloader_", then a number and the file type is ".tmp".

Hints
-
* If there is only the TSB directory on the card any new file will be stored there. FA-Downloader ignores this directory. Best practice is to define the DCIM directory manually on the card (via card reader).
* The WiFi range of this SD-Cards is fairly limited.  The range can be extended by using a pocket router that runs from rechargeable batteries. The computer and the card can be connected to this router that should be carried near the camera (shirt pocket etc.). The FlashAir card should be configured for APPMODE=5.
* Since the Transcend WIFI card is discontinued, the Toshiba FlashAir 4 is now the best choice.
* How long takes a transfer?  With the FlashAir 4 transfer times of 4 seconds per JPG are possible. Set the camera to RAW + smallest JPG. In burst mode 2 seconds are possible (depending on the camera).
