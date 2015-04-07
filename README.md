jsPrintWinApi
=============

jsPrintWinApi is a Firefox extension that allows the media type (e.g. "Plain Paper", "Thick", "Envelope" etc.) to be changed in Javascript when printing documents. This extension is designed to be used in conjuction with [jsPrintSetup](http://jsprintsetup.mozdev.org/reference.html) and does not provide capabilities that already exist in jsPrintSetup.

Install
-------------

To install jsPrintWinApi, download this GitHub project as a zip file and rename the file extension from .zip to .xpi. Then open the file in Firefox by pressing Ctrl+O or by dragging and dropping the file inside Firefox.

Usage
-------------

The extension adds a new property 'jsPrintWinApi' to the Javascript window object. The following functions are available:

getMediaTypeNames(printerName):
	Returns the available media types for a printer as an object whose keys are media type names and values are their numeric IDs.  

getMediaType(printerName):
	Returns the current media type as a number (media type numbers are printer-specific).

setMediaType(printerName, mediaType):
	Changes the per-user media type for a printer. The media type can be either a string or the numeric ID.	
	
Technical details
-----------------

jsPrintSetup does not let the media type to be changed because media types are driver-specific and Firefox does not expose any way (at least not any easy way) to change that setting programmatically.

jsPrintWinApi gets around this by calling the Windows Print Spooler API directly.  More specifically, the media type that Firefox uses for printing can be changed my modifying the [Per-User DEVMODE](https://msdn.microsoft.com/en-us/library/windows/desktop/dd162798%28v=vs.85%29.aspx) which is stored in the user's registry. Changing it does not require administrator rights and the value only affects the current user.

Normally this value can be changed by opening the Printer Properties dialog. This is not a satisfactory solution when there are a lot of documents that need to be automatically printed on different media types, and without any user intervention.

License
-----------------

This project is released under the MIT License. Please feel free to use, modify and merge the code freely.
