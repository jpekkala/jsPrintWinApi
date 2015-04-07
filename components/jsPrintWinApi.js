/**
 * Adds a custom property to the Javascript window object that enables access
 * to certain functions in the Windows Print Spooler API.
 *
 * See https://msdn.microsoft.com/en-us/library/windows/desktop/dd162861(v=vs.85).aspx
 * for more information about the API functions.
 */
	
const CLASS_ID = "{E5102ABA-4558-4947-8178-918674704052}";
const CONTRACT_ID = "@jukkapekkala.com/jsPrintWinApi;1";

Ci = Components.interfaces;
Components.utils.import("resource://gre/modules/XPCOMUtils.jsm");
Components.utils.import("resource://gre/modules/ctypes.jsm")
Components.utils.import("resource://gre/modules/devtools/Console.jsm");

//mappings between Windows data types and js-ctypes
const BYTE = ctypes.unsigned_char;
const WORD = ctypes.uint16_t;
const DWORD = ctypes.uint32_t;
const HANDLE = ctypes.voidptr_t;
//we use the unicode versions of functions when available and therefore when
//the API says LPTSTR, the correct type is WCHAR.ptr
const WCHAR = ctypes.jschar;

//this struct is long and we're only interested in one field so we pad everything else
const DEVMODE = ctypes.StructType("DEVMODEW", [
		{"padStart": ctypes.ArrayType(ctypes.uint8_t, 196)},
		{"dmMediaType": DWORD},
		{"padEnd": ctypes.ArrayType(ctypes.uint8_t, 20)}]);

//general printer info
const PRINTER_INFO_1 = ctypes.StructType("PRINTER_INFO_1", [
		{"Flags": DWORD}, 
		{"pDescription": WCHAR.ptr}, 
		{"pName": WCHAR.ptr}, 
		{"pComment": WCHAR.ptr}]);

//per-user DEVMODE 
const PRINTER_INFO_9 = ctypes.StructType("PRINTER_INFO_9", [
		{"pDevMode": DEVMODE.ptr}]);

const PRINTER_DEFAULTS = ctypes.StructType("PRINTER_DEFAULTS", [
		{"pDataType": WCHAR.ptr},
		{"pDevMode": DEVMODE.ptr},
		{"DesiredAccess": DWORD}]);

function jsPrintWinApi() { 
	this.winspool = ctypes.open("winspool.drv");
}

jsPrintWinApi.prototype = {
  classDescription: "jsPrintWinApi",
  classID:          Components.ID(CLASS_ID),
  contractID:       CONTRACT_ID,
  QueryInterface: XPCOMUtils.generateQI([Ci.nsIPrintWinApi]),

	_xpcom_factory: XPCOMUtils.generateSingletonFactory(jsPrintWinApi),

	classInfo: XPCOMUtils.generateCI({classID: Components.ID(CLASS_ID),
																			contractID: CONTRACT_ID,
																			interfaces: [Ci.nsIPrintWinApi],
																			flags: Ci.nsIClassInfo.DOM_OBJECT}),

	/**
	 * Returns the names of printers detected by Windows
	 */
	getPrinters: function() {
		var EnumPrinters = this.winspool.declare(
			"EnumPrintersW", 
			ctypes.winapi_abi, 
			ctypes.bool, //return type
			DWORD, //Flags
			WCHAR.ptr, //Name
			DWORD, //Level
			BYTE.ptr, //pPrinterEnum
			DWORD, //cbBuf
			DWORD.ptr, //pcbNeeded
			DWORD.ptr //pcReturned
		);

		//used in Flags
		const PRINTER_ENUM_LOCAL = 2;

		var name = WCHAR.ptr(0); //NULL, not needed
		var buffer = BYTE.ptr(0); //NULL
		var needed = DWORD(0); //needed buffer size (given by Windows)
		var returned = DWORD(0); //number of PRINTER_INFO structs in buffer

		//call with null buffer to get the needed buffer size
		EnumPrinters(PRINTER_ENUM_LOCAL, name, 1, buffer, 0, needed.address(), returned.address());

		buffer = ctypes.ArrayType(BYTE)(needed.value)

		//call again with correct buffer size
		EnumPrinters(PRINTER_ENUM_LOCAL, name, 1, buffer.addressOfElement(0), needed.value, needed.address(), returned.address());

		//reinterpret the byte buffer as an array of PRINTER_INFO structs
		var printers = ctypes.cast(buffer, ctypes.ArrayType(PRINTER_INFO_1, returned.value))

		//it's not possible to return Javascript arrays through XPCOM (they are wrapped in nsIArray)
		var text = "";
		for(var i = 0; i < returned.value; i++) {
			text += printers[i].pName.readString();
			if(i != returned.value - 1) text += "\n";
		}
		return text;
	},

	/**
	 * Retrieves a handle to a printer by name, or null if not found
   */
	openPrinter: function(printerName) {
		var OpenPrinter = this.winspool.declare(
			"OpenPrinterW", 
			ctypes.winapi_abi, 
			ctypes.bool, //return type
			WCHAR.ptr, //pPrinterName
			HANDLE.ptr, //phPrinter
			PRINTER_DEFAULTS.ptr //pDefault
		);

		var handle = HANDLE();

		if(!OpenPrinter(printerName, handle.address(), PRINTER_DEFAULTS.ptr(0))) {
			return null;
		} 

		return handle;
	},

	/**
	 * Closes the printer handle returned by openPrinter
   */
	closePrinter: function(handle) {
		var ClosePrinter = this.winspool.declare(
			"ClosePrinter",
			ctypes.winapi_abi,
			ctypes.bool, //return type
			HANDLE //hPrinter
		);
		
		return ClosePrinter(handle);
	},

	/**
	 * Gets a PRINTER_INFO struct as a byte buffer, or null.
   * Level specifies which of the nine PRINTER_INFO structs is returned.
   */
	getPrinter: function(handle, level) {
		var GetPrinter = this.winspool.declare(
			"GetPrinterW", 
			ctypes.winapi_abi, 
			ctypes.bool, //return type
			HANDLE, //hPrinter
			DWORD, //Level
			BYTE.ptr, //pPrinter
			DWORD, //cbBuf
			DWORD.ptr //pcbNeeded
		);

		var buffer = BYTE.ptr(0); //NULL
		var needed = DWORD(); //needed buffer size

		//call with null buffer to get the needed buffer size
		GetPrinter(handle, level, buffer, 0, needed.address());

		buffer = ctypes.ArrayType(BYTE)(needed.value);
	
		if(!GetPrinter(handle, level, buffer.addressOfElement(0), needed.value, needed.address())) {
			return null;
		}

		return buffer;
	},

	getMediaTypeNames: function(printerName) {
		var DeviceCapabilities = this.winspool.declare(
			"DeviceCapabilitiesW", 
			ctypes.winapi_abi,
			ctypes.int32_t, //return type should be DWORD (which is defined as unsigned
 											//in Windows Api) but the function returns -1 on failure
			WCHAR.ptr, //pDevice
			WCHAR.ptr, //pPort
			WORD, //fwCapability
			WCHAR.ptr, //pOutput
			DEVMODE.ptr //pDevMode
		);
	
		const DC_MEDIATYPENAMES = 34;	
		const DC_MEDIATYPES = 35;

		var portName = "";

		//NULL values must have the correct declared type in js-ctypes
		var outputPtr = WCHAR.ptr(0); //NULL
		var devmodePtr = DEVMODE.ptr(0); //NULL, not used

		var mediaCount = DeviceCapabilities(printerName, portName, DC_MEDIATYPES, outputPtr, devmodePtr);
		
		if(mediaCount == -1 || mediaCount == 0) return null;
		
		var mediaTypes = ctypes.ArrayType(DWORD)(mediaCount);
		//the function expects a (unicode) string instead of a DWORD array
		outputPtr = ctypes.cast(mediaTypes.address(), WCHAR.ptr);
		DeviceCapabilities(printerName, portName, DC_MEDIATYPES, outputPtr, devmodePtr);
	
		//each name requires a buffer that is 64 characters long
		var mediaTypeNames = ctypes.ArrayType(WCHAR)(mediaCount * 64);
		//pointer to char array is not the same as pointer to char in js-ctypes
		outputPtr = ctypes.cast(mediaTypeNames.address(), WCHAR.ptr);
		DeviceCapabilities(printerName, portName, DC_MEDIATYPENAMES, outputPtr, devmodePtr);

		var result = { __exposedProps__: {}};
		for(var i = 0; i < mediaCount; i++) {
			var name = "";
			//get the name from the buffer, it's not necessarily null-terminated
			for(var j = 0; j < 64; j++) {
				var ch = mediaTypeNames[64*i + j];
				if(ch == "\0") break;
				name += ch;
			}

			result[name] = mediaTypes[i];
			result["__exposedProps__"][name] = "r";	
		}
		
		return result;	
	},


	/**
	 * Get the active devmode as a pointer
   */
	getDevmodePtr: function(handle) {
		var buffer = this.getPrinter(handle, 9);	
		if(!buffer) return null;
	
		//buffer contains PRINTER_INFO_9 and that struct has only one field which is LPDEVMODE
		var devmodePtr = ctypes.cast(buffer, DEVMODE.ptr); 
	
		//per-user DEVMODE can be null, use the global DEVMODE in that case
		if(devmodePtr.isNull()) {
			buffer = this.getPrinter(handle, 8);
			if(!buffer) return null;
			//PRINTER_INFO_8 is identical to PRINTER_INFO_9
			devmodePtr = ctypes.cast(buffer, DEVMODE.ptr);
			if(devmodePtr.isNull()) return null;
		}
		
		return devmodePtr;
	},

	/**
	 * Returns the active media type, which is defined either in the per-user or the global devmode.
	 * If it's not available, the value -1 is returned instead.
   */
	getMediaType: function(printerName) {
		var handle = this.openPrinter(printerName);
		if(!handle) return -1;

		var devmodePtr = this.getDevmodePtr(handle);
		var mediaType	= devmodePtr != null ? devmodePtr.contents.dmMediaType : -1;
		this.closePrinter(handle);
		return mediaType;
	},

	/**
	 * Changes a printer's media type in the per-user DEVMODE.
	 * Returns true on success.
	 */
	setMediaType: function(printerName, mediaType) {
		var handle = this.openPrinter(printerName);
		if(!handle) return false;

		try {
			if(typeof mediaType === "string") {
				var types = this.getMediaTypeNames(printerName);
				if(mediaType in types) mediaType = types[mediaType];
			}
			if(typeof(mediaType) !== "number") return false;
	
			var devmodePtr = this.getDevmodePtr(handle);
			if(devmodePtr == null) return false;

			devmodePtr.contents.dmMediaType = mediaType;

			var SetPrinter = this.winspool.declare(
				"SetPrinterW", 
				ctypes.winapi_abi, 
				ctypes.bool,
				HANDLE, //hPrinter
				DWORD, //Level
				BYTE.ptr, //pPrinter
				DWORD //command
			);	

			//PRINTER_INFO_9 has only one field and that is a pointer to DEVMODE
			var bufferPtr = ctypes.cast(devmodePtr.address(), BYTE.ptr);
			return Boolean(SetPrinter(handle, 9, bufferPtr, 0));

		} finally {
			this.closePrinter(handle);
		}
	}

};

var components = [jsPrintWinApi];
if ("generateNSGetFactory" in XPCOMUtils)
  var NSGetFactory = XPCOMUtils.generateNSGetFactory(components);  // Firefox 4.0 and higher
else
  var NSGetModule = XPCOMUtils.generateNSGetModule(components);    // Firefox 3.x
