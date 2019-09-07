// Simple DLL Loader
var manifest = '<?xml version="1.0" encoding="UTF-16" standalone="yes"?> <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"> 	<assemblyIdentity type="win32" name="DynamicWrapperX" version="2.2.0.0"/> 	<file name="MessageBox64.dll">     	<comClass         	description="DynamicWrapperX Class"         	clsid="{89560001-A714-4a43-912E-978B935EDCCC}"         	threadingModel="Both"         	progid="DynamicWrapperX"/> 	</file>  </assembly>';

new ActiveXObject('WScript.Shell').Environment('Process')('TMP') = 'C:\\Users\\Research\\Documents\\MessageBox-master\\bin';

var ax = new ActiveXObject("Microsoft.Windows.ActCtx");

ax.ManifestText = manifest;

var dwx = ax.CreateObject("DynamicWrapperX");

//Pops Message Box From DLL Main()
