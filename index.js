'use-strict';
const shell = require('node-powershell');
const path = require('path');

function registerShell() {
	return new shell({
		executionPolicy: 'Bypass',
		noProfile: true
	});
}

targetPath("C:\\Users\\eld-longkw\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Pitch.lnk");

async function targetPath(shortcutPath){
    ps = registerShell();
    ps.addCommand(`$sh = New-Object -ComObject WScript.Shell`);
    ps.addCommand(`return $sh.CreateShortcut('${shortcutPath}').TargetPath`);
    ps.invoke()
    .then(output => console.log(output))
    .catch(err => {
        console.error(err);
    });
}