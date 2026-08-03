'use strict';

const { spawn } = require('child_process');

const WORKBENCH_URL =
    'https://yorpas.sharepoint.com/_layouts/15/workbench.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true';

const child = spawn('heft', ['start', '--clean', '--nobrowser'], {
    shell: true,
    stdio: ['inherit', 'pipe', 'pipe']
});

let opened = false;

child.stdout.on('data', (data) => {
    const text = data.toString();
    process.stdout.write(text);
    if (!opened && text.includes('Waiting for changes')) {
        opened = true;
        // "Start-Process <url>" (ShellExecute via the registered URL protocol
        // handler) silently does nothing on machines where that OS-level
        // association is broken/stale — a common Windows issue that shows no
        // error but also opens no window. Resolving Edge's real .exe path from
        // the registry and launching it directly (bypassing URL-association
        // resolution entirely) is what actually works reliably. The URL is
        // passed through an env var, never interpolated into the PowerShell
        // command string, so "&" in the query string can't be misparsed.
        spawn(
            'powershell',
            [
                '-NoProfile',
                '-Command',
                "$p = (Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\msedge.exe' " +
                    "-ErrorAction SilentlyContinue).'(default)'; " +
                "if (-not $p) { $p = (Get-ItemProperty 'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\msedge.exe' " +
                    "-ErrorAction SilentlyContinue).'(default)' }; " +
                'if ($p) { Start-Process -FilePath $p -ArgumentList $env:OPEN_WHEN_READY_URL } ' +
                'else { Start-Process $env:OPEN_WHEN_READY_URL }'
            ],
            {
                stdio: 'ignore',
                detached: true,
                env: { ...process.env, OPEN_WHEN_READY_URL: WORKBENCH_URL }
            }
        ).unref();
    }
});

child.stderr.on('data', (data) => process.stderr.write(data));

child.on('exit', (code) => process.exit(code ?? 0));
process.on('SIGINT', () => child.kill('SIGINT'));
