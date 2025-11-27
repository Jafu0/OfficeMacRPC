const RPC = require('discord-rpc');
const readline = require('readline');

// Client IDs for each app
const CLIENT_IDS = {
    'PowerPoint': '1443646079957729320',
    'Word': '1443649127169658950',
    'Excel': '1443648941118717972'
};

let rpc = null;
let currentClientId = null;
let lastStatus = null;
let showFileName = true;
let startTimestamp = null;
let currentDoc = null;

function getClientIdForApp(appName) {
    if (!appName) return null;
    if (appName.includes('PowerPoint')) return CLIENT_IDS.PowerPoint;
    if (appName.includes('Word')) return CLIENT_IDS.Word;
    if (appName.includes('Excel')) return CLIENT_IDS.Excel;
    return null;
}

// Connect to Discord with a specific Client ID
async function connectRpc(clientId) {
    if (rpc) {
        if (currentClientId === clientId) return;
        try {
            await rpc.destroy();
        } catch (e) { /* ignore */ }
        rpc = null;
    }

    if (!clientId) return;

    currentClientId = clientId;
    rpc = new RPC.Client({ transport: 'ipc' });

    rpc.on('ready', () => {
        console.log(JSON.stringify({ type: 'discord_status', connected: true, clientId: clientId }));
        if (lastStatus) {
            updatePresence(lastStatus);
        }
    });

    try {
        await rpc.login({ clientId: clientId });
    } catch (err) {
        console.log(JSON.stringify({ type: 'discord_status', connected: false, error: err.message }));
        rpc = null;
        currentClientId = null;
    }
}

function updatePresence(status) {
    if (status.clear || status.idle) {
        if (rpc) {
            rpc.clearActivity().catch(() => { });
        }
        startTimestamp = null;
        currentDoc = null;
        return;
    }

    const appName = status.app;
    const docName = status.doc;

    const targetClientId = getClientIdForApp(appName);

    if (targetClientId !== currentClientId) {
        connectRpc(targetClientId);
        return;
    }

    if (!rpc) return;

    // 4. Time Tracking
    if (docName !== currentDoc) {
        startTimestamp = new Date();
        currentDoc = docName;
    }

    // 5. Construct Presence Data

    let detailsText = "";
    let largeImageKey = "";
    let largeImageText = "";

    if (appName.includes('PowerPoint')) {
        detailsText = "Editing a Presentation";
        largeImageKey = "powerpoint";
        largeImageText = "Microsoft PowerPoint";
    } else if (appName.includes('Excel')) {
        detailsText = "Editing a Spreadsheet";
        largeImageKey = "excel";
        largeImageText = "Microsoft Excel";
    } else if (appName.includes('Word')) {
        detailsText = "Editing a Document";
        largeImageKey = "word";
        largeImageText = "Microsoft Word";
    }

    // Override with Custom Status if present
    if (status.customStatus && status.customStatus.length > 0) {
        detailsText = status.customStatus;
        // Log that we are using custom status
        console.log(JSON.stringify({ type: 'debug_log', message: "Using Custom Status: " + status.customStatus }));
    }

    let stateText = "";
    if (showFileName && docName && docName.length > 0) {
        stateText = docName;
    } else {
        stateText = "In a File";
    }

    rpc.setActivity({
        details: detailsText,
        state: stateText,
        startTimestamp: startTimestamp,
        largeImageKey: largeImageKey,
        largeImageText: largeImageText,
        smallImageKey: 'edit',
        smallImageText: 'Editing...',
        instance: false,
    }).catch((err) => {
        console.log(JSON.stringify({ type: 'discord_status', connected: false, error: err.message }));
    });
}
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
    terminal: false
});

rl.on('line', (line) => {
    try {
        const data = JSON.parse(line);

        if (data.type === 'config') {
            if (typeof data.showFileName === 'boolean') {
                showFileName = data.showFileName;
                if (lastStatus) updatePresence(lastStatus);
            }
        } else if (data.type === 'status') {
            // Debug: Echo back what we received
            console.log(JSON.stringify({ type: 'debug_log', message: JSON.stringify(data) }));

            lastStatus = data;
            updatePresence(data);
        }
    } catch (e) {
        // Ignore parse errors
    }
});
