import Cocoa
import Foundation
import ServiceManagement

let NODE_SCRIPT_PATH = "index.js"
let CHECK_INTERVAL: TimeInterval = 5.0
let IDLE_THRESHOLD: TimeInterval = 300.0 // 5 minutes

// States
var statusBarItem: NSStatusItem!
var nodeProcess: Process?
var checkTimer: Timer?
var idleTime: TimeInterval = 0
var showFileName: Bool = true
var customStatus: String? = nil

class AppDelegate: NSObject, NSApplicationDelegate, NSMenuDelegate {
    var lastNodeLog: String?
    
    let appConfig: [String: String] = [
        "com.microsoft.Word": "Microsoft Word",
        "com.microsoft.Excel": "Microsoft Excel",
        "com.microsoft.Powerpoint": "Microsoft PowerPoint"
    ]

    
    func applicationDidFinishLaunching(_ aNotification: Notification) {
        customStatus = UserDefaults.standard.string(forKey: "CustomStatus")
        
        statusBarItem = NSStatusBar.system.statusItem(withLength: NSStatusItem.squareLength)
        if let button = statusBarItem.button {
            if let iconPath = Bundle.main.path(forResource: "menubar_icon", ofType: "png"),
               let image = NSImage(contentsOfFile: iconPath) {
                image.isTemplate = true // Allows it to adapt to dark/light mode
                image.size = NSSize(width: 18, height: 18)
                button.image = image
            } else {
                button.title = "DO"
            }
        }
        
        setupMenu()
        startNodeProcess()
        startPolling()
    }
    
    func setupMenu() {
        let menu = NSMenu()
        menu.delegate = self
        
        let showFileItem = NSMenuItem(title: "Show File Name", action: #selector(toggleShowFile), keyEquivalent: "f")
        showFileItem.target = self
        menu.addItem(showFileItem)
        
        let setStatusItem = NSMenuItem(title: "Set Custom Status...", action: #selector(setCustomStatus), keyEquivalent: "s")
        setStatusItem.target = self
        menu.addItem(setStatusItem)
        
        let clearStatusItem = NSMenuItem(title: "Clear Custom Status", action: #selector(clearCustomStatus), keyEquivalent: "")
        clearStatusItem.target = self
        menu.addItem(clearStatusItem)
        
        menu.addItem(NSMenuItem.separator())
        
        let startupItem = NSMenuItem(title: "Launch on Startup", action: #selector(toggleStartup), keyEquivalent: "")
        startupItem.target = self
        menu.addItem(startupItem)
        
        menu.addItem(NSMenuItem.separator())
        
        let debugItem = NSMenuItem(title: "Debug Info", action: #selector(showDebugInfo), keyEquivalent: "d")
        debugItem.target = self
        menu.addItem(debugItem)
        
        menu.addItem(NSMenuItem.separator())
        menu.addItem(NSMenuItem(title: "Quit", action: #selector(quit), keyEquivalent: "q"))
        
        statusBarItem.menu = menu
        updateMenuState()
    }
    
    @objc func toggleShowFile(_ sender: NSMenuItem) {
        showFileName.toggle()
        updateMenuState()
        sendToNode(json: ["type": "config", "showFileName": showFileName])
    }
    
    @objc func setCustomStatus() {
        let alert = NSAlert()
        alert.messageText = "Set Custom Status"
        alert.informativeText = "Enter the text to display instead of 'Editing a Document':"
        alert.addButton(withTitle: "OK")
        alert.addButton(withTitle: "Cancel")
        
        let container = NSView(frame: NSRect(x: 0, y: 0, width: 250, height: 50))
        
        let input = NSTextField(frame: NSRect(x: 0, y: 26, width: 250, height: 24))
        input.stringValue = customStatus ?? ""
        container.addSubview(input)
        
        let checkbox = NSButton(frame: NSRect(x: 0, y: 0, width: 250, height: 20))
        checkbox.setButtonType(.switch)
        checkbox.title = "Save status across restarts"
        let savedStatus = UserDefaults.standard.string(forKey: "CustomStatus")
        checkbox.state = (savedStatus != nil || customStatus == nil) ? .on : .off
        container.addSubview(checkbox)
        
        alert.accessoryView = container
        
        let response = alert.runModal()
        if response == .alertFirstButtonReturn {
            let text = input.stringValue.trimmingCharacters(in: .whitespacesAndNewlines)
            customStatus = text.isEmpty ? nil : text
            
            if checkbox.state == .on {
                if let status = customStatus {
                    UserDefaults.standard.set(status, forKey: "CustomStatus")
                } else {
                    UserDefaults.standard.removeObject(forKey: "CustomStatus")
                }
            } else {
                UserDefaults.standard.removeObject(forKey: "CustomStatus")
            }
            
            updateMenuState()
        }
    }
    
    @objc func clearCustomStatus() {
        customStatus = nil
        UserDefaults.standard.removeObject(forKey: "CustomStatus")
        updateMenuState()
    }
    
    func updateMenuState() {
        if let menu = statusBarItem.menu {
            if let item = menu.item(withTitle: "Show File Name") {
                item.state = showFileName ? .on : .off
            }
            if let item = menu.item(withTitle: "Launch on Startup") {
                item.state = isLaunchItemPresent() ? .on : .off
            }
            if let item = menu.item(withTitle: "Clear Custom Status") {
                item.isHidden = (customStatus == nil)
            }
        }
    }
    
    @objc func toggleStartup(_ sender: NSMenuItem) {
        let currentState = sender.state == .on
        let newState = !currentState
        
        if setLaunchOnStartup(enabled: newState) {
            sender.state = newState ? .on : .off
        }
    }
    
    @objc func showDebugInfo() {
        let anyInput = CGEventType(rawValue: UInt32.max)!
        let idleTime = CGEventSource.secondsSinceLastEventType(.combinedSessionState, eventType: anyInput)
        
        let isTrusted = AXIsProcessTrusted()
        
        var info = "Debug Info:\n\n"
        info += "Version: \(Bundle.main.infoDictionary?["CFBundleShortVersionString"] as? String ?? "Unknown")\n"
        info += "Accessibility Trusted: \(isTrusted) " + (isTrusted ? "✅" : "❌") + "\n"
        if !isTrusted {
            info += "(Please grant Accessibility permissions in System Settings)\n"
        }
        info += "Idle Time: \(String(format: "%.1f", idleTime))s (Threshold: \(IDLE_THRESHOLD)s)\n"
        info += "Custom Status: \(customStatus ?? "None")\n"
        info += "Last Reported Bundle: \(lastReportedBundleId ?? "None")\n"
        info += "Node Process Running: \(nodeProcess?.isRunning == true)\n"
        info += "Discord RPC: \(discordConnected ? "Connected ✅" : "Disconnected ❌")\n"
        let safeLog = (lastNodeLog ?? "None").replacingOccurrences(of: "\n", with: " ").prefix(100)
        info += "Last Node Log: \(safeLog)\n"
        if let err = discordError {
            info += "RPC Error: \(err)\n"
        }
        info += "\n"
        
        info += "Detected Apps:\n"
        
        let runningApps = NSWorkspace.shared.runningApplications
        for (bundleId, name) in appConfig {
            if let app = runningApps.first(where: { $0.bundleIdentifier == bundleId }) {
                let isActive = app.isActive
                // Try to get document name and error
                let (docName, error) = getDocumentName(bundleId: bundleId, appName: name)
                info += "- \(name): Running (Active: \(isActive))\n"
                info += "  Doc Name: '\(docName)'\n"
                if let err = error {
                    info += "  Error: \(err)\n"
                }
            } else {
                info += "- \(name): Not Running\n"
            }
        }
        
        let alert = NSAlert()
        alert.messageText = "Office Rich Presence Debug"
        alert.informativeText = info
        alert.addButton(withTitle: "OK")
        if !isTrusted {
            alert.addButton(withTitle: "Open Settings")
        }
        
        let response = alert.runModal()
        if response == .alertSecondButtonReturn {
            // Open System Settings > Privacy & Security > Accessibility
            // The URL scheme for this changed in macOS Ventura, but this usually works to open Privacy
            if let url = URL(string: "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility") {
                NSWorkspace.shared.open(url)
            }
        }
    }
    
    @objc func quit() {
        nodeProcess?.terminate()
        NSApplication.shared.terminate(self)
    }
    
    // --- Node Process Management ---
    
    var discordConnected = false
    var discordError: String?
    
    func startNodeProcess() {
        // Look for node in standard paths
        let paths = ["/usr/local/bin/node", "/opt/homebrew/bin/node", "/usr/bin/node"]
        guard let validNodePath = paths.first(where: { FileManager.default.fileExists(atPath: $0) }) else {
            print("Error: Node.js not found")
            return
        }
        
        // Look for index.js in Bundle Resources or current directory
        let bundleScriptPath = Bundle.main.path(forResource: "index", ofType: "js")
        let cwdScriptPath = FileManager.default.currentDirectoryPath + "/" + NODE_SCRIPT_PATH
        let scriptPath = bundleScriptPath ?? cwdScriptPath
        
        guard FileManager.default.fileExists(atPath: scriptPath) else {
             print("Error: index.js not found at \(scriptPath)")
             return
        }
        
        let process = Process()
        process.executableURL = URL(fileURLWithPath: validNodePath)
        process.arguments = [scriptPath]
        
        let inputPipe = Pipe()
        process.standardInput = inputPipe
        
        let outputPipe = Pipe()
        process.standardOutput = outputPipe
        process.standardError = FileHandle.standardError
        
        outputPipe.fileHandleForReading.readabilityHandler = { [weak self] handle in
            let data = handle.availableData
            if data.isEmpty { return }
            if let str = String(data: data, encoding: .utf8) {
                // Split by newline in case multiple JSONs come at once
                let lines = str.components(separatedBy: .newlines)
                for line in lines {
                    if line.isEmpty { continue }
                    if let jsonData = line.data(using: .utf8),
                       let json = try? JSONSerialization.jsonObject(with: jsonData, options: []) as? [String: Any] {
                        
                        DispatchQueue.main.async {
                            if let type = json["type"] as? String {
                                if type == "discord_status" {
                                    self?.discordConnected = json["connected"] as? Bool ?? false
                                    if let error = json["error"] as? String {
                                        self?.lastNodeLog = "RPC Error: \(error)"
                                    }
                                } else if type == "debug_log" {
                                    self?.lastNodeLog = json["message"] as? String
                                }
                            }
                        }
                    }
                }
            }
        }
        
        do {
            try process.run()
            nodeProcess = process
            self.nodePipe = inputPipe
        } catch {
            print("Failed to start node process: \(error)")
        }
    }
    
    var nodePipe: Pipe?
    
    func sendToNode(json: [String: Any]) {
        guard let data = try? JSONSerialization.data(withJSONObject: json, options: []),
              let pipe = nodePipe else { return }
        
        var dataWithNewline = data
        dataWithNewline.append("\n".data(using: .utf8)!)
        
        pipe.fileHandleForWriting.write(dataWithNewline)
    }
    
    
    func startPolling() {
        checkTimer = Timer.scheduledTimer(withTimeInterval: CHECK_INTERVAL, repeats: true) { _ in
            self.poll()
        }
    }
    
    // Keep track of the last reported app to avoid flapping if multiple are open
    var lastReportedBundleId: String?
    
    func poll() {
        // Check Idle
        let anyInput = CGEventType(rawValue: UInt32.max)!
        let idleTime = CGEventSource.secondsSinceLastEventType(.combinedSessionState, eventType: anyInput)
        if idleTime > IDLE_THRESHOLD {
            sendToNode(json: ["type": "status", "idle": true])
            return
        }
        
        
        let runningApps = NSWorkspace.shared.runningApplications
        
        var activeApp: NSRunningApplication?
        var backgroundApp: NSRunningApplication?
        
        for (bundleId, _) in appConfig {
            if let app = runningApps.first(where: { $0.bundleIdentifier == bundleId }) {
                if app.isActive {
                    activeApp = app
                } else {
                    backgroundApp = app
                }
            }
        }
        
        var targetApp: NSRunningApplication?
        
        // Prioritize Active > Last Reported (if still running) > Any Background
        if let active = activeApp {
            targetApp = active
        } else if let last = lastReportedBundleId, let bg = runningApps.first(where: { $0.bundleIdentifier == last }) {
            targetApp = bg
        } else if let bg = backgroundApp {
            targetApp = bg
        }
        
        if let app = targetApp, let bundleId = app.bundleIdentifier, let appName = appConfig[bundleId] {
            // Use specific script based on Bundle ID
            let (docName, _) = getDocumentName(bundleId: bundleId, appName: appName)
            
            lastReportedBundleId = bundleId
            let payload: [String: Any] = [
                "type": "status",
                "app": appName,
                "doc": docName,
                "idle": false,
                "customStatus": customStatus ?? ""
            ]
            sendToNode(json: payload)
        } else {
            lastReportedBundleId = nil
            sendToNode(json: ["type": "status", "idle": false, "clear": true])
        }
    }
    
    func getDocumentName(bundleId: String, appName: String) -> (String, String?) {
        var scriptSource = ""
        
        switch bundleId {
        case "com.microsoft.Word":
            scriptSource = """
            tell application "Microsoft Word"
                if running then
                    try
                        return name of active document
                    on error errMsg
                        return "ERROR: " & errMsg
                    end try
                else
                    return ""
                end if
            end tell
            """
        case "com.microsoft.Excel":
            scriptSource = """
            tell application "Microsoft Excel"
                if running then
                    try
                        return name of active workbook
                    on error errMsg
                        return "ERROR: " & errMsg
                    end try
                else
                    return ""
                end if
            end tell
            """
        case "com.microsoft.Powerpoint":
            scriptSource = """
            try
                tell application "Microsoft PowerPoint"
                    return name of active presentation
                end tell
            on error
                try
                    tell application "Microsoft PowerPoint"
                        return name of active window
                    end tell
                on error
                    tell application "System Events"
                        tell process "Microsoft PowerPoint"
                            try
                                return name of window 1
                            on error errMsg
                                return "ERROR: " & errMsg
                            end try
                        end tell
                    end tell
                end try
            end try
            """
        default:
            // Fallback to System Events for unknown apps
            scriptSource = """
            tell application "System Events"
                if exists process "\(appName)" then
                    tell process "\(appName)"
                        try
                            if (count of windows) > 0 then
                                return name of window 1
                            else
                                return ""
                            end if
                        on error errMsg
                            return "ERROR: " & errMsg
                        end try
                    end tell
                else
                    return ""
                end if
            end tell
            """
        }
        
        var error: NSDictionary?
        if let script = NSAppleScript(source: scriptSource) {
            let result = script.executeAndReturnError(&error)
            if error == nil {
                let str = result.stringValue ?? ""
                if str.hasPrefix("ERROR: ") {
                    return ("", String(str.dropFirst(7)))
                }
                return (str, nil)
            } else {
                return ("", error?["NSAppleScriptErrorMessage"] as? String ?? "Unknown AppleScript Error")
            }
        }
        return ("", "Failed to create NSAppleScript")
    }
    
    // --- Launch Agent ---
    
    func getAgentPath() -> URL {
        let libraryDir = FileManager.default.urls(for: .libraryDirectory, in: .userDomainMask).first!
        return libraryDir.appendingPathComponent("LaunchAgents/cl.jafu.OfficeRichPresence.plist")
    }
    
    func isLaunchItemPresent() -> Bool {
        if #available(macOS 13.0, *) {
            return SMAppService.mainApp.status == .enabled
        } else {
            return FileManager.default.fileExists(atPath: getAgentPath().path)
        }
    }
    
    func setLaunchOnStartup(enabled: Bool) -> Bool {
        if #available(macOS 13.0, *) {
            do {
                if enabled {
                    try SMAppService.mainApp.register()
                } else {
                    try SMAppService.mainApp.unregister()
                }
                return true
            } catch {
                print("Failed to toggle SMAppService: \(error)")
                return false
            }
        } else {
            // Fallback for older macOS
            let path = getAgentPath()
            if enabled {
                let execPath = Bundle.main.executablePath ?? (FileManager.default.currentDirectoryPath + "/office-rich-presence")
                let dict: [String: Any] = [
                    "Label": "cl.jafu.OfficeRichPresence",
                    "ProgramArguments": [execPath],
                    "RunAtLoad": true,
                    "KeepAlive": false
                ]
                let plistData = try? PropertyListSerialization.data(fromPropertyList: dict, format: .xml, options: 0)
                do {
                    try plistData?.write(to: path)
                    return true
                } catch {
                    print("Failed to write launch agent: \(error)")
                    return false
                }
            } else {
                do {
                    try FileManager.default.removeItem(at: path)
                    return true
                } catch {
                    return false
                }
            }
        }
    }
}

let app = NSApplication.shared
let delegate = AppDelegate()
app.delegate = delegate
app.run()
