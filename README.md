# OfficeRichPresence for macOS

A lightweight, native macOS menu bar application that displays your Microsoft Office (Word, Excel, PowerPoint) status on Discord.

## Features

- **Native macOS App**: Built with **Swift**.
- **Multi-App Support**: Automatically switches between Word, Excel, and PowerPoint.
- **Custom Status**: Set a custom message (e.g., "Writing a Thesis") that overrides the default text.
- **Launch at Login**
- **Privacy**: Option to hide filenames.
- **Idle Detection**: Automatically clears status when you are away.

> [!IMPORTANT]
> **"Unidentified Developer" Warning**
>
> Because this is a free, open-source project, I do not have a paid Apple Developer Certificate ($99/year).
> When you first open the app, macOS will warn you that the developer cannot be verified.
>
> **To Open the App:**
> 1.  **Right-click** (or Control-click) the app icon.
> 2.  Select **Open**.
> 3.  Click **Open** in the dialog box.
>
> You only need to do this once.

## Build It Yourself (Xcode)

For the best experience and to verify the code yourself, you can build the app using Xcode.

### Prerequisites
- macOS 13.0 or later (Ventura/Sonoma/Sequoia)
- Xcode (latest version recommended)
- Node.js (installed via Homebrew or installer)

### Steps

1.  **Create Project**:
    - Open Xcode.
    - Select **Create a new Xcode project**.
    - Choose **macOS** -> **App**.
    - Product Name: `OfficeRichPresence`.
    - Interface: **Storyboard**.
    - Language: **Swift**.
    - Uncheck "Create Git repository on my Mac" (if you cloned this repo).

2.  **Clean Up Template**:
    - Delete `ViewController.swift`.
    - Delete `SceneDelegate.swift` (if present).
    - Open `Info.plist` (or the Info tab in Target settings) and remove the "Main storyboard file base name" entry (we don't use a storyboard).
    - *Optional*: You can delete `Main.storyboard` as well.

3.  **Add Code**:
    - Open `main.swift` (if it doesn't exist, create it).
    - Paste the contents of `Sources/main.swift` from this repository into your project's `main.swift`.
    - **Important**: Ensure there is no `@main` attribute in `AppDelegate.swift` if you kept it. The `main.swift` file handles the entry point.

4.  **Add Resources**:
    - Drag and drop `index.js` into your project navigator. Ensure "Copy items if needed" is checked and it is added to the "OfficeRichPresence" target.
    - Drag and drop `menubar_icon.png` into `Assets.xcassets` (or just the project navigator). Rename it to `menubar_icon` if needed.
    - Ensure `package.json` is **NOT** added to the target resources (it's not needed at runtime, only for `npm install`).

5.  **Install Dependencies**:
    - Open Terminal.
    - Navigate to your project folder (where `index.js` is).
    - Run `npm install` to install `discord-rpc`.
    - This creates a `node_modules` folder. You must ensure this folder is accessible to the app, or bundle it.
    - *Simpler Method*: The app looks for `index.js` in the bundle. It expects `node_modules` to be next to it or globally available. For a self-contained build, drag the `node_modules` folder into your Xcode project as a **Folder Reference** (blue folder icon), so it gets copied into the app bundle.

6.  **Entitlements (Sandbox)**:
    - Go to your Target settings -> **Signing & Capabilities**.
    - **Remove "App Sandbox"**. This app requires AppleScript to talk to Office apps, which is restricted in the Sandbox.
    - Add **"Hardened Runtime"** if you plan to notarize (optional for local use).
    - In `Info.plist`, add the key `Privacy - AppleEvents Sending Usage Description` with a value like "Needed to check Office status."

7.  **Build & Run**:
    - Press **Cmd+R** to build and run!

## Build It Yourself (Command Line)

If you prefer the command line, a script is provided:

```bash
./create_app.sh
```

This script compiles the Swift code, bundles the resources, and creates a ready-to-use `OfficeRichPresence.app` in the `build` directory.

## License

MIT License. Copyright (c) 2025 Jafu.
