# Hearthstone Battlegrounds Skip

A lightweight Windows console utility that lets you safely toggle Hearthstone's network connectivity to skip combat animations in Battlegrounds. The tool manages a Windows Firewall rule on your behalf, providing a simple guided workflow with colorful status feedback.

## Features

- One-key "Skip now" command that temporarily blocks and restores network access for `Hearthstone.exe`.
- Automatic firewall rule creation and status updates in real time.
- Guided prompts to locate your Hearthstone installation and configure disconnect duration.
- Friendly console interface with color-coded indicators for rule state and permissions.

## Getting Started

1. **Download**
   - Grab the published executable here: [Download HsSkipCombat.exe](https://github.com/hafnerpw/hs-skip-combat/releases/latest/download/HsSkipCombat.exe)

2. **Run the app**
   - Double-click the downloaded `HsSkipCombat.exe`, or launch it from a terminal:
     ```powershell
     .\HsSkipCombat.exe
     ```

3. **Follow the prompts**
   - The first run will help you locate `Hearthstone.exe` and set your preferred disconnect duration.
   - Administrative privileges are required the first time to create the firewall rule; the app will offer to relaunch elevated if needed.

## 🛠 Development

This project targets **.NET 9** on Windows.

### Build from source

```powershell
# Restore dependencies and build
dotnet build HsSkipCombat/HsSkipCombat.csproj

# Run locally
dotnet run --project HsSkipCombat/HsSkipCombat.csproj
```