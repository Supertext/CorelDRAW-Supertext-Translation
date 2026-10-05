# Installation guide

The round trip is one VBA module, `SupertextTranslation.bas`, containing two macros: **TranslationExport** and **TranslationImport**. Installing means importing the module into a CorelDRAW macro project once.

Menu names below follow CorelDRAW 2021 and later. Older versions have the same functions under **Tools › Macros** instead of **Tools › Scripts**.

- [1. Check that VBA is installed](#1-check-that-vba-is-installed)
- [2. Download](#2-download)
- [3. Import the module](#3-import-the-module)
- [4. Verify](#4-verify)
- [Optional: toolbar button and shortcuts](#optional-toolbar-button-and-shortcuts)
- [Rolling out to many machines](#rolling-out-to-many-machines)
- [Updating](#updating)
- [Uninstalling](#uninstalling)
- [Troubleshooting](#troubleshooting)

## 1. Check that VBA is installed

VBA is an optional component of CorelDRAW Graphics Suite. Press **Alt+F11** in CorelDRAW. If the Visual Basic editor opens, VBA is installed.

If nothing happens or the Scripts menu entries are greyed out: open **Windows Settings › Apps**, select CorelDRAW Graphics Suite, choose **Modify**, and enable **Visual Basic for Applications** in the feature list.

## 2. Download

Download the latest release from the [Releases page](https://github.com/Supertext/CorelDRAW-Supertext-Translation/releases), or clone the repository. You need `src/SupertextTranslation.bas`.

If you download the file through the browser's **Raw** view, make sure it is saved as `SupertextTranslation.bas` and not as `.bas.txt`. The file must keep its Windows line endings; the release download and a normal `git clone` on Windows both do.

## 3. Import the module

1. In CorelDRAW, press **Alt+F11** (or **Tools › Scripts › Script Editor**) to open the Visual Basic editor.
2. In the **Project** pane on the left, select **GlobalMacros (GlobalMacros.gms)**. This project is loaded every time CorelDRAW starts. If you prefer a separate project, see [Rolling out to many machines](#rolling-out-to-many-machines).
3. Choose **File › Import File…** and select `SupertextTranslation.bas`. A module named **SupertextTranslation** appears under **Modules**.
4. Choose **Debug › Compile GlobalMacros**. No message means it compiled.
5. Press **Ctrl+S** to save the project, and close the editor.

## 4. Verify

Open **Tools › Scripts › Scripts** (the Scripts docker). Under **GlobalMacros › SupertextTranslation** you should see:

- **TranslationExport**
- **TranslationImport**

Double-click one to run it. Alternatively, press **Alt+F8**, pick the macro from the list and click **Run**.

## Optional: toolbar button and shortcuts

1. Choose **Tools › Customization…** (or **Tools › Options › Customization**).
2. Open **Commands** and pick **Macros** in the category list.
3. Drag **GlobalMacros.SupertextTranslation.TranslationExport** onto any toolbar. Do the same for the import.
4. On the **Shortcut Keys** tab of the same page you can assign a keyboard shortcut to each macro.

## Rolling out to many machines

The simplest route is a dedicated macro project file, `SupertextTranslation.gms`, built once and copied to each machine:

1. On one machine, create a new macro project. In the Scripts docker menu, choose **New › New Macro Project** (the wording varies slightly between versions) and name it `SupertextTranslation`.
2. Import `SupertextTranslation.bas` into that project as in step 3 above, compile and save.
3. The project is saved in your user GMS folder, for example:

   ```
   %APPDATA%\Corel\CorelDRAW Graphics Suite 2024\Draw\GMS\SupertextTranslation.gms
   ```

   The version folder name differs between releases. In the VBA editor, the project's tooltip and the **Properties** pane show the exact path.

4. Copy the `.gms` file into the same folder on other machines, for example with PowerShell or Intune:

   ```powershell
   $suite = "CorelDRAW Graphics Suite 2024"
   $dest = "$env:APPDATA\Corel\$suite\Draw\GMS"
   New-Item -ItemType Directory -Force -Path $dest | Out-Null
   Copy-Item ".\SupertextTranslation.gms" -Destination $dest -Force
   ```

CorelDRAW loads every `.gms` in that folder at startup. Keep the `.bas` file in Git as the source of truth; the `.gms` is a build artefact (see the [developer guide](developer-guide.md#development-setup)).

## Updating

1. Open the VBA editor (**Alt+F11**).
2. Right-click the **SupertextTranslation** module and choose **Remove SupertextTranslation…**. Answer **No** when asked to export it.
3. Import the new `SupertextTranslation.bas`, compile and save as in step 3.

Check [CHANGELOG.md](../CHANGELOG.md) first. Text IDs already stored in documents stay valid across updates.

## Uninstalling

Remove the module as in [Updating](#updating), step 2, and save. If you used a separate `SupertextTranslation.gms`, close CorelDRAW and delete that file instead.

The hidden IDs remain in documents that were exported. They are invisible, don't affect printing or export, and can be ignored. See the [developer guide](developer-guide.md#text-ids) if you need to remove them.

## Troubleshooting

| Problem | Fix |
|---|---|
| Alt+F11 does nothing, Scripts menu greyed out | VBA isn't installed. See [step 1](#1-check-that-vba-is-installed). |
| "Macros are disabled" or a security warning | Allow macros for the project. Macro security settings are in CorelDRAW's options under VBA. GlobalMacros is normally trusted. |
| Import shows garbled code or a single long line | The file lost its Windows line endings (CRLF). Download the release file again instead of copying from the browser. |
| "Compile error: User-defined type not defined" | You imported into a project that isn't a CorelDRAW macro project. Use GlobalMacros or a `.gms` project. |
| Macro doesn't appear in the Scripts docker | Make sure you saved the project (Ctrl+S in the VBA editor) and that the module is called SupertextTranslation. |
| Error with a line number during export or import | Note the message, the CorelDRAW version and what you were doing, and report it in the repository's issues. |
