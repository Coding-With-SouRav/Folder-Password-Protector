# Demo Images

<img width="987" height="905" alt="Screenshot 2026-04-22 124153" src="https://github.com/user-attachments/assets/a7dfbd2f-6c2d-4c68-801b-e9cc633cae71" />
<img width="999" height="914" alt="Screenshot 2026-04-22 124122" src="https://github.com/user-attachments/assets/d77a54a0-f4b0-4a70-96e4-7ca37118b3f7" />
<img width="997" height="901" alt="Screenshot 2026-04-22 124313" src="https://github.com/user-attachments/assets/e7d8650b-3bdd-4f2e-a61e-e6e7116588e8" />

## Application Download Link
[Download Folder Locker]([https://github.com/Coding-With-SouRav/Folder-Password-Protector/releases/download/v1.0.0/Folder.Locker.exe](https://github.com/Coding-With-SouRav/Folder-Password-Protector/releases/download/v1.0.0/Folder.Locker.exe))

# How to Use the Folder Locker Application

This script creates a Windows application that password‑protects folders. When anyone tries to open a protected folder in File Explorer, a password dialog appears. The protection works even after the GUI is closed, thanks to a background monitor process.

## Using the Application

#### Default Master Password
The initial password is **`1234`**.

#### Protect a Folder
1. Go to the **Protected Folders** tab.
2. Click **Add Folder** and select any folder on your computer.
3. The folder appears in the list. It is now protected.

#### Open a Protected Folder
- When you double‑click the folder in Windows File Explorer, a password dialog pops up.
- Enter the **master password** (default `1234`) and click **OK**.
- The folder opens normally. For the next 1.5 seconds, that folder is temporarily allowed (grace period) so you can browse without re‑entering the password.

> If you enter the wrong password, the Explorer window that tried to open the folder is closed immediately.

#### Remove Protection
1. In the **Protected Folders** tab, select one or more folders from the list.
2. Click **Remove Selected**.
3. You will be asked for the master password – this prevents unauthorised removal.
4. Confirm the removal. The folders are no longer protected.

#### Change the Master Password
1. Go to the **Settings** tab.
2. Enter your **current password**, then the **new password** (minimum 4 characters) and confirm it.
3. Click **Change Password**.
4. All protected folders will now use the new password.

#### Background Monitor
- The monitor process runs independently (started as a separate `multiprocessing` process).
- It watches all File Explorer windows and intercepts attempts to open any folder listed as protected.
- **Even if you close the main GUI, the protection remains active** because the monitor continues to run in the background.
- To stop the monitor completely, you must end the process manually (e.g., via Task Manager). Look for a Python process running `main.py` or the packaged executable.
