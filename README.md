# All-Saved-WiFi-Passwords-Grabber-VBS

**Effortlessly Retrieve and Save All Wi-Fi Passwords on Windows!**

This Visual Basic Script (VBS) allows you to quickly fetch and save all Wi-Fi passwords stored on your Windows machine to a text file. The script automates command-line operations, reads Wi-Fi profile details, and stores them in a structured output file. No manual intervention needed!

## Features:
- **Fetch & Save:** Retrieves Wi-Fi passwords and saves them to `wifi_passwords.txt` for future reference.
- **Automated Workflow:** Runs `netsh` commands in the background to extract network profiles and their respective passwords.
- **Organized Output:** Saves detailed Wi-Fi profiles in a readable format.
- **File Management:** Creates and manages temporary files to handle command outputs efficiently.
- **Safe Cleanup:** Automatically deletes temporary files after use to keep your system tidy.

## How It Works:
1. The script creates a `wifi_passwords.txt` file in the same directory as the script, where all Wi-Fi details will be stored.
2. It extracts all saved Wi-Fi profiles and retrieves their corresponding passwords.
3. The information is saved in an organized manner with the Wi-Fi name and details listed.
4. After completion, the script cleans up any temporary files used during execution and displays a "Done" message.

## How to Use:
1. Download the `grab_wifi_passwords.vbs` file and place it in any folder.
2. Double-click the script to run it.
3. A file named `wifi_passwords.txt` will be created in the same directory as the script. Open this file to view the saved Wi-Fi profiles and passwords.
4. A message box will appear confirming the operation is done.

## Requirements:
- Windows OS
- Administrator privileges (needed to retrieve network profiles and passwords)

## Output File:
The retrieved Wi-Fi profiles and passwords are saved in `wifi_passwords.txt`. The file will include:
- Wi-Fi profile names
- Passwords (if available)
- Other network-related details (like authentication types)

Example of output in `wifi_passwords.txt`:

```text
Profile: YourWiFiNetwork
    SSID                   : YourWiFiNetwork
    Key Content            : yourpassword123
    Authentication         : WPA2-Personal
--------------------------------------------------
Profile: OfficeNetwork
    SSID                   : OfficeNetwork
    Key Content            : officesecret123
    Authentication         : WPA2-Personal
--------------------------------------------------
