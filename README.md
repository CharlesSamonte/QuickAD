# QuickAD

**QuickAD** is a lightweight PowerShell GUI tool designed to make managing Active Directory users faster and easier—no need to navigate the AD console or memorize complex command-line syntax.

Whether you're resetting passwords, moving users between OUs, or checking user details, QuickAD streamlines the process.

![PowerShell](https://img.shields.io/badge/Built%20with-PowerShell-blue)
![Windows](https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011-lightgrey)
![AD](https://img.shields.io/badge/Dependency-Active%20Directory-orange)

---

## Features
QuickAD currently supports:

- View user details by entering a username or full name  
- Reset user passwords to a default or custom value  
- Move users to a different Organizational Unit (OU)  
- List users with similar names for easy identification  
- Modify user attributes such as profile fields  
- Delete users (permanently or move to a "Deleted" OU)

---

## ⚠️ Important Notes

**This is not a plug-and-play solution.**  
QuickAD is an experimental script created as a learning project. It will **require customization** to work in your environment. Specifically, you will need to adjust:

- OU paths  
- Default passwords  
- Attribute names (if your AD schema is customized)  
- Any hardcoded references that apply to your organization  

There is currently **no configuration file**—all changes must be made directly in the script (`QuickAD.ps1`).

---

## How It Works

QuickAD is an interactive PowerShell GUI that uses pre-defined functions to interact with Active Directory. You input a user or select an action, and the GUI handles the AD logic for you.

---

## Requirements

- Windows 10 or 11  
- PowerShell with RSAT / Active Directory module installed  
- AD credentials with permission to view and modify user accounts  

---

## Getting Started

### Option 1: Run the Script Directly

1. Download `QuickAD.ps1` from this repository.  
2. Right-click and select **Run with PowerShell**.  
3. If prompted, enter valid AD credentials.  

### Option 2: Use the Shortcut

1. Clone or download the repository.  
2. Copy the included **QuickAD shortcut** to your desktop.  
3. Double-click the shortcut to quickly launch the tool.  

---

## Configuration

To customize QuickAD for your organization:

- Open `QuickAD.ps1` in your text editor  
- Locate sections with OU paths, passwords, and labels  
- Modify values to suit your environment  

> There is currently no external config file, so customization is done directly in the script.

---

## Usage Examples

- Look up a user’s department or job title in seconds  
- Reset a user's password to your organization's default  
- Move users into different OUs without using the AD console  

---

## Screenshots

### Main Interface
![Main Interface](images/main-interface.png)

### User Details View
![User Details](images/user-details.png)

### Reset Password
![Reset Password](images/reset-password.png)

### Move User to OU
![Move User to OU](images/move-ou.png)

---

## Security & Error Handling

- Prompts for AD credentials (if not already authenticated)  
- Displays real-time error messages in the GUI  
- No persistent logging is implemented at this time  

---

## License & Attribution

This project is free to use and modify for personal or internal use.  
If you find it useful, credit is appreciated but not required.

---

## File Info

- `QuickAD.ps1` – Full, self-contained PowerShell GUI script

---

## Contact

Have suggestions or issues?  
Feel free to open an issue or contribute via pull request.

---

## Final Note

QuickAD is a tool built for learning and convenience—**not a polished enterprise-grade solution**. You are encouraged to inspect the code, test it thoroughly, and tailor it to your specific environment before using it in any production context.
