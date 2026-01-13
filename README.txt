========================================
EmailDeleter - Setup Instructions
========================================

FIRST-TIME SETUP (Do this ONCE per computer):
---------------------------------------------

1. Locate the file "setup-env.bat" (it should be in the same folder as this README)

2. Double-click "setup-env.bat" to run it

3. Wait for the message "Setup Complete!" to appear

4. Close the window

5. You're done! You will NOT need to run this setup again unless you:
   - Reinstall Windows
   - Change computers
   - Are asked to update credentials by IT support


DAILY USE:
----------

Simply run "EmailDeleter.exe" from:
- A desktop shortcut
- A network location (e.g., \\SERVER\Apps\EmailDeleter\EmailDeleter.exe)
- Any local folder

No setup or configuration is needed after the first-time setup.


IMPORTANT NOTES FOR NETWORK DEPLOYMENT:
----------------------------------------

If your IT department has placed EmailDeleter on a network share:
- You still need to run setup-env.bat ONCE on YOUR computer
- The BAT file configures YOUR computer to work with the app
- After setup, you can run EmailDeleter.exe from the network location


TROUBLESHOOTING:
----------------

Problem: EmailDeleter shows an error about "Graph API credentials not found"
Solution: Run setup-env.bat again on your computer

Problem: The app worked before but stopped working after reinstalling Windows
Solution: Run setup-env.bat again (Windows reinstall clears the configuration)

Problem: I changed computers and the app doesn't work
Solution: Run setup-env.bat on the new computer

For other issues, please contact your IT support.


TECHNICAL DETAILS:
------------------

The setup-env.bat file configures environment variables on your computer
that EmailDeleter needs to connect to Microsoft Graph API. These variables
are stored securely in the Windows registry and persist across restarts.

The setup only needs to be performed once per computer, and the configuration
will remain until Windows is reinstalled or the variables are manually removed.
