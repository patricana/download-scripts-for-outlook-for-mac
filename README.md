# What is it?
This repository contains an AppleScript file that downloads the emails in inbox or sent items (depending on your selection) to your disk.

# Why I need this?
You may need for the following use cases
- if need to back-up your full inbox to disk before you make changes to your outlook account
- if you have an auto-delete policy in your outlook that you don't want to (or can't) change but would like to keep important emails from being auto-deleted
- if you need do fast text-based search in your emails (by having a text copy on desk)

# How do I use it?
1. Download the script on your mac
2. Open Script Editor and copy the script there. Save as script file if you want to make some changes of your own. Otherwise, export as "Application"
3. Open outlook and download all emails (e.g., somemtimes only headers or partial emails are downloaded based on settings, make sure to download full emails first. This may take some time depending on the items you may have)
4. This script saves items as text, or EML (that includes attachments and formatting etc.), or both (depending on your selection). However, it requires two "categories" to be created in outlook with "txt" and "eml" names to keep track of which emails have been downloaded. You first need to create these manually in your outlook.
5. Run the Application file you had saved
6. Check if any errors are reported (text extraction almost always succeeds but eml downloads throws some erros on longer runs, like when trying to download the full inbox. These errors are caused by Microsoft's poor APIs for AppleScript and restarting outlook and running script again typically solves them).
