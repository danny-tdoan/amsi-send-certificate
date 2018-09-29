# Overview
A simple script to send certificates to participants following a predefined template. To be generic, the functionality of this script is to send emails to different recipients, with different attachments, using an email template. This expands the functionality of MailMerge, and allows different attachments for different recipients.

The main functionality:
1. Prepare the attachments for recipients following a mapping <name/id>:<attachment>
2. Prepare the email as a draft and save in Draft box. The drafts are to be reviewed before sending out
3. Send email without draft review.

# Usage
The program is initially written to be used on Windows, although it can be easily expanded to other environment, by preparing other `sh` scripts.

Use Step 1 and Step 2 bat files. The scripts are self-explanatory.

Please refer to the instruction under `instructions` for more details.