# MailMergeAndScheduler
The Google App Script, creates a mail merge. However google limits the number of mail that could be sent, hence the script will schedule the rest of mail after 24 hours.

Problem Statement - For Toastmasters club, before every meeting a mail is sent to people who are not members but attend meeting. The mail should have meeting number, date, venue and agenda (attachment). If the mail is personalized, it makes more imapct. Morever, if a mail is sent to more than 100-200 people in BCC gmail marks it as spam. Moreover, we couldnt use the existing add on as they had set the limit as 50, hence it will require multiple scheduling. So I came up with a highly specific mail merge and schedule script.

Structure of File - 

The googlesheet should have two sheet - 

Sheet # 1 - Name "Mail Merge" - The sheet will contain the following columns - 

First Name,
Last Name,
Email Address,
Meeting Number,
File Attachments,
Schedule Date,
Mail Merge Status

If you want to delete any column, please make the change in App Script too.

Sheet # 2 - Name "Meeting Info". The sheet will contain the link to agenda attachment (it has to be drive).
