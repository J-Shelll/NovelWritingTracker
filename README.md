# Novel Writing Tracker

My first ever coding project. This is my original idea to help someone practice the writing progress on their novel or other book.
Previously, I used a tracker like this in Excel and I implemented Excel formulas to automate the process. I now created a project where just running
the script it calculates and updates total wordcount today, today's wordcount and if the 500 word target was met. The project works well from my testing.

I used ChatGPT 5.5 Thinking on this project.


## What it does

This script:

- Opens a Microsoft Word document
- Counts the current number of words
- Opens an Excel tracker
- Finds today's date in the tracker
- Writes the current total word count
- Calculates words written today by subtracting yesterday's total
- Checks whether the daily 500-word target was met
- Writes `Met` or `Short` into the tracker

## Technologies used

- PowerShell
- Microsoft Word COM automation
- Microsoft Excel COM automation

## Requirements

- Windows
- Microsoft Word
- Microsoft Excel
- PowerShell

## File setup

Place these files in the same folder:

```text
NovelScript.ps1
Novel.docx
NovelTracker.xlsx
```
