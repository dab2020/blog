---
layout: post
title: "How I Managed To Automate Creating Event ID Cards For Over 1500 People"
date: 2024-04-05 10:00:00 -0500
categories: tech
tags: tech programming
---
Anyone who has organised events knows that one of the most time consuming tasks for their Media and IT department is creating ID Cards. The task appears to be quite simple on the surface but can get technical really quickly. Here's how I saved our team tens if not hundreds of man hours and uncountable frustration by automating the process.

### Background

Like I said, the task appears to be simple. A good baseline is using Photoshop's Data Set Feature, which allows us to define variables in a template and change them, for each dataset. It also allows us to import datasets either through .csv files or tab delimited .txt files. It works perfectly fine without any major hiccups until you get to images where all the difficulty really begins to reveal itself. The thing is the data set requires all files to be stored locally and with their file path given. For most events who are using Google forms or similar software this slowly becomes a nightmare . like in our case we had all the files on google drive with randomised file names.

---



### First Attempt

Now I first had a go at solving this in January of 2023, when I was working as the Co-Director IT for an event, I conjured up this really janky solution written in Visual Basic .NET. I isolated the fileID for the pictures and inserted them into this URL `"https://drive.google.com/uc?export=download&id="` to generate a direct download link. I exported the google sheet as an excel file and used `Microsoft.Office.Interop` to interact with it. I looped through the rows and saved the files with the delegate name in the photoshop data set folder. I then used the photoshop data set ability as normal.

I no longer have the exact code but this a early draft.

```vbnet
 Dim eDC As New Application
        Dim workbook As Workbook = excelApp.Workbooks.Open("file.xlsx")
        Dim worksheet As Worksheet = workbook.Sheets(1)
        Dim rowCount As Integer = worksheet.UsedRange.Rows.Count

        ' Loop through all rows in the second column where the link is stored
        For i As Integer = 2 To rowCount
            ' Get the link and the name
            Dim link As String = worksheet.Cells(i, 2).Value
            Dim name As String = worksheet.Cells(i, 1).Value

            ' Download the file
            Dim client As New System.Net.WebClient
            client.DownloadFile(link, name & ".jpg")

            ' Replace the link with the file path
            worksheet.Cells(i, 2).Value = "2"
        Next

        ' Save the changes
        workbook.Save()
        workbook.Close()
        excelApp.Quit()
```

---



### Second (Better) Attempt

I had while working on a different project familiarised myself with the Google app script platform, if you don't know about it or have been sleeping on it, like I had, DON'T. The way apps script allows you to interact with Google Services is unmatched. Appscript is written in javascript, which I has some basic understanding of from my web dev days. Even if you don't, it's a pretty easy language to pick up within a day or two.
Our Registration data was stored in a Google Sheet that was auto-populated from the Google Forms Responses. I duplicated this sheet and formatted it to include only the participants name, picture and team number. Now like I said, automating the text replacement is a pretty straight forward process. The problem with images was the randomised file names and paths. I wrote a simple app script that allowed us to get the names of the files from the fileID's of the pictures. The entire folder containing the pictures could then be downloaded and cross referenced using that sheet and imported as a dataset into photoshop.

We are going to use the `SpreadsheetApp`  service baked into Appscript. It treats google sheets as a grid saved as a 2 dimensional array. To get data from the sheet we must first get access to where the data is stored, then get the range of the spreadsheet 

```js
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var dataRange = sheet.getDataRange();
var data = dataRange.getValues();
```

The values from the sheet are now stored in the `data` variable. The way my sheet was formatted was such that the third column contained the fileIDs of the pictures. So we are now going to iterate through that and get the fileID

```js
for (var i = 1; i < data.length; i++) {
    var fileId = data[i][2];
    // Rest of The Code Described Below
}
```

Within the same loop, we are gonna make sure that the fileID variable is not blank, and then using the `Drive.App.getFileById` Function get the name of the file. I made it into a function called `getFilename()`

```javascript
    if (fileId) {
        var name = getFilename(fileId)
        // More Code
    }
```

With the Function `getFilename()` :

```js
function getFilenam(fileId) {
  var file = DriveApp.getFileById(fileId);
  var fileNam = file.getName();
  return fileNam;
}
```

We can now set the value of the name into the google sheet. I inserted it into column 4.

```javascript
   sheet.getRange(i + 1, 4).setValue(name);
```

---



For a a good tutorial on how to use the photoshop dataset feature, a good starting point is this video:

<div>
<iframe width="944" height="531" src="https://www.youtube.com/embed/Acjn7HoqrJ8" title="Auto-Create 100s of Custom Designs using "Variables" in Photoshop!" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" referrerpolicy="strict-origin-when-cross-origin" allowfullscreen></iframe>
</div>
