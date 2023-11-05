# AppScript Email Automation

Code to send bulk emails from a google sheet with randomised delay in between each email to prevent being detected as spam.

## Environment Setup

1. Make a <a href="https://sheet.new">new sheet</a>
2. Rename the first subsheet to `Email Configuration`
3. Add another sheet and rename it to `Errors`
4. In the `Email Configuration` subsheet add the following headers
<table>
    <tr>
        <th style="border:1px white solid">Name</th>
        <th style="border:1px white solid">Email</th>
        <th style="border:1px white solid">cc</th>
        <th style="border:1px white solid">bcc</th>
        <th style="border:1px white solid">Subject</th>
        <th style="border:1px white solid">Content</th>
        <th style="border:1px white solid">HTML Content</th>
        <th style="border:1px white solid">Attachments</th>
    </tr>
</table>
5. Go to **Extensions** menu and select **Appscript**
6. In the **Code.gs** file paste the code from <a href="./Code.gs">Code.gs</a>
7. Make a new file **Constants.gs** and paste the code from <a href="./Constants.gs">Constants.gs</a>
8. Save the project
9. Go to the google sheet and refresh the page
10. Now you will see a menu called **Email** with an option **Send Email**

## Configurable Elements

1. You can configure all the headers in the sheet by changing the code in the **Constants.gs** file by changing the respective header value enclosed in Double Quotes (""),
2. There is a random delay added between each email so that Google doesn't consider all your emails as sent from bots. The max delay can be configured from **Constants.gs** file

## Glossary and Conventions to be followed (\* required)

- Meaning of each header is explained below
  1. Name: Name of the receiver
  2. Email\*: Email of the receiver (should be comma separated if multiple)
  3. cc: CC's to be included (should be comma separated if multiple)
  4. bcc: BCC's to be included (should be comma separated if multiple)
  5. Subject: Subject of the email
  6. Content\*: Content of the email (for devices that do not support HTML content)
  7. HTML Content\*: HTML content of the email (for devices that support HTML Content)
  8. Attachments: IDs of files uploaded on drive (comma separated if multiple)

## Error Handling

- Whenever an error occurs while sending emails, the code logs them to the **Errors** subsheet.
- All the columns of the **Email Configuration** sheet along with one more Error column are logged in the Errors sheet.
