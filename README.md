# Julian Timehop
This is a simple Python script that produces an email from a list of diary entries about my son's activities/achievements. It downloads a Word document from Dropbox, parses a table in the document, and sends an appropriate email. It is run every day on my Linux server (using cron).

See [this blog post](http://blog.rtwilson.com/creating-an-email-service-for-my-sons-childhood-memories-with-python) for more details.

## Requirements:
 - python-docx
 - pandas
 - emails
 - [dropbox](https://pypi.org/project/dropbox/)

You will also need to set the following environment variables:

 - `SMTP_PASSWORD`
 - `DROPBOX_KEY`