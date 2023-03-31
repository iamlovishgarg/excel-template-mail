- [DESCRIPTION](#description)
- [CONFIGURATION](#configuration)
- [Generating App Password](#generating-app-password)
- [DEVELOPER INSTRUCTIONS](#developer-instructions)
    - [Data Excel File](#data-excel-file)
    - [Template Text File](#template-text-file)
    - [Credentials File](#credentials-file)
    - [Filenames File](#filenames-file)
    - [Main File](#main-file)
    - [Summary](#summary)
- [FAQ](#faq)


# DESCRIPTION
Excel-Template-Mail is a 3 step easy process to send personalized emails in bulk directly. It is like https://merge.email/ but just free. Although it doesn't have the same amount of features like attaching files to email, but they can be added in future. 

# CONFIGURATION

Just open command prompt or terminal in current working directory and run this command:
```bash
pip install -r requirements.txt
```

# Generating App Password

To login into your google account(with 2 step verification "on", as old way of using third party apps is depricated and also is not secure) through python, you need to get a 16 character password. <br /><br />
You can go on this link to generate app password:<br />
https://myaccount.google.com/u/4/apppasswords

OR

You can also check this video out if you have any query:<br />
https://www.youtube.com/watch?v=g_j6ILT-X0k

# DEVELOPER INSTRUCTIONS

There is nothing much to do after making app password. You just need to provide the data. A sample dataset is provided in the repository. 

## Data Excel File
Here you can write the data, the variables which will be filled in template, like:
names | ai_marks | ml_marks | python_marks | emails
| :--- | ---: | :---: | :---: | :---:
Ashish  | 35 | 36 | 60 | ashish@gmail.com
Shubham  | 35 | 37 | 70 | shubham@gmail.com
Lovish  | 33 | 35 | 90 | lovish@gmail.com

## Template Text File
Here you can write the message you want to send. You can write variables in this file through <code>{}</code> syntax, like this:
```html
Hello {names},

It is to inform you that, you have gotten marks as follows:
AI: {ai_marks}
ML: {ml_marks}
Python: {python_marks}

Regards,
Teacher
```

## Credentials File
Email and the generated app password is saved here. 

## Filenames File
Excel file and template file name is saved here(in case you want to change it, you don't have to open main file). 

## Main File
This is the main script which does all the work. You don't need to edit anything in this. 

## Summary
So, the flow of the app is like, fill in the excel file with unique values(values which are supposed to change), put template with variables in <code>{}</code> syntax in "template.txt" file. <u>Column names in excel file would be variable names</u> in "template.txt". You can put <b>as many columns and rows as you want</b>. "emails" column would be <b>ignored</b> while making template as it is used to send emails and is not used in the template. 

# FAQ

1. Can we make as many columns and rows as we want? <br/>
Answer: <b>Yes.</b>

2. Can we attach files? <br />
Answer: <b>As of now, no.</b>

3. Which language is it made in? <br />
Answer: <b>Python.</b>

4. Does it uses new way of authentication in gmail as old way(third party apps) is depricated? 
Answer: <b>Yes, the process of doing it is written above in "Generating App Password" Column.</b>

5. I am getting error, how to resolve it? <br />
Answer: <b>For every error, i have written an assertion. So, just check the last line, you will know what the reason is. </b>