# Smart Mail Automation System
A tool that automatically sends personalized emails with attachment to every event participant, removing the need for manual mailing.

## Description
This project automates the distribution of personalized emails with attachments—such as participation certificates—to event attendees.

It reads recipient details from an Excel/CSV file, generates customized email content, attaches the correct certificate for each person, and sends everything securely through SMTP.

By eliminating manual sending, it saves time, reduces errors, and ensures every participant receives their documents quickly and professionally.

## Features

#### 1. Automated Email Sending

Sends emails to multiple recipients without manual intervention.

#### 2. Personalized Messages

Customizes the email content for each participant using their details.

#### 3. Attachment Support

Sends files like certificates or documents to every recipient individually.

#### 4. Excel/CSV Integration

Reads participant data (name, email, etc.) directly from data files.

#### 5. Secure SMTP Communication

Uses SSL for safe and reliable email delivery.

#### 6. Error Reduction 

Minimizes mistakes caused by manual mailing tasks.

#### 7. Time Efficient

Delivers certificates to all participants in minutes instead of hours.

#### 8. Scalable

Easily handles small groups or large batches of participants.

#### 9. Customizable

Email subject, body, and attachment names can be modified as needed.

## Getting Started
This is an example of how you may give instructions on setting up your project locally. To get a local copy up and running follow these simple example steps.
### Dependencies
* Windows
### Prerequisites
#### 1. Ensure python is installed
```bash
python --version
```
#### 2. Install pandas
```bash
pip install pandas
```

### Installing
* Upload your student data (CSV file) under `event_data` folder

* (OPTIONAL) Modify body of the mail in `mail_template.txt`

* Upload `myEventt_mail_data.csv` under `event_data` folder

* Upload attachment files under `myEvent_certificates` folder

* In `send_cert_with_SMTP` , `resend_cert_with_SMTP` file, update the following:
  
  + `SENDER_EMAIL = "INSERT EMAIL"`
    
  + `APP_PASSWORD = "INSERT PASSWD"`
    
  + `CSV_PATH = "mail_req_csv/myEventt_mail_data.csv"`
    
  + `CERT_FOLDER = Path(BASE_DIR) / "myEvent_certificates"`

  + `SUBJECT = "Insert Subject of mail"`

  + `EVENT_NAME = "My event name"`
    
* In `send_cert_without_SMTP` file, make sure you have Outlook Desktop installed on your system and login with the mail credentials you want to send mails from


### Executing program
#### Run
```
python send_cert_with_SMTP.py
```
## Contributing
Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are greatly appreciated.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement". Don't forget to give the project a star! Thanks again!

* Fork the Project
* Create your Feature Branch `git checkout -b feature/AmazingFeature`
* Commit your Changes `git commit -m 'Add some AmazingFeature`
* Push to the Branch `git push origin feature/AmazingFeature`
* Open a Pull Request
## Author
[Vibitha Varshini S S](https://github.com/ssvibitha)
