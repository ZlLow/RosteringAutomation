# Roster Automation
This project is an automated roster system. It automates the creation of a spreadsheet template that will retrieve an individual availability. 

From these availability, it will automatically create a roster list and generate a timesheet accordingly.
## About
This project is purely a scripting project utilizing Google Sheet and [Google App Script](#acknowledgement).

<sub>It is  **created entirely** on [Google App Script](#acknowledgement).
It would not be able to run on any platforms except **Google App Script**.</sub>




## Structure of the Project
Folders are arranged based on the functionalities of the script.
Within these folders contains individual script.

| Folder Name | Description                                                                        |
|-------------|------------------------------------------------------------------------------------|
| Error       | Handles all errors and log all errors                                              |
| JSON        | Contains the JSON file of the Project. <br/>Can be imported into Google App SCript |
| Menu        | Creation of All Menus within Spreadsheets                                          |
| Misc        | Iterative Folder Search and File Search                                            |
| Spreadsheet | Automation of both Spreadsheet and Sheet                                           |
| Triggers    | Event Type Actions.<br/> [Click Here to Learn More][trigger]                       |


## Getting Started
### Prerequisites
- Creation of [Google Account][googleaccount]
- [Google App Script][googleappscript]


### Running Script
[Documentations](documentation/docs.pdf)

## Acknowledgement
- [Google App Script][googleappscript]
- [Google App Script Documents](https://developers.google.com/apps-script/reference)


[googleappscript]: https://www.google.com/script/start/
[googleaccount]: https://accounts.google.com/v3/signin/identifier?dsh=S1108141963%3A1690474069932711&continue=https%3A%2F%2Fmyaccount.google.com%3Futm_source%3Daccount-marketing-page%26utm_medium%3Dgo-to-account-button&ifkv=AeDOFXgx6tIotZMmgy1NnzbHeU-x2KFrkigZLFTvWFuCru4WBC4wHkWDo8qnr6Fi_mWZ5ZIVXevlZg&service=accountsettings&flowName=GlifWebSignIn&flowEntry=ServiceLogin
[trigger]: https://developers.google.com/apps-script/guides/triggers
