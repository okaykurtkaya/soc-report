# SOC REPORT

### INTRODUCTION

If you have a ticket system created with **Redmine** or if you want to report a watch list for your **Zabbix** servers or if you want to create a monthly SOC report covering all of them, you can review my project.

![RESULT](https://github.com/okay-kurtkaya/soc-report/blob/main/REPORT-RESULTS/12.png)

### PROGRAM REQUIREMENTS

* Before starting the program, the directory where you run the program should have the **images** folder containing the image files and the main **Word** file.

* If a new customer arrives, this is not a problem for the program and a monthly report is generated. While creating the monthly report, you should put this logo in the images folder, since each customer has a logo. When saving the customer's logo to the images folder, make sure that it is in the following format. For example, let's assume that a company like Cyber Gladiators has been added to our system. In this case, the logo should be saved in the images folder as follows; It should be in the form of siber_gladiators.png.
  * If there are Turkish characters in the customer name as in the example above, use the English alphabet structure instead.
  * The customer name can be more than two words. In this case, when registering the logo, use **underscore** (_) instead of spaces, as in the upper part.

### HOW TO USE THE PROGRAM ?

1. In the login section of the program, you must log in with your **username** and **password** belonging to the ticket system.

2. After logging in, you will be asked for the date range for which reports will be generated:
    * _start_date part_: The start date of the report. This date should be specified as **year-month-day**. For example, when you want to receive a report for March 2023, your start date should be as follows; **2023-3-1**
    * _last_date part_: The end date of the report. This date should be specified as **year-month-day**. For example, when you want to receive a report for March 2023, your end date should be as follows; **2023-3-31**
    * _base docx template name part_: Word file containing codes, used while generating the report. If any changes are made to the codes in this file, the structure may be completely broken. In this case, the program cannot generate the monthly report. Do not make any changes to this file. In this part of the program, the name of the main Word file is requested. When entering the name of this file, you must enter it with the extension of the file. For example, you should type **SOC-REPORT-TEMPLATE.docx**.
    * After entering the desired values correctly, the program will start to generate the monthly report of **each customer in a different folder**.


### AND THE END

This is my first project. Surely there may be some **deficiencies** or **excesses**. I will improve them as I see them over time. I hope this project will be a helpful resource for you.

`printf(n3gat1v3o)`
