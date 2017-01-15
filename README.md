# AttendanceManagerMonster

## The rationale behind it

On Aug 2015 ["she codes;"](http://www.she-codes.org/) launched the program "she codes; Academy" - free-of-charge Web and Android study courses for women without any previous programming background. The need for a well-designed software that automatically sends learning tasks and monitors the student’s progress, arose during the very first week of "she codes; academy" launch.

*AttendanceManagerMonster*, a GUI app for managers is used for:

1.	Analyzing students’ progress.
2.	Sending weekly emails to students.
3.	Modifying attendance files

and more...

It was created for "she codes;" WIS branch on Feb 8, 2016.  
*AttendanceManagerMonster* and also *[TasksMonster](https://github.com/DarLador/TasksMonster)* applications were the source of inspiration for the latter development of [shecodesconnect.com](https://shecodesconnect.com/) - which is an implementation of this platform at national level.

## The application

*AttendanceManagerMonster* application was built in Anaconda Python 3.5 on windows OS.

## Application process at a glance

**Intro screen** -- The manager fills her name and password.

<img width="250" alt="intro" src="https://cloud.githubusercontent.com/assets/17408143/21900607/2241efac-d8fd-11e6-9b0b-a858ca3b7517.PNG">

If the name & password are correct the app will recognize the manager and selects students from the manager’s related study group (web or android).

The opened frame consists of 8 tabs. Below is a short description for each tab.

### 1: Summary
Shows a general description of each student (progress, absences etc.).

<img width="446" alt="1" src="https://cloud.githubusercontent.com/assets/17408143/21900620/2d0634b6-d8fd-11e6-81e6-ff6ac0cc21d6.png">

### 2: Detailed report
In this tab the app sorts students by their class and shows their attendance details (presences/absences).

<img width="446" alt="tab2_detailedreport_tree" src="https://cloud.githubusercontent.com/assets/17408143/21900629/38471c14-d8fd-11e6-891c-0fa29bc0060b.PNG">

For example:

<img width="447" alt="3" src="https://cloud.githubusercontent.com/assets/17408143/21900639/42273048-d8fd-11e6-80d8-b5ffa698eb85.png">

### 3: Meetups
In this tab the app shows student's attendance from the very first day of she codes academy. We can filter the table by parameters such as: “Was absent in the recent 1/2/3/4 weeks / one month and more” or by “Number of absences (0-31)”

<img width="446" alt="4" src="https://cloud.githubusercontent.com/assets/17408143/21900650/48910a3a-d8fd-11e6-99af-20490d5377e8.png">

For example:

<img width="447" alt="5" src="https://cloud.githubusercontent.com/assets/17408143/21900658/520dbec8-d8fd-11e6-9b86-bfa8f7e70366.png">

Let’s say we want to send an email to Tami and ask her how she is doing. We can select her name and click on the ‘email’ button. Then we can compose the email in a new frame.

<img width="644" alt="6" src="https://cloud.githubusercontent.com/assets/17408143/21900665/5666be66-d8fd-11e6-9bcf-086b97911265.png">

### 4: Progress report
Here the app sorts out students by the class they’re in and for each student it shows the date when each study task were sent to them. This screen is highly important as for the manager can identify students who spend too much time on the same lesson and help them out.

<img width="446" alt="7" src="https://cloud.githubusercontent.com/assets/17408143/21900672/5e8f119c-d8fd-11e6-8669-076f60907b52.png">

### 5: Complete lessons
Students who miss a meetup can complete it in another branch. Let’s say Netaly can’t attend to the meetup on 17/2/2016. Instead she visited the TLV – Check Point Building branch (Web) on 14/02/2016.

<img width="445" alt="8" src="https://cloud.githubusercontent.com/assets/17408143/21961277/14ef575a-db0e-11e6-8611-d857debfe8fd.png">

After we fill the data and click on the Submit button, the app gets refreshed.

This is how the attendance report looks after this change:

<img width="446" alt="9" src="https://cloud.githubusercontent.com/assets/17408143/21900684/66a23670-d8fd-11e6-962d-439c84e795ac.png">

### 6: Delete/restore
If a student no longer attends classes, the managers can remove her from the app. The student will not be permanently deleted from the database. Therefore, we can restore her data if she returns.

<img width="446" alt="10" src="https://cloud.githubusercontent.com/assets/17408143/21900692/6d6eeaca-d8fd-11e6-9adf-44c11b22df50.png">

### 7: Weekly emails
This feature enables sending emails to current students.

<img width="447" alt="tab6_weekly emails" src="https://cloud.githubusercontent.com/assets/17408143/21901334/f1144da0-d8ff-11e6-9157-2d8959241c2a.PNG">

Managers must define how recent of attendance qualifies as “current students".

<img width="448" alt="tab6_send_to_list" src="https://cloud.githubusercontent.com/assets/17408143/21901342/f9a8b47e-d8ff-11e6-9a0e-c1fb68e9746e.png">

### 8: Notes
General notes about the app.
