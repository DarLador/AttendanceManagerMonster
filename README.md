# AttendanceManagerMonster

## The rationale behind it

On Aug 2015 "she codes;" launched "she codes; Academy" - free-of-charge Web and Android study tracks for women without any previous programming background. The need for well designed software which automatically sends tasks and monitors the students progress arose in the very first day of "she codes; academy".

AttendanceManagerMonster is a GUI app for managers is used:

1. To analyse students progress.
2. To send weekly emails to students.
3. Modify attendance files

and more...

It was created for "she codes;" WIS branch on Feb 8, 2016. AttendanceManagerMonster and [TasksMonster](https://github.com/DarLador/TasksMonster) applications were the source of inspiration for shecodesconnect.com - implementation of this platform at national level.

## The application

AttendanceManagerMonster application was built in Anaconda Python 3.5 on windows. As soon as time permits I will upload a new version with related files so it will work on any PC.

## Application process at a glance

**Intro screen** -- The manager fills her name and password.

<img width="250" alt="intro" src="https://cloud.githubusercontent.com/assets/17408143/21900607/2241efac-d8fd-11e6-9b0b-a858ca3b7517.PNG">

If name & password are correct the app recognize the manager and selects students from the manager’s group (web or android). 

The opened frame consists of 8 tabs. Below is a short description for each tab.
### 1: Summary
Shows general description of each student (progress, absences etc.).

<img width="446" alt="1" src="https://cloud.githubusercontent.com/assets/17408143/21900620/2d0634b6-d8fd-11e6-81e6-ff6ac0cc21d6.png">

### 2: Detailed report
In this tab the app sort students by group and show the present / absent details.

<img width="446" alt="tab2_detailedreport_tree" src="https://cloud.githubusercontent.com/assets/17408143/21900629/38471c14-d8fd-11e6-891c-0fa29bc0060b.PNG">

For example:

<img width="447" alt="3" src="https://cloud.githubusercontent.com/assets/17408143/21900639/42273048-d8fd-11e6-80d8-b5ffa698eb85.png">

### 3: Meetups
In this tab the app shows student's attendance from the very first day of she codes academy. We can filter the table by: “Was absence in the recent 1/2/3/4 weeks / one month and more” or by “Number of absences (0-31)” 

<img width="446" alt="4" src="https://cloud.githubusercontent.com/assets/17408143/21900650/48910a3a-d8fd-11e6-99af-20490d5377e8.png">

For example:

<img width="447" alt="5" src="https://cloud.githubusercontent.com/assets/17408143/21900658/520dbec8-d8fd-11e6-9b86-bfa8f7e70366.png">

Let’s say we want to send an email to Tami and ask her how she is. We can select her name and click on the ‘email’ button. Then we can compose the email in a new frame.

<img width="644" alt="6" src="https://cloud.githubusercontent.com/assets/17408143/21900665/5666be66-d8fd-11e6-9bcf-086b97911265.png">

### 4: Progress report
Here the app sorts students by group and for each student it shows the date when each task was sent. This screen is highly important as the manager can identify the students who spend too much time on the same lesson.

<img width="446" alt="7" src="https://cloud.githubusercontent.com/assets/17408143/21900672/5e8f119c-d8fd-11e6-8669-076f60907b52.png">

### 5: Complete lessons
Students who miss class can complete it in another branch. Let’s say Netaly can’t attend to the class on 17/2/2016. Instead she visited TLV – Check Point Building (Web) on 14/02/2016.

<img width="445" alt="8" src="https://cloud.githubusercontent.com/assets/17408143/21900676/62de0b18-d8fd-11e6-95a1-0cc834c5e531.png">

After we fill the data and click on Submit button, the app is refreshed. 

This is how the attendance report looks after this change:

<img width="446" alt="9" src="https://cloud.githubusercontent.com/assets/17408143/21900684/66a23670-d8fd-11e6-962d-439c84e795ac.png">

### 6: Delete/restore
If a student no longer attends classes the managers can remove her. A student will not be permanently deleted. Therefore we can restore her data if she returns.

<img width="446" alt="10" src="https://cloud.githubusercontent.com/assets/17408143/21900692/6d6eeaca-d8fd-11e6-9adf-44c11b22df50.png">

### 7: Weekly emails
Enables sanding emails to current students.

<img width="447" alt="tab6_weekly emails" src="https://cloud.githubusercontent.com/assets/17408143/21901334/f1144da0-d8ff-11e6-9157-2d8959241c2a.PNG">

Managers must define the term "current students".

<img width="448" alt="tab6_send_to_list" src="https://cloud.githubusercontent.com/assets/17408143/21901342/f9a8b47e-d8ff-11e6-9a0e-c1fb68e9746e.png">

### 8: Notes
General notes about the app.
