# Google-Apps-Script
Repository holding scripts to automate gmail use cases to make your gmail account feel professional

## Newslender Automation
This Google Apps Script automates the process of sending personalized newsletters via Gmail. It reads a draft email containing "Haptigation Newsletter" in the subject line and sends it to all contacts with a specific label, personalizing each email with the recipient's name.
Implementing it this way ensures data safety and allows for personalization in the newsletter.

Check out the detailed documentation here: [Docusaurus](https://haptigation-student-project.github.io/Documentation/docs/Software/Google%20Apps%20Script/newsletter-automation)

---

## E-Mail Acknowledgement Auto-Responder
This Google Apps Script responds to any message ingoing that hasn't been read within 1 minute. It reads a draft email containing "E-Mail Response Automation Draft - DO NOT DELETE" in the subject line. 

Check out the detailed documentation here: [Docusaurus](https://haptigation-student-project.github.io/Documentation/docs/Software/Google%20Apps%20Script/email-acknowledgement-automation)

---

## User Deletion Automation
This Google Apps Script must be triggered manually. It deletes EVERYTHING in GMAIL. The contact, all ingoing emails and all outgoing emails. Before deletion it sends out a confirmation email.

Check out the detailed documentation here [Docusaurus](https://haptigation-student-project.github.io/Documentation/docs/Software/Google%20Apps%20Script/user-deletion-automation)

---

## Unsibscribe Newsletter Automation
This Google Apps Script automates the removal of the "Newsletter Subsriber" label from a contact. It does this by reading a [Google Forms](https://docs.google.com/forms/d/e/1FAIpQLSd9MLAY40kCpw3iFn5atip4SO0vLRMWJK3G2-bQKwDecgRUFg/viewform?usp=sharing&ouid=101591382494264700789), checks for the inserted email and removes the label. Afterwards a confirmation email is being sent.

Check out the detailed documentation here [Docusaurus](https://haptigation-student-project.github.io/Documentation/docs/Software/Google%20Apps%20Script/unsubscribe-newsletter-automation)

## Attachments for Docusaurus
### Newsletter Automation
<img width="1227" height="134" alt="save icon" src="https://github.com/user-attachments/assets/d5c373b4-24ea-4bc5-aea0-cbb9c12048b4" />
<img width="1227" height="144" alt="run function" src="https://github.com/user-attachments/assets/d0c9576f-31a4-489e-b2e4-4e1a9b0df6c7" />
### Auto Responder
<img width="1302" height="127" alt="save icon" src="https://github.com/user-attachments/assets/9f713499-f400-4042-8c1b-537d62180816" />
<img width="1223" height="141" alt="run function setup" src="https://github.com/user-attachments/assets/1fa2cedf-4218-4e49-bcb0-c0d7946974e2" />
<img width="366" height="406" alt="trigger icon" src="https://github.com/user-attachments/assets/e68ba06e-6fba-4c1f-a721-9042d59b4796" />
<img width="1227" height="137" alt="run function remove" src="https://github.com/user-attachments/assets/21a86e3f-8425-453f-9fb3-1490296bdf72" />
### Unsubscribe Newsletter
<img width="1222" height="140" alt="save icon" src="https://github.com/user-attachments/assets/98219f3c-a5d6-4209-ab6c-5ecf8cab1471" />
<img width="1262" height="135" alt="run function setup" src="https://github.com/user-attachments/assets/e3403182-3f40-412a-bd85-8ffd3a148921" />
<img width="1896" height="389" alt="trigger" src="https://github.com/user-attachments/assets/b0dd5e15-da57-4df9-b22e-b5ef68b83d0e" />
<img width="1220" height="132" alt="run function remove" src="https://github.com/user-attachments/assets/f1484a30-633c-4e13-b597-5ebba7e86ceb" />
### User Data Deletion
<img width="1230" height="141" alt="save icon" src="https://github.com/user-attachments/assets/54b4e907-aac5-48b8-9eaa-e87428e4a31b" />
