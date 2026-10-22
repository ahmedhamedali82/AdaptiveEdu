# AdaptiveEdu

Adaptive Telegram-based learning system - DEPI Graduation Project (Final version, 22 October 2026).

A student registers once on the Telegram bot, chooses a subject, receives an AI-generated lesson as a PDF by e-mail with a 10-question quiz, uploads his answers as an Excel file, and gets graded automatically. Pass (more than 6/10) unlocks the next lesson; Fail keeps the same lesson open for a retake. A new lesson cannot be requested while a previous one is unanswered, and a new subject can always be started from lesson 1.

- Bot: https://t.me/AdaptiveEdubot
- Workflow: `AdaptiveEdu.json` - n8n workflow ADAPTIVE_14 (48 nodes). Before importing, put your own PDFShift API key in the HTTP Request node (the committed file contains `YOUR_PDFSHIFT_API_KEY`).

## Team

| Team member | E-mail |
|---|---|
| Ehab Mohamed Abdellatif Zaki (Team Leader) | eh_latif@hotmail.com |
| Ahmed Hamed Mohamed Elfeky | ahmed.hamedmohamed1@gmail.com |
| Aamn Ehab Sayed Mohamed | aamnehab100@gmail.com |
| Abdelrahman Khaled Abdelmenium Younis | abdelrahmanyonis1@gmail.com |
| Mahmoud Abdelsalam Elsayed Alsisi | alsisiagric84@live.com |

## Final materials in this repository

| File | Description |
|---|---|
| `AdaptiveEdu_Demo.mp4` | Final demo video (4:21, 1080p) with synchronised Egyptian male voice-over |
| `AdaptiveEdu_Project_Documentation.docx` | Full project documentation (architecture, sheets, workflow steps, live evidence) |
| `AdaptiveEdu_Project_Plan.pptx` | Project presentation (19 slides) |
| `AdaptiveEdu.json` | Final n8n workflow export (sanitised - no secrets) |
| `Students_Data.xlsx` | Students sheet (codes, Telegram IDs, e-mails, status) |
| `Lessons_Data.xlsx` | Curriculum: 8 subjects and their ordered lessons |
| `student_progress.xlsx` | Progress log: answer keys, grades, success flags |
| `ST005_8_1.xlsx`, `ST005_8_2.xlsx` | Answer files used in the final live demo |
| `Screenshots/` | 17 screenshots from the final live run (bot, sheets, Drive, e-mails, PDF) |

## Live demo scenario (as recorded)

1. ST005 registers - Telegram chat ID saved automatically, subjects list received.
2. Subject 8 (Physical Education), lesson 1 Sports Skills - PDF + quiz by e-mail.
3. First answers graded 4/10 (Fail) - retake with corrected answers graded 10/10 (Pass).
4. Lesson 2 (Health & Movement) delivered automatically; wrong answers scored 5/10 (Fail).
5. Requesting another lesson while lesson 2 is open is refused by the bot.
6. Subject 7 starts Programming & AI from lesson 1 (Coding Fundamentals).
