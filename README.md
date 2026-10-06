# AdaptiveEdu

DEPI Graduation Project — a Telegram study bot for Egyptian 1st Secondary.

**GitHub (fixed URL):** https://github.com/ahmedhamedali82/AdaptiveEdu

## Idea

The student sends an ID. The bot sends **the lesson that is due**, then a short quiz.

- **Pass (5 or more / 9)** → next lesson  
- **Fail** → same lesson, simpler text  
- Scores are saved in Google Sheets  

## Two Google Sheets

1. **Students**  
   https://docs.google.com/spreadsheets/d/17A-hQ6jQ7QJrcZ4f5f02O82PKKAtQ2gTiJHtv71LWI4/edit?usp=sharing  
   Tab: `Students` — n8n matches on `studentID`

2. **Curriculum**  
   https://docs.google.com/spreadsheets/d/10JNorDKgaesXGobuq_9sHipzRp-DOjxGAFZEV5D7Hwk/edit?usp=sharing  
   Tabs: `Subjects` (read) · `Lessons` (read) · `Results` (append only)

## How the student uses it

1. Send `S1-1001`  
   Or `S1-1001 Integrated Sciences`  
2. Bot sends lesson text + 3 questions  
3. Reply with three letters: `C B A`  
4. Bot sends Pass / Fail and writes the sheets  

## n8n

**Delivery: 2 weeks.**

Import `AdaptiveEdu_Telegram_Bot.json`.

Then connect:

- Telegram bot credential on **Telegram Trigger**, **Send lesson + quiz**, **Send result**  
- Google Sheets credential on all Sheets nodes  

Keep tab names exactly: `Students`, `Lessons`, `Results`.

## Files

- `AdaptiveEdu_Project_Plan.pptx` — 2-week presentation for the panel  
- `AdaptiveEdu_Telegram_Bot.json` — n8n workflow  
- `AdaptiveEdu_Students.xlsx` / `AdaptiveEdu_Curriculum.xlsx` — source files for the two sheets  
