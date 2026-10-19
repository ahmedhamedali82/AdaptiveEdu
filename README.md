# AdaptiveEdu

DEPI Graduation Project — live Telegram study bot for Egyptian 1st Secondary.

**Bot:** https://t.me/AdaptiveEdubot  
**GitHub:** https://github.com/ahmedhamedali82/AdaptiveEdu

## Idea

1. Student sends ID (`ST005`)  
2. Chooses subject code `1–8`  
3. Gemini builds the next lesson as **PDF + 10 MCQs**  
4. Student receives the **PDF attached by email** and uploads Excel answers (`ST005_4_1.xlsx`)  
5. **Pass if grade > 6 / 10** · else **Fail** and restudy the same lesson  

Results go by **Gmail**. Progress is saved in Google Sheets.

## Three Google Sheets

1. **Students_Data** — https://docs.google.com/spreadsheets/d/1ddIwenYjiwKryP7Pogxe-7qsjKnTmtfCtAlDoplSn5Q  
2. **Lessons_Data** — https://docs.google.com/spreadsheets/d/1akeM65odle57kxtR7M5LqFMnxCsHX3M1cJKKltExFEg  
3. **student_progress** — https://docs.google.com/spreadsheets/d/1II6ckyiKFucKsC38PC_0a9tH1h8H0JZe3qaclLxCExY  

## Excel answer file

- Name: `ST{id}_{subject}_{lesson}.xlsx` (example `ST005_4_1.xlsx`)  
- One row, 10 capital letters, **no header**

## n8n

Import `AdaptiveEdu.json` (workflow **ADAPTIVE_14**, ~46 nodes).

Connect Telegram · Google Sheets · Gmail · Drive · Gemini · PDFShift (HTTP Request) credentials.

Lesson PDF path: Gemini Markdown → Format Lesson HTML → PDFShift convert → Gmail **PDF attachment**.

## Files

- `AdaptiveEdu_Project_Plan.pptx` — panel slides (journey + n8n steps + live shots)  
- `AdaptiveEdu_Project_Documentation.docx` — full project documentation  
- `AdaptiveEdu_Video_demo_steps.docx` — 2–5 min video demo steps (EN + AR)  
- `AdaptiveEdu.json` — n8n workflow  
- `diagrams/demo/` — bot, lesson email, Pass/Fail emails, Excel style, canvas  
