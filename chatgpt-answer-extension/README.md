# AutoBid ChatGPT Answer Worker

Chrome extension that runs on a ChatGPT tab, reads pending AutoBid questions from Google Sheets through your Apps Script Web App, asks ChatGPT for answers, and saves JSON answers back to the same row.

## Setup

1. Paste `../docs/apps-script-autobid-bridge.gs` into the existing Apps Script project.
2. Add the `autobidListJobs`, `autobidSaveQuestions`, `autobidReadAnswers`, and `autobidSaveAnswers` branches to `doPost(e)`.
3. Deploy a new Apps Script Web App version.
4. Load this folder as an unpacked Chrome extension:

```text
~/Documents/AutoBid/chatgpt-answer-extension
```

5. Open ChatGPT in Chrome.
6. In this extension popup, set:

```text
Apps Script Web App URL
Extension Secret
Sheet tab
Start row
End row
```

7. Click `Run once`, or `Start loop`.

## Expected Sheet Flow

The job-page AutoBid extension writes pending fields to:

```text
autobid_questions
```

This worker writes answers to:

```text
autobid_answers
```

The value saved in `autobid_answers` has this shape:

```json
{
  "answers": [
    { "field_id": "ab_12_example", "value": "Answer text" }
  ]
}
```
