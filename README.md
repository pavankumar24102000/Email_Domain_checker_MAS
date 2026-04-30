# 📧 Email Domain Checker

A simple tool that automatically checks a list of email addresses to tell you which ones are **safe**, which are **risky/disposable**, and which are **unknown** — and saves the results in a nice color-coded Excel file.

---

## 🤔 What does it do? (In Plain English)

Imagine you have an Excel sheet with hundreds of email addresses, and you want to know which ones are real, trustworthy emails versus fake or temporary ones. Doing this manually would take hours.

This program does it for you, automatically:

1. **Reads** your Excel file full of email addresses.
2. **Opens** a website (verifymail.io) for each email and checks its status.
3. **Records** the result for each email.
4. **Saves** everything into a new Excel file with colors so you can see results at a glance.

Think of it as a robot assistant that opens a website hundreds of times and copies down the answer for you.

---

## 🎨 What the colors mean

When the program finishes, you'll get an Excel file where each email is highlighted:

| Color | Meaning | Examples of Status |
|-------|---------|--------------------|
| 🟢 **Green** | Good email — safe to use | Safe, Valid, Deliverable, Good |
| 🔴 **Red** | Bad or risky email — avoid | Temporary, Disposable, Invalid, Risky, Failed |
| 🟡 **Yellow** | Unclear — needs manual review | Anything that doesn't match the above |

---

## 📋 What you need before running it

1. **Java** installed on your computer.
2. **Google Chrome** browser installed.
3. An **Excel file (.xlsx)** with the email addresses listed in **Column A** (first column).
4. The required libraries: **Selenium** (for browsing) and **Apache POI** (for Excel).

---

## 📁 File locations (you may need to change these)

The program currently looks for files in these spots:

- **Input file** (your list of emails): `C:\Users\pavan.k1\Downloads\Domain check 30 apr.xlsx`
- **Output file** (the results): `C:\Users\pavan.k1\Downloads\output.xlsx`

👉 If your computer username or file location is different, open the code and change these two lines:

```java
String inputPath = "C:\\Users\\YOUR_NAME\\Downloads\\YOUR_FILE.xlsx";
String outputPath = "C:\\Users\\YOUR_NAME\\Downloads\\output.xlsx";
```

---

## ▶️ How to run it

1. Put your email list in an Excel file with emails in **Column A** (one email per row).
2. Update the file paths in the code (see above).
3. Run the program.
4. A Chrome window will open and start checking each email automatically — **don't close it!**
5. Wait for it to finish. You'll see progress in the console.
6. When done, open the `output.xlsx` file in your Downloads folder.

---

## 🖥️ What you'll see while it runs

While the program is working, your console (the black text window) will show something like this:

```
=============================================================
Email                                              Status
=============================================================
john@example.com                                   Safe to send       [🟢 GREEN]
fake@tempmail.com                                  Disposable email   [🔴 RED]
test@unknown.org                                   Unrecognized       [🟡 YELLOW]
...
=============================================================
✅ SUMMARY
=============================================================
  🟢 Green  (Safe/Valid)       : 45
  🔴 Red    (Temporary/Invalid): 12
  🟡 Yellow (Unknown/Other)    : 8
  📧 Total Processed           : 65
=============================================================
✅ Output saved to: C:\Users\pavan.k1\Downloads\output.xlsx
```

At the end, you get a clean summary telling you how many emails fall into each category.

---

## ⚠️ Things to keep in mind

- **Don't close the Chrome window** while the program is running — it needs it open to do its work.
- **Internet required** — it checks each email through a website.
- **It takes time** — about 5–10 seconds per email. So 100 emails ≈ 10–15 minutes.
- **The website might block too many quick requests.** If it stops working, wait a while before trying again.
- **Emails marked yellow** should be reviewed by hand — the program wasn't sure about them.

---

## 🛠️ Troubleshooting

| Problem | What to do |
|---------|------------|
| "File not found" error | Check the file path and make sure your input Excel file actually exists there. |
| Chrome doesn't open | Make sure Chrome is installed and ChromeDriver matches your Chrome version. |
| Many emails come back as "FAILED TO LOAD" | The website may be slow or blocking you. Try again later or with fewer emails. |
| Output file won't open | Close any Excel windows that already have `output.xlsx` open, then re-run. |

---

## 📝 Summary

**Input:** An Excel file with email addresses.
**Output:** An Excel file telling you which emails are safe (green), risky (red), or unclear (yellow), plus a count of each.
**Effort:** Click run and walk away. ☕

That's it!
