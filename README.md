# 🔍 Excel-3

![Status](https://img.shields.io/badge/status-active-brightgreen.svg)
![Excel](https://img.shields.io/badge/Microsoft-Excel-blue.svg)

## ✨ Project Description

**Excel-3** is a guide to the Find & Select feature in Excel. Here you'll find practical tips, instructions, and illustrations about finding, replacing, and removing values.

> 📚 **Goal:** Help you master Excel's Find, Replace, Go To Special, and related tools—ideal for all users!

---

## 📒 Table of Contents

- [Find](#-find)
- [Replace](#-replace)
- [Go To Special](#-go-to-special)
- [Wildcards](#-wildcards)
- [Delete Blank Rows](#-delete-blank-rows)
- [Row Differences](#-row-differences)
- [Screenshots](#-screenshots)
- [Requirements](#-requirements)
- [Author](#-author)

---

## 🔎 Find

1. On the Home tab, in the Editing group, click Find & Select. (Or press **CTRL+F**)
2. Click Find.

![Find](Screenshots/Find.png)

The 'Find and Replace' dialog box will appear.

3. Type the text you want to find. For example, type Melon.

![Find1](Screenshots/Find1.png)

Excel selects the first occurrence.

4. Click Find Next

![Find2](Screenshots/Find2.png)

5. To get a list of all the occurrences, click 'Find All'.  
*(Note: Excel highlighted "Watermelon" even though we wanted to search the word "Melon". To only get exact match, go to Options in Find and Replace dialog box and select "Match entire cell contents.")*

![Options](Screenshots/Options.png)

---

## 🔄 Replace

1. On the Home tab, in the Editing group, click Find & Select.
2. Click Replace.
3. The 'Find and Replace' dialog will appear.
4. Type the text you want to find and replace it with.
5. Click Find Next.

![Replace](Screenshots/Replace.png)

6. Click Replace to make a single replacement.

![Replace1](Screenshots/Replace1.png)

*(Note: Use Replace All to replace all occurrences.)*

---

## 🧭 Go To Special

Use Excel’s Go To Special feature to quickly select all cells with data validation, formulas, conditional formatting, etc.  
Here we select cells with formulas.

1. Select range of cells.
2. On the Home tab, in the Editing group, click Find & Select.
3. Click Go To Special.
4. Select Formulas and click OK.

![Special](Screenshots/Special.png)

Excel selects all cells with formulas.

---

## 🃏 Wildcards

There are 3 wildcards in Excel:  
- **?** matches exactly one character  
- **\*** matches zero or more characters  
- **~** matches wildcards (for example, `~?` finds a literal question mark).

---

## 🗑️ Delete Blank Rows

1. On the Home tab, in the Editing group, click Find & Select.
2. Click Go To Special.
3. Select Blanks and click OK.

![Special1](Screenshots/Special1.png)

4. On the Home tab, in the Cells group, click Delete then Delete Sheet Rows.

![Delete](Screenshots/Delete.png)

**Result:**

![Result](Screenshots/Result.png)

*(Note: this method also deletes rows with one or more blank cells.)*

Now let's remove only rows that are completely empty:

1. Insert COUNTA function and drag it down.

![Counta](Screenshots/Counta.png)

When COUNTA returns 0 it means the row is empty. We need to filter these rows.

2. Select cell E1.
3. On the Data tab, in the Sort & Filter group, click Filter. Arrows in column headers will appear.

![Filter](Screenshots/Filter.png)

4. Click the arrow in E1 column.
5. Click on Select All to clear all the check boxes, and click the check box next to 0.

![Filter1](Screenshots/Filter1.png)

Click OK.

**Result:**  
![Result1](Screenshots/Result1.png)

Delete these rows, then click Filter again to remove it.

**Result:**  
![Result2](Screenshots/Result2.png)

---

## ➡️ Row Differences

1. Select range H1:J7  
   *(Note: because we selected the range H1:J7 by clicking on cell H1 first, cell H1 is the active cell. As a result, the comparison cells are in column H.)*
2. On the Home tab, in the Editing group, click Find & Select.
3. Click Go To Special.
4. Select Row differences and click OK.

![Special2](Screenshots/Special2.png)

**Result:**  
![Result3](Screenshots/Result3.png)

**Colored differences:**  
![Result4](Screenshots/Result4.png)

---

## 📷 Screenshots

You can find all screenshots in the `/Screenshots` folder.

---

## ℹ️ Requirements

- Microsoft Excel (recommended: 2021/365 for best compatibility)

---

## 👨‍💻 Author

Project and documentation by **Kuba27x**  
Repository: [Kuba27x/Excel-3](https://github.com/Kuba27x/Excel-3)

---
