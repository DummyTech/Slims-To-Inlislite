# 📌 Important Notes
Due to a formatting issue in the script, you **WILL HAVE TO RESAVE THE RESULTS** using spreadsheet software like **WPS Office** or **Microsoft Excel** for it to work correctly.

# ➕ Additional Notes
- Exported files from SLiMS/Senayan must be placed in the same directory as the script.
- Loading data into Inlislite may take some time, but it will work.
- Exported book data is split into batches of 100 records to reduce load time during import to Inlislite.

# Why choose this method instead of swapping columns manually in Excel?
### ✅ Advantages
The advantages of using this script instead of manually rearranging data columns in spreadsheet software are:
- This software operates locally and modify the export file, not the database php like other method. So if there is a fatal error, at most, it will only break the export file, not the database.
- Once the column order is set, there's no need to redo it in the future.
- Column rules can be customized in the code to fit your needs, including support for complex conditions or formulas written in python.
- Automatically formats the output to match the Inlislite import format, which is required for successful data import.

### ⚠️ Limitations & Nice-to-have
Here are some known limitations and potential areas for improvement:
- The script does not migrate photos or snippets of books and members.
- Require basic python language knowledge to edit each column formula.
- Data can't be directly imported into Inlislite due to improperly format data.

At the end of the day, it all depends on your specific needs.

---
### Thank you for using our software! 🎉🎉🥳🥳
### DummyTech >u<
