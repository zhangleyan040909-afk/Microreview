# Microreview

Microreview is a local Windows desktop app for reviewing student practice materials and exporting score sheets.

## Main Features

- Import a student CSV with `姓名` and `学号`.
- Group students by grade using the first two digits of the student ID, for example `24721001 -> 2024级`.
- Select a local materials folder and auto-match files by:
  - student ID in the file name
  - related keywords for questions 1-8
- Review 8 questions for each student.
- Record completion status, score, and notes.
- Export all scores to CSV.

## Material Matching

The app ignores spaces, full-width spaces, underscores, and hyphens when matching file names.

Each file name should include the student ID and one of the keywords below:

- Q1: 爱国主义教育实践, 爱国主义, 爱国
- Q2: 志愿服务实践, 志愿服务, 志愿
- Q3: 劳动与生活实践, 劳动实践, 劳动
- Q4: 管理服务实践, 管理服务, 管理
- Q5: 勤工助学实践, 勤工助学
- Q6: 应急救护实践, 应急救护
- Q7: 跨文化交流实践, 跨文化交流
- Q8: 参与国家重大活动实践, 国家重大活动

Examples:

```text
24721001_张三_爱 国主义 教育实践.txt
张三 24721001 志愿-服务.docx
24721001　劳动_实践.pdf
```

## Running the Packaged App

Use the packaged directory:

```text
dist/Microreview/Microreview.exe
```

Do not move `Microreview.exe` out of the `dist/Microreview` folder because it depends on the `_internal` directory.

## Build From Source

Requirements:

```powershell
python -m pip install customtkinter pyinstaller pillow
```

Build:

```powershell
pyinstaller --noconfirm --onedir --windowed --name Microreview --icon microreview.ico desktop_app_modern.py
```

The output will be:

```text
dist/Microreview/Microreview.exe
```

## Notes

The app stores local working data in:

```text
%APPDATA%\StudentReviewApp\review_data.json
```

This keeps existing user data compatible with earlier builds.
