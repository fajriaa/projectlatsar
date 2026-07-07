import json
from app import parse_ocr_fields

# Simulated OCR output combining the screenshots you provided
sample_text = '''Extraction Name: Evidence Pack (1)
Model name and specifications: Samsung SM-P205
Extraction Type: File System
Extraction Start and End Date/time: 20/02/2026 11:21:11(UTC+7)
Path: F:\\DATA KERJA\\2026\\...\\
Image Hash: Verified 23/02/26

Content
Device Data
Chats                157
Contacts             862
Installed Applications    373(12)

Instant Messages     5
User Accounts        21
Timeline             Tidak ditemukan
Applications         2622
Archives             273
Databases            1376
Images               34609
Text                 6908

Device Data Files
All Files            113,362
Applications         5,499
Images               57,853 (4,788)
Text                 10,142
Videos               452
'''

res = parse_ocr_fields(sample_text)
print(json.dumps(res, indent=2, ensure_ascii=False))
