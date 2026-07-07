import re
import json

def parse_ocr_fields(text):
    data = {}
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    pairs = {}
    current_section = None
    for i, ln in enumerate(lines):
        if re.search(r'device data files', ln, re.IGNORECASE):
            current_section = 'files'
            continue
        if re.search(r'device data\b', ln, re.IGNORECASE):
            current_section = 'device'
            continue
        m = re.match(r'^([A-Za-z0-9\s\./\\\-]{2,60})\s*[:\-]?\s{1,}\s*(.+)$', ln)
        if m:
            label = m.group(1).strip()
            val = m.group(2).strip()
            pairs[label] = {'val': val, 'section': current_section, 'idx': i}
            continue
        m2 = re.match(r'^([A-Za-z\s]{2,40})\s+([0-9][0-9\.,\(\)\s]+)$', ln)
        if m2:
            pairs[m2.group(1).strip()] = {'val': m2.group(2).strip(), 'section': current_section, 'idx': i}
            continue
        cols = re.split(r'\s{2,}', ln)
        if len(cols) >= 2:
            left = cols[0].strip()
            right = cols[-1].strip()
            pairs[left] = {'val': right, 'section': current_section, 'idx': i}

    def find_in_pairs(keys):
        for k in keys:
            for label in pairs:
                if k.lower() in label.lower():
                    return pairs[label]
        return None

    name = find_in_pairs(['Extraction Name', 'Nama Instansi', 'Nama', 'Owner', 'Device'])
    if name:
        data['nama_instansi'] = name

    alamat = find_in_pairs(['Path', 'Alamat', 'Address', 'Location'])
    if alamat:
        data['alamat_instansi'] = alamat

    merek = find_in_pairs(['Model name', 'Model name and specifications', 'Merek/Model', 'Merek', 'Model', 'Device'])
    if merek:
        data['merek_model'] = merek

    imei = find_in_pairs(['IMEI', 'Imei'])
    if imei:
        data['imei'] = re.sub(r'\s+', '', imei)

    tel = find_in_pairs(['Telp', 'Telepon', 'Phone'])
    if tel:
        data['telepon'] = tel

    stats_map = {
        'chats': ['chats', 'chat'],
        'contacts': ['contacts', 'contact'],
        'installed_applications': ['installed applications', 'installed app', 'installed apps'],
        'instant_messages': ['instant messages', 'instant message'],
        'user_accounts': ['user accounts', 'user account'],
        'timeline': ['timeline'],
        'applications': ['applications', 'application'],
        'archives': ['archives', 'archive'],
        'databases': ['databases', 'database'],
        'images': ['images', 'image'],
        'text': ['text', 'texts'],
        'all_files': ['all files', 'allfiles']
    }
    stats = {}
    stats_meta = {}
    for label, meta in pairs.items():
        low = label.lower()
        matched = None
        for k, aliases in stats_map.items():
            for a in aliases:
                if a in low:
                    matched = k
                    break
            if matched:
                break
        if matched:
            existing = stats_meta.get(matched)
            if existing:
                if existing.get('section') != 'files' and meta.get('section') == 'files':
                    stats[matched] = meta['val']
                    stats_meta[matched] = meta
                else:
                    pass
            else:
                stats[matched] = meta['val']
                stats_meta[matched] = meta

    if not stats:
        for label, meta in pairs.items():
            val = meta.get('val') if isinstance(meta, dict) else meta
            if re.search(r'\d', val):
                nlabel = re.sub(r'[^a-z0-9]+', '_', label.lower()).strip('_')
                stats[nlabel] = val

    if stats:
        def clean_num(s):
            if s is None:
                return s
            t = re.sub(r'[\s,]+', '', str(s))
            return t
        stats_clean = {k: clean_num(v) for k, v in stats.items()}
        data['indexing_stats'] = stats
        data['indexing_stats_clean'] = stats_clean

    date_re = re.search(r'((?:\d{2,4}[\-/]\d{1,2}[\-/]\d{1,4}))', text)
    if date_re:
        data['detected_date'] = date_re.group(1)

    data['raw_lines'] = lines[:50]
    data['pairs'] = pairs
    return data


if __name__ == '__main__':
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
