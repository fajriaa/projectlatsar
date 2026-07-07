from app import strip_indexing_from_text
s='''Barang bukti elektronik ini diterima ...
--- Indeks: contacts: 14,392 (118) (2 cack: 1.81413)
--- Ringkasan Indexing:
- Contacts: 14,392 (118) (2 cack: 1.81413)
1539
Lainnya teks normal.'''
print('----ORIG----')
print(s)
print('----STRIPPED----')
print(strip_indexing_from_text(s))
