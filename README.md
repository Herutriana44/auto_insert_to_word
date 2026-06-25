# Auto Insert to Word

Script Python untuk memasukkan gambar dan PDF dari folder ke dalam dokumen Word dengan otomatis.

## Fitur
- ✅ Auto-deteksi file gambar (JPG, PNG, GIF, BMP, TIFF)
- ✅ Auto-deteksi file PDF
- ✅ Resize otomatis ke 70% dari ukuran asli
- ✅ Header "Lampiran" dengan numbering otomatis
- ✅ Sort file berdasarkan nama

## Instalasi

```bash
pip install -r requirements.txt
```

**Catatan untuk PDF:**
- Jika ingin memproses PDF, pastikan `poppler-utils` terinstall
- Di Ubuntu/Debian: `sudo apt-get install poppler-utils`
- Di Termux: `pkg install poppler`

## Cara Pakai

```bash
python insert_to_word.py <folder_path> [output_file.docx]
```

### Contoh:

```bash
# Basic usage
python insert_to_word.py "/path/to/dokumentasi penelitian"

# Custom output file
python insert_to_word.py "/path/to/dokumentasi penelitian" hasil.docx
```

## Output

Script akan membuat file `.docx` dengan struktur:
```
Dokumentasi Penelitian (Heading 0)

Lampiran 1: gambar1.jpg (Heading 1)
[gambar dengan lebar 70%]

Lampiran 2: dokumen.pdf (Heading 1)
[preview halaman pertama PDF dengan lebar 70%]

...
```

## Troubleshooting

**Error: pdf2image tidak terinstall**
- Install: `pip install pdf2image`
- Install poppler: `pkg install poppler` (Termux) atau `apt install poppler-utils` (Linux)

**File tidak terdeteksi**
- Pastikan ekstensi file sesuai: .jpg, .png, .pdf, dll
- Cek path folder sudah benar
