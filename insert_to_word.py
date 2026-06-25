#!/usr/bin/env python3
"""
Script untuk menggabungkan gambar/PDF dari folder ke dalam dokumen Word.
- Auto-deteksi tipe file (image/pdf)
- Ukuran gambar: 70% dari aslinya
- Header 1: "Lampiran"
"""
import os
import sys
import glob
from pathlib import Path
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import tempfile

try:
    from pdf2image import convert_from_path
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False
    print("Warning: pdf2image tidak terinstall. PDF tidak akan diproses.")
    print("Install dengan: pip install pdf2image")


def create_lampiran_header(doc, num, name):
    """Buat header level 1 untuk lampiran"""
    header = doc.add_heading(f"Lampiran {num}: {name}", level=1)
    header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return header


def embed_image(doc, image_path, size_percent=0.7):
    """Embed image dengan ukuran 70% dari aslinya"""
    try:
        img = Image.open(image_path)

        # Hitung ukuran baru (70%)
        width, height = img.size
        new_width = int(width * size_percent)
        new_height = int(height * size_percent)

        # Resize gambar
        img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

        # Simpan ke temporary file
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            img_resized.save(tmp.name, 'PNG')
            temp_path = tmp.name

        # Insert ke docx
        doc.add_paragraph()
        # Konversi pixel ke inches (96 DPI standard)
        width_inches = new_width / 96.0
        doc.add_picture(temp_path, width=Inches(min(width_inches, 6.5)))
        doc.add_paragraph()

        # Hapus temporary file
        os.remove(temp_path)
        return True

    except Exception as e:
        print(f"Error embedding image {image_path}: {e}")
        return False


def embed_pdf(doc, pdf_path, size_percent=0.7):
    """Embed PDF sebagai screenshot halaman pertama"""
    if not PDF_SUPPORT:
        print(f"Skip PDF (tidak ada pdf2image): {pdf_path}")
        return False

    try:
        # Convert PDF ke images (ambil halaman pertama)
        pages = convert_from_path(pdf_path, dpi=150, first_page=1, last_page=1)

        if not pages:
            return False

        page = pages[0]
        width, height = page.size
        new_width = int(width * size_percent)
        new_height = int(height * size_percent)

        # Resize
        img_resized = page.resize((new_width, new_height), Image.Resampling.LANCZOS)

        # Simpan ke temporary file
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            img_resized.save(tmp.name, 'PNG')
            temp_path = tmp.name

        # Insert ke docx
        doc.add_paragraph()
        width_inches = new_width / 96.0
        doc.add_picture(temp_path, width=Inches(min(width_inches, 6.5)))
        doc.add_paragraph()

        # Hapus temporary file
        os.remove(temp_path)
        return True

    except Exception as e:
        print(f"Error embedding PDF {pdf_path}: {e}")
        return False


def process_folder(folder_path, output_file='dokumentasi_penelitian.docx'):
    """Proses folder dan generate dokumen Word"""

    if not os.path.exists(folder_path):
        print(f"Error: Folder tidak ditemukan: {folder_path}")
        return

    # Buat dokumen baru
    doc = Document()
    doc.add_heading('Dokumentasi Penelitian', level=0)
    doc.add_paragraph()

    # Temukan semua file
    image_extensions = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'JPG', 'JPEG', 'PNG', 'GIF', 'BMP', 'TIFF']
    pdf_extensions = ['pdf', 'PDF']

    all_files = []

    # Scan untuk images
    for ext in image_extensions:
        pattern = os.path.join(folder_path, f'*.{ext}')
        all_files.extend(glob.glob(pattern))

    # Scan untuk PDFs
    for ext in pdf_extensions:
        pattern = os.path.join(folder_path, f'*.{ext}')
        all_files.extend(glob.glob(pattern))

    # Sort berdasarkan nama file
    all_files = sorted(set(all_files))

    if not all_files:
        print(f"Tidak ada file gambar atau PDF ditemukan di: {folder_path}")
        return

    print(f"Ditemukan {len(all_files)} file untuk diproses...")

    lampiran_num = 1
    success_count = 0

    for file_path in all_files:
        filename = os.path.basename(file_path)
        extension = file_path.lower().split('.')[-1]

        print(f"Memproses [{lampiran_num}]: {filename}")

        # Buat header lampiran
        create_lampiran_header(doc, lampiran_num, filename)

        # Embed berdasarkan tipe file
        if extension in ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff']:
            success = embed_image(doc, file_path)
        elif extension == 'pdf':
            success = embed_pdf(doc, file_path)
        else:
            success = False

        if success:
            success_count += 1

        lampiran_num += 1

    # Simpan dokumen
    doc.save(output_file)
    print(f"\n{'='*50}")
    print(f"Dokumen Word berhasil dibuat: {output_file}")
    print(f"Total file diproses: {success_count}/{len(all_files)}")
    print(f"{'='*50}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python insert_to_word.py <folder_path> [output_file.docx]")
        print("\nContoh:")
        print("  python insert_to_word.py '/path/to/dokumentasi penelitian'")
        print("  python insert_to_word.py '/path/to/dokumentasi penelitian' output.docx")
        sys.exit(1)

    folder_path = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'dokumentasi_penelitian.docx'

    process_folder(folder_path, output_file)
