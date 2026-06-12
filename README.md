# PyPDF Studio V2 - Procesor Universal de Documente

Aplicație desktop Python pentru manipulare PDF/DOCX cu interfață grafică.

## Funcționalități

| Tab | Descriere |
|-----|-----------|
| **Unire PDF** | Combină multiple fișiere PDF într-unul singur |
| **Divizare PDF** | Extrage pagini după intervale, pagină cu pagină, sau la fiecare N pagini |
| **Completare & Semnătură** | Completează formulare PDF/DOCX manual sau automat. Semnătură prin desenat sau imagine. Plasare semnătură prin **click pe previzualizare**. |
| **Securitate & Metadate** | Protejare cu parolă, editare metadate |
| **Conversie** | PDF↔DOCX, OCR pentru scanări |

## Cerințe sistem

- Windows 10/11
- Python 3.10+
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) (doar pentru funcția OCR)

## Instalare pas cu pas

### 1. Creare folder proiect

```cmd
mkdir C:\Proiecte\PyPDF-Studio
cd C:\Proiecte\PyPDF-Studio
```

### 2. Mediu virtual (recomandat)

```cmd
python -m venv venv
venv\Scripts\activate
```

### 3. Instalare dependințe

```cmd
pip install -r requirements.txt
```

**Doar esențialele** (fără OCR/conversie):
```cmd
pip install pypdf PyMuPDF python-docx Pillow
```

### 4. Tesseract OCR (opțional, doar pentru scanări)

1. Descarcă de la: https://github.com/UB-Mannheim/tesseract/wiki
2. Instalează în `C:\Program Files\Tesseract-OCR`
3. Adaugă în PATH:
   - Deschide "Edit environment variables"
   - La `Path` adaugă: `C:\Program Files\Tesseract-OCR`
4. Verifică în CMD: `tesseract --version`

### 5. Rulează aplicația

```cmd
python pypdf_studio_v2.py
```

## Compilare .EXE (portabil)

```cmd
pip install pyinstaller
pyinstaller --noconsole --onefile --windowed --name "PyPDF_Studio_V2" pypdf_studio_v2.py
```

Executabilul va fi în folderul `dist/`.

## Cum completezi un document

### Mod Manual (recomandat pentru formulare administrative)

1. Deschide tab-ul **Completare & Semnătură**
2. Încarcă PDF-ul sau DOCX-ul
3. Apasă **"Adaugă câmp"** și scrie eticheta (ex: `Nume`, `CNP`, `Nr. înmatriculare`)
4. Completează valoarea în dreapta etichetei
5. (Opțional) Desenează semnătura sau încarc-o din imagine
6. (Opțional) Dă **click pe previzualizare** unde vrei să apară semnătura
7. Apasă **"Salvează document completat"**

### Mod Automat (doar pentru PDF-uri digitale clare)

1. Selectează modul **"Detectare automată"**
2. Apasă **"Scanează"** – aplicația caută cuvinte-cheie (CNP, Nume, Adresă etc.)
3. Completează valorile detectate
4. Salvează

### Mod OCR (pentru scanări)

1. Selectează modul **"OCR"**
2. Apasă **"Scanează"** – citește textul din imagine folosind Tesseract
3. Completează valorile
4. Salvează

## Notă importantă

Pentru documente administrative românești (precum declarații de la primărie, DGITL, etc.), **modul Manual este cel mai fiabil** deoarece aceste documente nu au câmpuri de formular standardizate.

## Publicare GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/iondraghici/PyPDF-Universal-Studio.git
git push -u origin main
```

## Licență

Proiect personal. Folosiți pe propria răspundere.
