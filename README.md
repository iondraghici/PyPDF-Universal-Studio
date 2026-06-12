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

---

## Instalare pas cu pas

### 1. Creare folder proiect

```cmd
mkdir C:\Proiecte\PyPDF-Studio
cd C:\Proiecte\PyPDF-Studio
```

### 2. Descărcare cod sursă din GitHub

> ⚠️ **IMPORTANT!** Acest pas este obligatoriu. Fără el, fișierul `pypdf_studio_v2.py` nu va exista.

**Opțiunea A - Git (recomandat):**
```cmd
git clone https://github.com/iondraghici/PyPDF-Universal-Studio.git .
```

**Opțiunea B - Manual (fără Git):**
1. Accesați: https://github.com/iondraghici/PyPDF-Universal-Studio
2. Click pe butonul verde **<> Code** → **Download ZIP**
3. Extrageți conținutul ZIP în `C:\Proiecte\PyPDF-Studio`

**Verificare:** Rulați `dir` și confirmați că vedeți `pypdf_studio_v2.py`.

### 3. Mediu virtual (recomandat)

```cmd
python -m venv venv
venv\Scripts\activate
```

După activare, prompt-ul va arăta astfel:
```
(venv) C:\Proiecte\PyPDF-Studio>
```

### 4. Instalare dependințe

**Toate funcționalitățile (inclusiv conversie și compilare .exe):**
```cmd
pip install -r requirements.txt
```

**Doar esențialele** (fără OCR/conversie/compilare):
```cmd
pip install pypdf PyMuPDF python-docx Pillow
```

### 5. Tesseract OCR (opțional, doar pentru scanări)

1. Descărcați de la: https://github.com/UB-Mannheim/tesseract/wiki
2. Instalați în `C:\Program Files\Tesseract-OCR`
3. Adăugați în PATH:
   - Deschideți "Edit environment variables"
   - La variabila `Path` adăugați: `C:\Program Files\Tesseract-OCR`
4. Verificați în CMD: `tesseract --version`

### 6. Rulare aplicație

```cmd
python pypdf_studio_v2.py
```

> ⚠️ **Atenție:** Rulați această comandă doar din terminal cu venv activat (vezi pasul 3). Dacă deschideți aplicația prin dublu-click pe fișier, Python-ul sistemului nu va găsi bibliotecile instalate în venv.

---

## Ieșire din mediul virtual (venv)

Pentru a dezactiva venv și a reveni la Python-ul sistemului:

```cmd
deactivate
```

Prompt-ul va reveni la forma normală (fără `(venv)` în față).

---

## Ștergere și recreare mediu virtual (venv)

Dacă doriți să ștergeți complet mediul virtual și să reinstalați de la zero:

### 1. Dezactivați venv (dacă este activ)
```cmd
deactivate
```

### 2. Ștergeți folderul venv
```cmd
rmdir /s /q venv
```

### 3. Recreați venv și reinstalați dependințele
```cmd
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

---

## Compilare .EXE (portabil)

```cmd
pip install pyinstaller
pyinstaller --noconsole --onefile --windowed --name "PyPDF_Studio_V2" pypdf_studio_v2.py
```

Executabilul va fi în folderul `dist/`.

> **Notă:** Rulați aceste comenzi cu venv activat. PyInstaller trebuie instalat în același mediu virtual cu celelalte dependințe.

---

## Troubleshooting

### Eroare: `python: can't open file 'pypdf_studio_v2.py': [Errno 2] No such file or directory`

**Cauză:** Nu ați descărcat codul sursă din GitHub (pasul 2).  
**Soluție:** Urmați [pasul 2 de mai sus](#2-descărcare-cod-sursă-din-github).

### Eroare: `Biblioteci esențiale lipsesc: pypdf, PyMuPDF, python-docx, Pillow.`

**Cauză:** Aplicația rulează cu Python-ul sistemului, nu cu cel din venv.  
**Soluție:** Rulați aplicația din terminal cu venv activat: `venv\Scripts\activate` apoi `python pypdf_studio_v2.py`.

### Eroare: `Could not find a version that satisfies the requirement pdf2docx>=1.5.0`

**Cauză:** Fișierul `requirements.txt` conține o versiune inexistentă.  
**Soluție:** Modificați în `requirements.txt` linia `pdf2docx>=1.5.0` în `pdf2docx>=0.5.0`.

### Eroare: `WARNING: Cache entry deserialization failed`

**Cauză:** Cache-ul pip este corupt.  
**Soluție:**
```cmd
pip cache purge
python.exe -m pip install --upgrade pip
```

### Eroare: `'pyinstaller' is not recognized as an internal or external command`

**Cauză:** PyInstaller nu este instalat în venv-ul activ.  
**Soluție:**
```cmd
venv\Scripts\activate
pip install pyinstaller
```

---

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
