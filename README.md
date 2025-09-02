# Word Document Comparison

<img width="1682" height="567" alt="image" src="https://github.com/user-attachments/assets/8c31e3cc-7844-4c93-b74b-98628aa8c604" />

## Project Structure
```
word-document-comparison/
│── compare_basic.py              # membandingkan 2 dokumen Word (versi lama & baru) secara default
│── compare_with_options.py       # membandingkan dokumen dengan opsi tambahan (misalnya ignore formatting)
│── extract_revisions.py          # mengekstrak hasil revisi (insert & delete) ke file txt
│── versi_lama.docx               # contoh dokumen lama
│── versi_baru.docx               # contoh dokumen baru
│── requirements.txt              # dependensi python
```

## Installation Steps
1. Clone the repository:
```shell
git clone https://github.com/fahriialfiansyah/word-document-comparison.git
```
2. Navigate to the project directory:
```shell
cd word-document-comparison
```
3. Install the required dependencies:
```shell
pip install -r requirements.txt
```
4. Place your Word documents (versi_lama.docx and versi_baru.docx) in the project root directory.
5. Run a comparison script:
    - Basic comparison:
      ```shell
      python compare_basic.py
      ```
    - Comparison with options (ignore formatting, etc.):
      ```shell
      python compare_with_options.py
      ```
    - Extract insertions and deletions into text files:
      ```shell
      python extract_revisions.py
      ```
6. The results will be saved as:
   - ```shell hasil_perbandingan.docx``` or ```shell hasil_perbandingan_ignore_formatting.docx``` (Word comparison output)
   - ```shell revisi_penambahan.txt``` and ```shell revisi_penghapusan.txt``` (revision details)
