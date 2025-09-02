from spire.doc import *

# load dokumen versi lama
dokumen_lama = Document()
dokumen_lama.LoadFromFile("versi_lama.docx")

# load dokumen versi terbaru
dokumen_baru = Document()
dokumen_baru.LoadFromFile("versi_baru.docx")

# bandingkan dokumen
dokumen_lama.Compare(dokumen_baru, "Editor")

# simpan hasil perbandingan
dokumen_lama.SaveToFile("hasil_perbandingan.docx")