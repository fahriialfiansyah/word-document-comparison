from spire.doc import *

# load dokumen versi lama
dokumen_lama = Document()
dokumen_lama.LoadFromFile("versi_lama.docx")

# load dokumen versi terbaru
dokumen_baru = Document()
dokumen_baru.LoadFromFile("versi_baru.docx")

# set opsi perbandingan
opsi = CompareOptions()
opsi.IgnoreFormatting = True  # abaikan perubahan format seperti font, color, dan style

# bandingkan dokumen dengan opsi
dokumen_lama.Compare(dokumen_baru, "Editor", DateTime.get_Now(), opsi)

# simpan hasil perbandingan
dokumen_lama.SaveToFile("hasil_perbandingan_ignore_formatting.docx")