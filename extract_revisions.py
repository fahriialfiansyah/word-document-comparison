from spire.doc import *
from spire.doc.common import *

def save_text(fname: str, content: str):
    """simpan teks ke file"""
    with open(fname, "w", encoding="utf-8") as fp:
        fp.write(content)

def format_revision(index: int, revision, content: str) -> str:
    """format informasi revisi"""
    return (
        f"Index: {index}\n"
        f"Type: {revision.Type.name}\n"
        f"Author: {revision.Author}\n"
        f"Content: {content}\n"
    )

input_file = "hasil_perbandingan.docx"     # file input word
output_insert = "revisi_penambahan.txt"    # file output untuk revisi penambahan
output_delete = "revisi_penghapusan.txt"   # file output untuk revisi penghapusan

# load file hasil perbandingan
document = Document()
document.LoadFromFile(input_file)

# list untuk menyimpan hasil revisi
insert_revisions = ["Revisi Penambahan:"]
delete_revisions = ["Revisi Penghapusan:"]

# inisialisasi index revisi
index_insert = 0
index_delete = 0

# loop tiap section dalam dokumen
for k in range(document.Sections.Count):
    sec = document.Sections.get_Item(k)

    # loop tiap objek dalam body section
    for m in range(sec.Body.ChildObjects.Count):
        doc_item = sec.Body.ChildObjects.get_Item(m)

        # jika objek berupa paragraf
        if isinstance(doc_item, Paragraph):
            para = doc_item

            # revisi penambahan di tingkat paragraf
            if para.IsInsertRevision:
                index_insert += 1
                insert_revisions.append(format_revision(index_insert, para.InsertRevision, para.Text))

            # revisi penghapusan di tingkat paragraf
            elif para.IsDeleteRevision:
                index_delete += 1
                delete_revisions.append(format_revision(index_delete, para.DeleteRevision, para.Text))

            # loop objek dalam paragraf
            for j in range(para.ChildObjects.Count):
                obj = para.ChildObjects.get_Item(j)

                # jika objek berupa teks
                if isinstance(obj, TextRange):
                    text_range = obj

                    # revisi penambahan di tingkat teks
                    if text_range.IsInsertRevision:
                        index_insert += 1
                        insert_revisions.append(format_revision(index_insert, text_range.InsertRevision, text_range.Text))

                    # revisi penghapusan di tingkat teks
                    elif text_range.IsDeleteRevision:
                        index_delete += 1
                        delete_revisions.append(format_revision(index_delete, text_range.DeleteRevision, text_range.Text))

# simpan hasil revisi ke file txt
save_text(output_insert, "\n".join(insert_revisions))
save_text(output_delete, "\n".join(delete_revisions))