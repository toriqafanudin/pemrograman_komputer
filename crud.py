import openpyxl

# Fungsi untuk membuat tabel (Create)
def create_table():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DataMahasiswa"
    ws.append(["ID", "Nama", "Umur"])
    wb.save("data_mahasiswa_openpyxl.xlsx")

# Fungsi untuk menambahkan data (Create)
def insert_data(nama, umur):
    wb = openpyxl.load_workbook("data_mahasiswa_openpyxl.xlsx")
    ws = wb["DataMahasiswa"]
    id_baru = ws.max_row + 1
    ws.append([id_baru, nama, umur])
    wb.save("data_mahasiswa_openpyxl.xlsx")

# Fungsi untuk membaca data (Read)
def read_data():
    wb = openpyxl.load_workbook("data_mahasiswa_openpyxl.xlsx")
    ws = wb["DataMahasiswa"]
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        data.append(row)
    return data

# Fungsi untuk memperbarui data (Update)
def update_data(id, nama, umur):
    wb = openpyxl.load_workbook("data_mahasiswa_openpyxl.xlsx")
    ws = wb["DataMahasiswa"]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if row[0] == id:
            row_index = row[0] + 1
            ws.cell(row=row_index, column=2, value=nama)
            ws.cell(row=row_index, column=3, value=umur)
            wb.save("data_mahasiswa_openpyxl.xlsx")
            break

# Fungsi untuk menghapus data (Delete)
def delete_data(id):
    wb = openpyxl.load_workbook("data_mahasiswa_openpyxl.xlsx")
    ws = wb["DataMahasiswa"]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if row[0] == id:
            row_index = row[0] + 1
            ws.delete_rows(row_index)
            wb.save("data_mahasiswa_openpyxl.xlsx")
            break

# Contoh Penggunaan
create_table()
insert_data("John Doe", 25)
insert_data("Jane Doe", 22)

print("Data awal:")
print(read_data())

update_data(1, "John Smith", 26)

print("\nData setelah pembaruan:")
print(read_data())

delete_data(2)

print("\nData setelah penghapusan:")
print(read_data())
