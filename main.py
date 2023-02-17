########################################################################################################################
# Source code: Aplikasi Python untuk Generator File Excel Recap Data SIRUP di Kabupaten Mojokerto                      #
# ---------------------------------------------------------------------------------------------------------------------#
# Aplikasi ini dibuat oleh:                                                                                            #
# Nama          : Muhammad Fajrul Falah, S.ST., M.Tr.T., dkk.                                                          #
# Jabatan       : Tenaga Operator Komputer                                                                             #
# Instansi      : Biro Pengadaan Barang dan Jasa Setda Kab. Mojokerto                                                  #
# Email         : sunrise.pacet@gmail.com                                                                              #
# URL Web       : https://github.com/fajrulfalah93                                                                     #
# ---------------------------------------------------------------------------------------------------------------------#
# Hak cipta milik Allah SWT, source code ini silahkan dicopy, di download atau di distribusikan ke siapa saja untuk    #
# bahan belajar, atau untuk dikembangkan lagi lebih lanjut, btw tidak untuk dijual ya.                                 #
#                                                                                                                      #
# Jika teman-teman mengembangkan lebih lanjut source code ini, agar berkenan untuk men-share code yang teman-teman     #
# kembangkan lebih lanjut sebagai bahan belajar untuk kita semua.                                                      #
# ---------------------------------------------------------------------------------------------------------------------#
# @ Mojokerto, 2023                                                                                                    #
########################################################################################################################

import requests
import datetime as dt
import time
import xlsxwriter

start_time = time.time()

now = dt.datetime.now()
dt_string = now.strftime("%d-%m-%Y-%H-%M")
today = dt.date.today()
yearNow = today.year
maxDate = yearNow + 2
minDate = 2013
mod = 2017

dictOfDate = {
    'January': 'Januari',
    'February': 'Februari',
    'March': 'Maret',
    'April': 'April',
    'May': 'Mei',
    'June': 'Juni',
    'July': 'Juli',
    'August': 'Agustus',
    'September': 'September',
    'October': 'Oktober',
    'November': 'November',
    'December': 'Desember',
}

option = ["default", "pgl", "ep", "slk", "tdr", "tdc", "dk", "pnl", "swa"]
opsi = ["Pilih Metode Pemilihan", "Pengadaan Langsung", "E-Purchasing", "Seleksi", "Tender", "Tender Cepat",
        "Dikecualikan", "Penunjukan Langsung", "Swakelola"]
mtd = "swa"

########################################################################################################################
# Ambil Data Sirup (Data ini berasal dari API JSON yang ditarik dari Information Systems Branch (ISB) LKPP             #
# ---------------------------------------------------------------------------------------------------------------------#
# Mapping Data JSON (Data SIRUP berada pada key 'aaData'                                                               #
########################################################################################################################

recap = 'https://sirup.lkpp.go.id/sirup/datatablectr/datatableruprekapkldi?idKldi=D170&tahun=' + str(
    yearNow)
with requests.get(recap, stream=True) as gr:
    getRecap = gr.json()
    dtRecap = getRecap['aaData']

provider = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediakldi?idKldi=D170&tahun=' + str(
    yearNow)
with requests.get(provider, stream=True) as gp:
    getProv = gp.json()
    dtProv = getProv['aaData']

selfMan = 'https://sirup.lkpp.go.id/sirup/datatablectr/datarupswakelolakldi?idKldi=D170&tahun=' + str(
    yearNow)
with requests.get(selfMan, stream=True) as gs:
    getSelf = gs.json()
    dtSelf = getSelf['aaData']

provSelf = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediaswakelolaallrekapkldi?idKldi=D170&tahun=' + str(
    yearNow)
with requests.get(provSelf, stream=True) as gps:
    getProvSelf = gps.json()
    dtProvSelf = getProvSelf['aaData']

########################################################################################################################
# Alternatif Source untuk Data SIRUP                                                                                   #
########################################################################################################################
# recap = 'https://lpse.or.id/sirup/?mod=datatableruprekapkldi&idKldi=D170&tahun=' + str(yearNow)
# provider = 'https://lpse.or.id/sirup/?mod=dataruppenyediakldi&idKldi=D170&tahun=' + str(yearNow)
# selfMan = 'https://lpse.or.id/sirup/?mod=datarupswakelolakldi&idKldi=D170&tahun=' + str(yearNow)
# provSelf = 'https://lpse.or.id/sirup/?mod=dataruppenyediaswakelolaallrekapkldi&idKldi=D170&tahun=' + str(yearNow)

########################################################################################################################
# Inisialisasi Variable Data Tampung untuk Kalkulasi Data SIRUP                                                        #
########################################################################################################################

countDik = {}
countEpc = {}
countPdl = {}
countPnl = {}
countSlk = {}
countTdr = {}
countTdc = {}
countSwa = {}
countTotal = {}

countDikTotal = {}
countEpcTotal = {}
countPdlTotal = {}
countPnlTotal = {}
countSlkTotal = {}
countTdrTotal = {}
countTdcTotal = {}
countSwaTotal = {}
countTotalTotal = {}

########################################################################################################################
# Menjumlahkan Data Paket dan Pagu pada Data Penyedia berdasarkan Metode Pemilihannya                                  #
########################################################################################################################

for p in dtProv:
    if (countDik.get(p[1], 'None') == 'None' or
            countEpc.get(p[1], 'None') == 'None' or
            countPdl.get(p[1], 'None') == 'None' or
            countPnl.get(p[1], 'None') == 'None' or
            countSlk.get(p[1], 'None') == 'None' or
            countTdr.get(p[1], 'None') == 'None' or
            countTdc.get(p[1], 'None') == 'None'):
        countDik[p[1]] = {'paket': 0, 'pagu': 0}
        countEpc[p[1]] = {'paket': 0, 'pagu': 0}
        countPdl[p[1]] = {'paket': 0, 'pagu': 0}
        countPnl[p[1]] = {'paket': 0, 'pagu': 0}
        countSlk[p[1]] = {'paket': 0, 'pagu': 0}
        countTdr[p[1]] = {'paket': 0, 'pagu': 0}
        countTdc[p[1]] = {'paket': 0, 'pagu': 0}

    if p[4] == 'Dikecualikan':
        countDik[p[1]]['pagu'] += int(p[3])
        countDik[p[1]]['paket'] += 1

    if p[4] == 'e-Purchasing':
        countEpc[p[1]]['pagu'] += int(p[3])
        countEpc[p[1]]['paket'] += 1

    if p[4] == 'Pengadaan Langsung':
        countPdl[p[1]]['pagu'] += int(p[3])
        countPdl[p[1]]['paket'] += 1

    if p[4] == "Penunjukan Langsung":
        countPnl[p[1]]['pagu'] += int(p[3])
        countPnl[p[1]]['paket'] += 1

    if p[4] == "Seleksi":
        countSlk[p[1]]['pagu'] += int(p[3])
        countSlk[p[1]]['paket'] += 1

    if p[4] == "Tender":
        countTdr[p[1]]['pagu'] += int(p[3])
        countTdr[p[1]]['paket'] += 1

    if p[4] == "Tender Cepat":
        countTdc[p[1]]['pagu'] += int(p[3])
        countTdc[p[1]]['paket'] += 1

########################################################################################################################
# Menjumlahkan Data Paket dan Pagu pada Data Penyedia dalam Swakelola berdasarkan Metode Pemilihannya                  #
########################################################################################################################

for p in dtProvSelf:
    if (countDik.get(p[1], 'None') == 'None' or
            countEpc.get(p[1], 'None') == 'None' or
            countPdl.get(p[1], 'None') == 'None' or
            countPnl.get(p[1], 'None') == 'None' or
            countSlk.get(p[1], 'None') == 'None' or
            countTdr.get(p[1], 'None') == 'None' or
            countTdc.get(p[1], 'None') == 'None'):
        countDik[p[1]] = {'paket': 0, 'pagu': 0}
        countEpc[p[1]] = {'paket': 0, 'pagu': 0}
        countPdl[p[1]] = {'paket': 0, 'pagu': 0}
        countPnl[p[1]] = {'paket': 0, 'pagu': 0}
        countSlk[p[1]] = {'paket': 0, 'pagu': 0}
        countTdr[p[1]] = {'paket': 0, 'pagu': 0}
        countTdc[p[1]] = {'paket': 0, 'pagu': 0}

    if p[5] == 'Dikecualikan':
        countDik[p[1]]['pagu'] += int(p[3])
        countDik[p[1]]['paket'] += 1

    if p[5] == 'e-Purchasing':
        countEpc[p[1]]['pagu'] += int(p[3])
        countEpc[p[1]]['paket'] += 1

    if p[5] == 'Pengadaan Langsung':
        countPdl[p[1]]['pagu'] += int(p[3])
        countPdl[p[1]]['paket'] += 1

    if p[5] == "Penunjukan Langsung":
        countPnl[p[1]]['pagu'] += int(p[3])
        countPnl[p[1]]['paket'] += 1

    if p[5] == "Seleksi":
        countSlk[p[1]]['pagu'] += int(p[3])
        countSlk[p[1]]['paket'] += 1

    if p[5] == "Tender":
        countTdr[p[1]]['pagu'] += int(p[3])
        countTdr[p[1]]['paket'] += 1

    if p[5] == "Tender Cepat":
        countTdc[p[1]]['pagu'] += int(p[3])
        countTdc[p[1]]['paket'] += 1

########################################################################################################################
# Menjumlahkan Data Paket dan Pagu pada Data Swakelola                                                                 #
########################################################################################################################

for p in dtSelf:
    if countSwa.get(p[1], 'None') == 'None':
        countSwa[p[1]] = {'paket': 0, 'pagu': 0}

    countSwa[p[1]]['pagu'] += int(p[3])
    countSwa[p[1]]['paket'] += 1

########################################################################################################################
# Membuat List Satuan Kerja                                                                                            #
########################################################################################################################

recapKeysTemp = []

for p in dtRecap:
    recapKeysTemp.append(p[1])

recapKeys = list(dict.fromkeys(recapKeysTemp))

########################################################################################################################
# Menjumlahkan Data Paket dan Pagu pada Masing-Masing Satuan Kerja                                                     #
########################################################################################################################

for key in recapKeys:
    if countTotal.get(key, 'None') == 'None':
        countTotal[key] = {'paket': 0, 'pagu': 0}

    countTotal[key]['paket'] = ((countDik[key]['paket'] if countDik.get(key, 'None') != 'None' else 0) +
                                (countEpc[key]['paket'] if countEpc.get(key, 'None') != 'None' else 0) +
                                (countPdl[key]['paket'] if countPdl.get(key, 'None') != 'None' else 0) +
                                (countPnl[key]['paket'] if countPnl.get(key, 'None') != 'None' else 0) +
                                (countSlk[key]['paket'] if countSlk.get(key, 'None') != 'None' else 0) +
                                (countTdr[key]['paket'] if countTdr.get(key, 'None') != 'None' else 0) +
                                (countTdc[key]['paket'] if countTdc.get(key, 'None') != 'None' else 0) +
                                (countSwa[key]['paket'] if countSwa.get(key, 'None') != 'None' else 0))
    countTotal[key]['pagu'] = ((countDik[key]['pagu'] if countDik.get(key, 'None') != 'None' else 0) +
                               (countEpc[key]['pagu'] if countEpc.get(key, 'None') != 'None' else 0) +
                               (countPdl[key]['pagu'] if countPdl.get(key, 'None') != 'None' else 0) +
                               (countPnl[key]['pagu'] if countPnl.get(key, 'None') != 'None' else 0) +
                               (countSlk[key]['pagu'] if countSlk.get(key, 'None') != 'None' else 0) +
                               (countTdr[key]['pagu'] if countTdr.get(key, 'None') != 'None' else 0) +
                               (countTdc[key]['pagu'] if countTdc.get(key, 'None') != 'None' else 0) +
                               (countSwa[key]['pagu'] if countSwa.get(key, 'None') != 'None' else 0))

########################################################################################################################
# Menambahkan Judul Kolom untuk Excel yang Akan Dibuat                                                                 #
########################################################################################################################

data = [
    [
        "No.",
        "Satuan Kerja",
        "Paket Pengadaan Langsung",
        "Nilai Pengadaan Langsung",
        "Paket E-Purchasing",
        "Nilai E-Purchasing",
        "Paket Seleksi",
        "Nilai Seleksi",
        "Paket Tender",
        "Nilai Tender",
        "Paket Tender Cepat",
        "Nilai Tender Cepat",
        "Paket Dikecualikan",
        "Nilai Dikecualikan",
        "Paket Penunjukan Langsung",
        "Nlai Penunjukan Langsung",
        "Paket Swakelola",
        "Nilai Swakelola",
        "Paket Terumumkan",
        "Pagu Terumumkan"
    ]
]

########################################################################################################################
# Menambahkan Isi Data Berdasarkan Kolom yang Telah Dibuat                                                             #
########################################################################################################################

detailRekap = []

index = 1

for key in recapKeys:
    detailRekap.append([
        index,
        key,
        (countPdl[key]['paket'] if countPdl.get(key, 'None') != 'None' else 0),
        (countPdl[key]['pagu'] if countPdl.get(key, 'None') != 'None' else 0),
        (countEpc[key]['paket'] if countEpc.get(key, 'None') != 'None' else 0),
        (countEpc[key]['pagu'] if countEpc.get(key, 'None') != 'None' else 0),
        (countSlk[key]['paket'] if countSlk.get(key, 'None') != 'None' else 0),
        (countSlk[key]['pagu'] if countSlk.get(key, 'None') != 'None' else 0),
        (countTdr[key]['paket'] if countTdr.get(key, 'None') != 'None' else 0),
        (countTdr[key]['pagu'] if countTdr.get(key, 'None') != 'None' else 0),
        (countTdc[key]['paket'] if countTdc.get(key, 'None') != 'None' else 0),
        (countTdc[key]['pagu'] if countTdc.get(key, 'None') != 'None' else 0),
        (countDik[key]['paket'] if countDik.get(key, 'None') != 'None' else 0),
        (countDik[key]['pagu'] if countDik.get(key, 'None') != 'None' else 0),
        (countPnl[key]['paket'] if countPnl.get(key, 'None') != 'None' else 0),
        (countPnl[key]['pagu'] if countPnl.get(key, 'None') != 'None' else 0),
        (countSwa[key]['paket'] if countSwa.get(key, 'None') != 'None' else 0),
        (countSwa[key]['pagu'] if countSwa.get(key, 'None') != 'None' else 0),
        (countTotal[key]['paket'] if countTotal.get(key, 'None') != 'None' else 0),
        (countTotal[key]['pagu'] if countTotal.get(key, 'None') != 'None' else 0)
    ])
    index += 1

########################################################################################################################
# Menjumlahkan Kolom Data sebagai Data 'TOTAL' Berdasarkan Kolom yang Telah Dibuat                                     #
########################################################################################################################

countPdlTotal = {'paket': 0, 'pagu': 0}
countEpcTotal = {'paket': 0, 'pagu': 0}
countSlkTotal = {'paket': 0, 'pagu': 0}
countTdrTotal = {'paket': 0, 'pagu': 0}
countTdcTotal = {'paket': 0, 'pagu': 0}
countDikTotal = {'paket': 0, 'pagu': 0}
countPnlTotal = {'paket': 0, 'pagu': 0}
countSwakelolaTotal = {'paket': 0, 'pagu': 0}
countTotalTotal = {'paket': 0, 'pagu': 0}

for pin in detailRekap:
    countPdlTotal['paket'] += pin[2]
    countPdlTotal['pagu'] += pin[3]
    countEpcTotal['paket'] += pin[4]
    countEpcTotal['pagu'] += pin[5]
    countSlkTotal['paket'] += pin[6]
    countSlkTotal['pagu'] += pin[7]
    countTdrTotal['paket'] += pin[8]
    countTdrTotal['pagu'] += pin[9]
    countTdcTotal['paket'] += pin[10]
    countTdcTotal['pagu'] += pin[11]
    countDikTotal['paket'] += pin[12]
    countDikTotal['pagu'] += pin[13]
    countPnlTotal['paket'] += pin[14]
    countPnlTotal['pagu'] += pin[15]
    countSwakelolaTotal['paket'] += pin[16]
    countSwakelolaTotal['pagu'] += pin[17]
    countTotalTotal['paket'] += pin[18]
    countTotalTotal['pagu'] += pin[19]

detailRekap.append([
    "",
    "TOTAL",
    countPdlTotal['paket'],
    countPdlTotal['pagu'],
    countEpcTotal['paket'],
    countEpcTotal['pagu'],
    countSlkTotal['paket'],
    countSlkTotal['pagu'],
    countTdrTotal['paket'],
    countTdrTotal['pagu'],
    countTdcTotal['paket'],
    countTdcTotal['pagu'],
    countDikTotal['paket'],
    countDikTotal['pagu'],
    countPnlTotal['paket'],
    countPnlTotal['pagu'],
    countSwakelolaTotal['paket'],
    countSwakelolaTotal['pagu'],
    countTotalTotal['paket'],
    countTotalTotal['pagu']
])

data.extend(detailRekap)

########################################################################################################################
# Men-Generate File Excel Data Rekapitulasi SIRUP Sesuai dengan Waktu Pengambilan (Tanggal dan Jam)                    #
########################################################################################################################

workbook = xlsxwriter.Workbook('rekap-' + dt_string + '.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0
for d1, d2, d3, d4, d5, d6, d7, d8, d9, d10, d11, d12, d13, d14, d15, d16, d17, d18, d19, d20 in data:
    worksheet.write(row, col, d1)
    worksheet.write(row, col + 1, d2)
    worksheet.write(row, col + 2, d3)
    worksheet.write(row, col + 3, d4)
    worksheet.write(row, col + 4, d5)
    worksheet.write(row, col + 5, d6)
    worksheet.write(row, col + 6, d7)
    worksheet.write(row, col + 7, d8)
    worksheet.write(row, col + 8, d9)
    worksheet.write(row, col + 9, d10)
    worksheet.write(row, col + 10, d11)
    worksheet.write(row, col + 11, d12)
    worksheet.write(row, col + 12, d13)
    worksheet.write(row, col + 13, d14)
    worksheet.write(row, col + 14, d15)
    worksheet.write(row, col + 15, d16)
    worksheet.write(row, col + 16, d17)
    worksheet.write(row, col + 17, d18)
    worksheet.write(row, col + 18, d19)
    worksheet.write(row, col + 19, d20)
    row += 1

workbook.close()

tmpTP = []
tmpTPds = []
tmpSwa = []

if mtd == "slk":
    dataPenyedia = []
    dataPds = []

    for pen in dtProv:
        for word, replacement in dictOfDate.items():
            pen[7] = pen[7].replace(word, replacement)
        if pen[4] == "Seleksi":
            dataPenyedia.append([pen[1], pen[2], int(pen[3]), pen[5], pen[6], pen[7]])
    tmpTP = sorted(dataPenyedia)

    for pen in dtProvSelf:
        for word, replacement in dictOfDate.items():
            pen[4] = pen[4].replace(word, replacement)
        if pen[5] == "Seleksi":
            dataPds.append([pen[1], pen[2], int(pen[3]), pen[6], pen[9], pen[4]])
    tmpTPds = sorted(dataPds)

if mtd == "tdr":
    dataPenyedia = []
    dataPds = []

    for pen in dtProv:
        for word, replacement in dictOfDate.items():
            pen[7] = pen[7].replace(word, replacement)
        if pen[4] == "Tender":
            dataPenyedia.append([pen[1], pen[2], int(pen[3]), pen[5], pen[6], pen[7]])
    tmpTP = sorted(dataPenyedia)

    for pen in dtProvSelf:
        for word, replacement in dictOfDate.items():
            pen[4] = pen[4].replace(word, replacement)
        if pen[5] == "Tender":
            dataPds.append([pen[1], pen[2], int(pen[3]), pen[6], pen[9], pen[4]])
    tmpTPds = sorted(dataPds)

if mtd == "tdc":
    dataPenyedia = []
    dataPds = []

    for pen in dtProv:
        for word, replacement in dictOfDate.items():
            pen[7] = pen[7].replace(word, replacement)
        if pen[4] == "Tender Cepat":
            dataPenyedia.append([pen[1], pen[2], int(pen[3]), pen[5], pen[6], pen[7]])
    tmpTP = sorted(dataPenyedia)

    for pen in dtProvSelf:
        for word, replacement in dictOfDate.items():
            pen[4] = pen[4].replace(word, replacement)
        if pen[5] == "Tender Cepat":
            dataPds.append([pen[1], pen[2], int(pen[3]), pen[6], pen[9], pen[4]])
    tmpTPds = sorted(dataPds)

if mtd == "pgl":
    dataPenyedia = []
    dataPds = []

    for pen in dtProv:
        for word, replacement in dictOfDate.items():
            pen[7] = pen[7].replace(word, replacement)
        if pen[4] == "Pengadaan Langsung":
            dataPenyedia.append([pen[1], pen[2], int(pen[3]), pen[5], pen[6], pen[7]])
    tmpTP = sorted(dataPenyedia)

    for pen in dtProvSelf:
        for word, replacement in dictOfDate.items():
            pen[4] = pen[4].replace(word, replacement)
        if pen[5] == "Pengadaan Langsung":
            dataPds.append([pen[1], pen[2], int(pen[3]), pen[6], pen[9], pen[4]])
    tmpTPds = sorted(dataPds)

if mtd == "ep":
    dataPenyedia = []
    dataPds = []

    for pen in dtProv:
        for word, replacement in dictOfDate.items():
            pen[7] = pen[7].replace(word, replacement)
        if pen[4] == "e-Purchasing":
            dataPenyedia.append([pen[1], pen[2], int(pen[3]), pen[5], pen[6], pen[7]])
    tmpTP = sorted(dataPenyedia)

    for pen in dtProvSelf:
        for word, replacement in dictOfDate.items():
            pen[4] = pen[4].replace(word, replacement)
        if pen[5] == "e-Purchasing":
            dataPds.append([pen[1], pen[2], int(pen[3]), pen[6], pen[9], pen[4]])
    tmpTPds = sorted(dataPds)

if mtd == "dk":
    dataPenyedia = []
    dataPds = []

    for pen in dtProv:
        for word, replacement in dictOfDate.items():
            pen[7] = pen[7].replace(word, replacement)
        if pen[4] == "Dikecualikan":
            dataPenyedia.append([pen[1], pen[2], int(pen[3]), pen[5], pen[6], pen[7]])
    tmpTP = sorted(dataPenyedia)

    for pen in dtProvSelf:
        for word, replacement in dictOfDate.items():
            pen[4] = pen[4].replace(word, replacement)
        if pen[5] == "Dikecualikan":
            dataPds.append([pen[1], pen[2], int(pen[3]), pen[6], pen[9], pen[4]])
    tmpTPds = sorted(dataPds)

if mtd == "pnl":
    dataPenyedia = []
    dataPds = []

    for pen in dtProv:
        for word, replacement in dictOfDate.items():
            pen[7] = pen[7].replace(word, replacement)
        if pen[4] == "Penunjukan Langsung":
            dataPenyedia.append([pen[1], pen[2], int(pen[3]), pen[5], pen[6], pen[7]])
    tmpTP = sorted(dataPenyedia)

    for pen in dtProvSelf:
        for word, replacement in dictOfDate.items():
            pen[4] = pen[4].replace(word, replacement)
        if pen[5] == "Penunjukan Langsung":
            dataPds.append([pen[1], pen[2], int(pen[3]), pen[6], pen[9], pen[4]])
    tmpTPds = sorted(dataPds)

if mtd == "swa":
    dataSwa = []

    for pen in dtSelf:
        for word, replacement in dictOfDate.items():
            pen[6] = pen[6].replace(word, replacement)
        if pen[7][(len(pen[7]) - 1)] == ",":
            edit = pen[7][:-1]
        else:
            edit = pen[7]
        dataSwa.append([pen[1], edit, pen[2], int(pen[3]), pen[4], pen[5], pen[6]])
    tmpSwa = sorted(dataSwa)

if (len(tmpTP) == 0) or (len(tmpTPds) == 0):
    tmpJoin = tmpSwa
    header = [
        [
            "No.",
            "Satuan Kerja",
            "Kegiatan",
            "Paket",
            "Pagu",
            "Sumber Dana",
            "Kode RUP",
            "Waktu Pemilihan"
        ]
    ]
    order = 1
    for key in tmpJoin:
        key.insert(0, order)
        order += 1

    gabung = header + tmpJoin

    workbook = xlsxwriter.Workbook('detail-' + mtd + '-' + dt_string + '.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    for d1, d2, d3, d4, d5, d6, d7, d8 in gabung:
        worksheet.write(row, col, d1)
        worksheet.write(row, col + 1, d2)
        worksheet.write(row, col + 2, d3)
        worksheet.write(row, col + 3, d4)
        worksheet.write(row, col + 4, d5)
        worksheet.write(row, col + 5, d6)
        worksheet.write(row, col + 6, d7)
        worksheet.write(row, col + 7, d8)
        row += 1

    workbook.close()
else:
    kolek = tmpTP + tmpTPds
    tmpJoin = sorted(kolek)
    header = [
        [
            "No.",
            "Satuan Kerja",
            "Paket",
            "Pagu",
            "Sumber Dana",
            "Kode RUP",
            "Waktu Pemilihan"
        ]
    ]
    order = 1
    for key in tmpJoin:
        key.insert(0, order)
        order += 1

    gabung = header + tmpJoin

    workbook = xlsxwriter.Workbook('detail-' + mtd + '-' + dt_string + '.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    for d1, d2, d3, d4, d5, d6, d7 in gabung:
        worksheet.write(row, col, d1)
        worksheet.write(row, col + 1, d2)
        worksheet.write(row, col + 2, d3)
        worksheet.write(row, col + 3, d4)
        worksheet.write(row, col + 4, d5)
        worksheet.write(row, col + 5, d6)
        worksheet.write(row, col + 6, d7)
        row += 1

    workbook.close()

print("--- %s detik ---" % (time.time() - start_time))
