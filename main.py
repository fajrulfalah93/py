import requests as rq
import urllib3
import ssl
import datetime as dt
import json
import time
import xlsxwriter

start_time = time.time()

# class CustomHttpAdapter(requests.adapters.HTTPAdapter):
#     # "Transport adapter" that allows us to use custom ssl_context.
#     def __init__(self, ssl_context=None, **kwargs):
#         self.poolmanager = None
#         self.ssl_context = ssl_context
#         super().__init__(**kwargs)
#
#     def init_poolmanager(self, connections, maxsize, block=False):
#         self.poolmanager = urllib3.poolmanager.PoolManager(
#             num_pools=connections, maxsize=maxsize,
#             block=block, ssl_context=self.ssl_context)
#
#
# def get_legacy_session():
#     ctx = ssl.create_default_context(ssl.Purpose.SERVER_AUTH)
#     ctx.options |= 0x4  # OP_LEGACY_SERVER_CONNECT
#     session = requests.session()
#     session.mount('https://', CustomHttpAdapter(ctx))
#     return session

now = dt.datetime.now()
dt_string = now.strftime("%d-%m-%Y-%H-%M")
today = dt.date.today()
yearNow = today.year
maxDate = yearNow + 2
minDate = 2013
mod = 2017

mapObj = {
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

# for x in range(minDate,maxDate):
#     recap = 'https://sirup.lkpp.go.id/sirup/datatablectr/datatableruprekapkldi?idKldi=D170&tahun=' + str(x)
#     getRecap = get_legacy_session().get(recap).json()
#     with open("recap" + str(x) + ".json", "w") as outfile:
#         json.dump(getRecap, outfile)
#
# for x in range(minDate,maxDate):
#     provider = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediakldi?idKldi=D170&tahun=' + str(x)
#     getProv = get_legacy_session().get(provider).json()
#     with open("provider" + str(x) + ".json", "w") as outfile:
#         json.dump(getProv, outfile)
#
# for x in range(minDate,maxDate):
#     selfMan = 'https://sirup.lkpp.go.id/sirup/datatablectr/datarupswakelolakldi?idKldi=D170&tahun=' + str(x)
#     getSelf = get_legacy_session().get(selfMan).json()
#     with open("selfMan" + str(x) + ".json", "w") as outfile:
#         json.dump(getSelf, outfile)
#
# for x in range(mod,maxDate):
#     provSelf = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediaswakelolaallrekapkldi?idKldi=D170&tahun=' + str(x)
#     getProvSelf = get_legacy_session().get(provSelf).json()
#     with open("provSelf" + str(x) + ".json", "w") as outfile:
#         json.dump(getProvSelf, outfile)

recap = 'https://sirup.lkpp.go.id/sirup/datatablectr/datatableruprekapkldi?idKldi=D170&tahun=' + str(
    yearNow)
# getRecap = get_legacy_session().get(recap).json()
getRecap = rq.get(recap).json()
# with open("D:/learn/pbj/data/recap" + str(yearNow) + ".json", "w") as outfile:
#     json.dump(getRecap, outfile)

provider = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediakldi?idKldi=D170&tahun=' + str(
    yearNow)
# getProv = get_legacy_session().get(provider).json()
getProv = rq.get(provider).json()
# with open("D:/learn/pbj/data/provider" + str(yearNow) + ".json", "w") as outfile:
#     json.dump(getProv, outfile)

selfMan = 'https://sirup.lkpp.go.id/sirup/datatablectr/datarupswakelolakldi?idKldi=D170&tahun=' + str(
    yearNow)
# getSelf = get_legacy_session().get(selfMan).json()
getSelf = rq.get(selfMan).json()
# with open("D:/learn/pbj/data/selfMan" + str(yearNow) + ".json", "w") as outfile:
#     json.dump(getSelf, outfile)

provSelf = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediaswakelolaallrekapkldi?idKldi=D170&tahun=' + str(
    yearNow)
# getProvSelf = get_legacy_session().get(provSelf).json()
getProvSelf = rq.get(provSelf).json()
# with open("D:/learn/pbj/data/provSelf" + str(yearNow) + ".json", "w") as outfile:
#     json.dump(getProvSelf, outfile)

# recap = 'https://sirup.lkpp.go.id/sirup/datatablectr/datatableruprekapkldi?idKldi=D170&tahun=' + str(yearNow)
# provider = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediakldi?idKldi=D170&tahun=' + str(yearNow)
# selfMan = 'https://sirup.lkpp.go.id/sirup/datatablectr/datarupswakelolakldi?idKldi=D170&tahun=' + str(yearNow)
# provSelf = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediaswakelolaallrekapkldi?idKldi=D170&tahun=' + str(
#     yearNow)

# # recap = 'https://lpse.or.id/sirup/?mod=datatableruprekapkldi&idKldi=D170&tahun=' + str(yearNow)
# # provider = 'https://lpse.or.id/sirup/?mod=dataruppenyediakldi&idKldi=D170&tahun=' + str(yearNow)
# # selfMan = 'https://lpse.or.id/sirup/?mod=datarupswakelolakldi&idKldi=D170&tahun=' + str(yearNow)
# # provSelf = 'https://lpse.or.id/sirup/?mod=dataruppenyediaswakelolaallrekapkldi&idKldi=D170&tahun=' + str(
# #     yearNow)

dtRecap = getRecap['aaData']
dtProv = getProv['aaData']
dtSelf = getSelf['aaData']
dtProvSelf = getProvSelf['aaData']

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

for p in dtSelf:
    if countSwa.get(p[1], 'None') == 'None':
        countSwa[p[1]] = {'paket': 0, 'pagu': 0}

    countSwa[p[1]]['pagu'] += int(p[3])
    countSwa[p[1]]['paket'] += 1

recapKeysTemp = []

for p in dtRecap:
    recapKeysTemp.append(p[1])

recapKeys = list(dict.fromkeys(recapKeysTemp))

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

workbook = xlsxwriter.Workbook('rekap-' + dt_string + '.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0
for d1, d2, d3, d4, d5, d6, d7, d8, d9, d10, d11, d12, d13, d14, d15, d16, d17, d18, d19, d20 in (data):
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

# # makeAPICall2 = async () = > {
# # try {
# # if (mtd == = "default") {
# # return;
# # }
# #
# # let
# # response = await axios.get(
# # "https://lpse.or.id/sirup/?mod=dataruppenyediakldi&idKldi=D170&tahun=" +
# # year
# # );
# # const
# # penyedia2 = response.data.aaData;
# # response = await axios.get(
# # "https://lpse.or.id/sirup/?mod=dataruppenyediaswakelolaallrekapkldi&idKldi=D170&tahun=" +
# # year
# # );
# # const
# # pdsw2 = response.data.aaData;
# #
# # response = await axios.get(
# # "https://lpse.or.id/sirup/?mod=datarupswakelolakldi&idKldi=D170&tahun=" +
# # year
# # );
# # const
# # swa2 = response.data.aaData;
# #
# # let
# # tmpTP = [];
# # let
# # tmpTPds = [];
# # let
# # tmpSwa = [];
# #
# # if (mtd === "slk")
# # {
# # let
# # tenderPenyedia = [];
# # let
# # tenderPds = [];
# #
# # penyedia2.forEach((pen) = > {
# #     let
# # edit = pen[7].replace(
# #        / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function(matched)
# # {
# # return mapObj[matched];
# # }
# # );
# # if (pen[4] === "Seleksi") {
# # tenderPenyedia.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[5],
# # pen[6],
# # edit,
# # ]);
# # }
# # });
# # tmpTP = tenderPenyedia.sort();
# #
# # pdsw2.forEach((pen) = > {
# # let edit = pen[4].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[5] === "Seleksi") {
# # tenderPds.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[6],
# # pen[9],
# # edit,
# # ]);
# # }
# # });
# # tmpTPds = tenderPds.sort();
# # }
# #
# # if (mtd == = "tdr") {
# # let tenderPenyedia =[];
# # let tenderPds =[];
# #
# # penyedia2.forEach((pen) = > {
# # let edit = pen[7].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[4] === "Tender") {
# # tenderPenyedia.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[5],
# # pen[6],
# # edit,
# # ]);
# # }
# # });
# # tmpTP = tenderPenyedia.sort();
# #
# # pdsw2.forEach((pen) = > {
# # let edit = pen[4].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[5] === "Tender") {
# # tenderPds.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[6],
# # pen[9],
# # edit,
# # ]);
# # }
# # });
# # tmpTPds = tenderPds.sort();
# # }
# #
# # if (mtd == = "tdc") {
# # let tenderPenyedia =[];
# # let tenderPds =[];
# #
# # penyedia2.forEach((pen) = > {
# # let edit = pen[7].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[4] === "Tender Cepat") {
# # tenderPenyedia.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[5],
# # pen[6],
# # edit,
# # ]);
# # }
# # });
# # tmpTP = tenderPenyedia.sort();
# #
# # pdsw2.forEach((pen) = > {
# # let edit = pen[4].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[5] === "Tender Cepat") {
# # tenderPds.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[6],
# # pen[9],
# # edit,
# # ]);
# # }
# # });
# # tmpTPds = tenderPds.sort();
# # }
# #
# # if (mtd == = "pgl") {
# # let pglPenyedia =[];
# # let pglPds =[];
# #
# # penyedia2.forEach((pen) = > {
# # let edit = pen[7].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[4] === "Pengadaan Langsung") {
# # pglPenyedia.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[5],
# # pen[6],
# # edit,
# # ]);
# # }
# # });
# # tmpTP = pglPenyedia.sort();
# # pdsw2.forEach((pen) = > {
# # let edit = pen[4].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[5] === "Pengadaan Langsung") {
# # pglPds.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[6],
# # pen[9],
# # edit,
# # ]);
# # }
# # });
# # tmpTPds = pglPds.sort();
# # }
# #
# # if (mtd == = "ep") {
# # let epPenyedia =[];
# # let epPds =[];
# #
# # penyedia2.forEach((pen) = > {
# # let edit = pen[7].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[4] === "e-Purchasing") {
# # epPenyedia.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[5],
# # pen[6],
# # edit,
# # ]);
# # }
# # });
# # tmpTP = epPenyedia.sort();
# # pdsw2.forEach((pen) = > {
# # let edit = pen[4].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[5] === "e-Purchasing") {
# # epPds.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[6],
# # pen[9],
# # edit,
# # ]);
# # }
# # });
# # tmpTPds = epPds.sort();
# # }
# #
# # if (mtd == = "dk") {
# # let dkPenyedia =[];
# # let dkPds =[];
# #
# # penyedia2.forEach((pen) = > {
# # let edit = pen[7].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[4] === "Dikecualikan") {
# # dkPenyedia.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[5],
# # pen[6],
# # edit,
# # ]);
# # }
# # });
# # tmpTP = dkPenyedia.sort();
# # pdsw2.forEach((pen) = > {
# # let edit = pen[4].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[5] === "Dikecualikan") {
# # dkPds.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[6],
# # pen[9],
# # edit,
# # ]);
# # }
# # });
# # tmpTPds = dkPds.sort();
# # }
# #
# # if (mtd == = "pnl") {
# # let pnlPenyedia =[];
# # let pnlPds =[];
# #
# # penyedia2.forEach((pen) = > {
# # let edit = pen[7].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[4] === "Penunjukan Langsung") {
# # pnlPenyedia.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[5],
# # pen[6],
# # edit,
# # ]);
# # }
# # });
# # tmpTP = pnlPenyedia.sort();
# # pdsw2.forEach((pen) = > {
# # let edit = pen[4].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[5] === "Penunjukan Langsung") {
# # pnlPds.push([
# # pen[1],
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[6],
# # pen[9],
# # edit,
# # ]);
# # }
# # });
# # tmpTPds = pnlPds.sort();
# # }
# #
# # if (mtd == = "swa") {
# # let swaDetail =[];
# # let edit2 = "";
# #
# # swa2.forEach((pen) = > {
# # let edit = pen[6].replace(
# # / January | February | March | April | May | June | July | August | September | October | November | December /,
# # function (matched) {
# # return mapObj[matched];
# # }
# # );
# # if (pen[7].charAt(pen[7].length - 1) === ",") {
# # edit2 = pen[7].slice(0, -1);
# # } else {
# # edit2 = pen[7];
# # }
# # swaDetail.push([
# # pen[1],
# # edit2,
# # pen[2],
# # parseInt(pen[3], 0),
# # pen[4],
# # pen[5],
# # edit,
# # ]);
# # });
# # tmpSwa = swaDetail.sort();
# # }
# #
# # let
# # tmpJoin = [];
# # let
# # header = [];
# #
# # if (tmpTP.length === 0 | | tmpTPds == = 0)
# # {
# #     tmpJoin = tmpSwa.sort();
# # header = [
# #     [
# #         "No.",
# #         "Satuan Kerja",
# #         "Kegiatan",
# #         "Paket",
# #         "Pagu",
# #         "Sumber Dana",
# #         "Kode RUP",
# #         "Waktu Pemilihan",
# #     ],
# # ];
# # } else {
# #     let
# # kolek = tmpTP.concat(tmpTPds);
# # tmpJoin = kolek.sort();
# # header = [
# #     [
# #         "No.",
# #         "Satuan Kerja",
# #         "Paket",
# #         "Pagu",
# #         "Sumber Dana",
# #         "Kode RUP",
# #         "Waktu Pemilihan",
# #     ],
# # ];
# # }
# #
# # let
# # i = 1;
# # tmpJoin.forEach(myFunction);
# # function
# # myFunction(item)
# # {
# #     item.unshift(i);
# # i + +;
# # }
# #
# # let
# # gabung = header.concat(tmpJoin);
# #
# # var
# # wb = utils.book_new(), \
# # ws = utils.aoa_to_sheet(gabung);
# # utils.book_append_sheet(wb, ws, "MySheet1");
# # writeFileXLSX(wb, "Detail.xlsx");
# # } catch(e)
# # {
# #     console.log(e);
# # }
# # };

print("--- %s detik ---" % (time.time() - start_time))
