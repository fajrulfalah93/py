import requests as rq
import datetime as dt
# import json

today = dt.date.today()
# yearNow = today.year
yearNow = 2022
maxDate = yearNow + 2

recap = 'https://sirup.lkpp.go.id/sirup/datatablectr/datatableruprekapkldi?idKldi=D170&tahun=' + str(yearNow)
provider = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediakldi?idKldi=D170&tahun=' + str(yearNow)
selfMan = 'https://sirup.lkpp.go.id/sirup/datatablectr/datarupswakelolakldi?idKldi=D170&tahun=' + str(yearNow)
provSelf = 'https://sirup.lkpp.go.id/sirup/datatablectr/dataruppenyediaswakelolaallrekapkldi?idKldi=D170&tahun=' + str(
    yearNow)

# recap = 'https://lpse.or.id/sirup/?mod=datatableruprekapkldi&idKldi=D170&tahun=' + str(yearNow)
# provider = 'https://lpse.or.id/sirup/?mod=dataruppenyediakldi&idKldi=D170&tahun=' + str(yearNow)
# selfMan = 'https://lpse.or.id/sirup/?mod=datarupswakelolakldi&idKldi=D170&tahun=' + str(yearNow)
# provSelf = 'https://lpse.or.id/sirup/?mod=dataruppenyediaswakelolaallrekapkldi&idKldi=D170&tahun=' + str(
#     yearNow)

getRecap = rq.get(recap).json()
getProv = rq.get(provider).json()
getSelf = rq.get(selfMan).json()
getProvSelf = rq.get(provSelf).json()

# with open("recap.json", "w") as outfile:
#     json.dump(getRecap, outfile)
#
# with open("provider.json", "w") as outfile:
#     json.dump(getProv, outfile)
#
# with open("selfMan.json", "w") as outfile:
#     json.dump(getSelf, outfile)
#
# with open("provSelf.json", "w") as outfile:
#     json.dump(getProSelf, outfile)

dtRecap = getRecap['aaData']
dtProv = getProv['aaData']
dtSelf = getSelf['aaData']
dtProvSelf = getProvSelf['aaData']

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

# for p in dtProv:
#     if (countDik.get(p[1], 'None') == 'None' or
#             countEpc.get(p[1], 'None') == 'None' or
#             countPdl.get(p[1], 'None') == 'None' or
#             countPnl.get(p[1], 'None') == 'None' or
#             countSlk.get(p[1], 'None') == 'None' or
#             countTdr.get(p[1], 'None') == 'None' or
#             countTdc.get(p[1], 'None') == 'None'):
#         countDik[p[1]] = {'paket': 0, 'pagu': 0}
#         countEpc[p[1]] = {'paket': 0, 'pagu': 0}
#         countPdl[p[1]] = {'paket': 0, 'pagu': 0}
#         countPnl[p[1]] = {'paket': 0, 'pagu': 0}
#         countSlk[p[1]] = {'paket': 0, 'pagu': 0}
#         countTdr[p[1]] = {'paket': 0, 'pagu': 0}
#         countTdc[p[1]] = {'paket': 0, 'pagu': 0}
#
#     if p[4] == 'Dikecualikan':
#         countDik[p[1]]['pagu'] += int(p[3])
#         countDik[p[1]]['paket'] += 1
#
#     if p[4] == 'e-Purchasing':
#         countEpc[p[1]]['pagu'] += int(p[3])
#         countEpc[p[1]]['paket'] += 1
#
#     if p[4] == 'Pengadaan Langsung':
#         countPdl[p[1]]['pagu'] += int(p[3])
#         countPdl[p[1]]['paket'] += 1
#
#     if p[4] == "Penunjukan Langsung":
#         countPnl[p[1]]['pagu'] += int(p[3])
#         countPnl[p[1]]['paket'] += 1
#
#     if p[4] == "Seleksi":
#         countSlk[p[1]]['pagu'] += int(p[3])
#         countSlk[p[1]]['paket'] += 1
#
#     if p[4] == "Tender":
#         countTdr[p[1]]['pagu'] += int(p[3])
#         countTdr[p[1]]['paket'] += 1
#
#     if p[4] == "Tender Cepat":
#         countTdc[p[1]]['pagu'] += int(p[3])
#         countTdc[p[1]]['paket'] += 1
#
# print(countDik)
# print(countEpc)
# print(countPdl)
# print(countPnl)
# print(countSlk)
# print(countTdr)
# print(countTdc)

# for p in dtProvSelf:
#     if (countDik.get(p[1], 'None') == 'None' or
#             countEpc.get(p[1], 'None') == 'None' or
#             countPdl.get(p[1], 'None') == 'None' or
#             countPnl.get(p[1], 'None') == 'None' or
#             countSlk.get(p[1], 'None') == 'None' or
#             countTdr.get(p[1], 'None') == 'None' or
#             countTdc.get(p[1], 'None') == 'None'):
#         countDik[p[1]] = {'paket': 0, 'pagu': 0}
#         countEpc[p[1]] = {'paket': 0, 'pagu': 0}
#         countPdl[p[1]] = {'paket': 0, 'pagu': 0}
#         countPnl[p[1]] = {'paket': 0, 'pagu': 0}
#         countSlk[p[1]] = {'paket': 0, 'pagu': 0}
#         countTdr[p[1]] = {'paket': 0, 'pagu': 0}
#         countTdc[p[1]] = {'paket': 0, 'pagu': 0}
#
#     if p[4] == 'Dikecualikan':
#         countDik[p[1]]['pagu'] += int(p[3])
#         countDik[p[1]]['paket'] += 1
#
#     if p[4] == 'e-Purchasing':
#         countEpc[p[1]]['pagu'] += int(p[3])
#         countEpc[p[1]]['paket'] += 1
#
#     if p[4] == 'Pengadaan Langsung':
#         countPdl[p[1]]['pagu'] += int(p[3])
#         countPdl[p[1]]['paket'] += 1
#
#     if p[4] == "Penunjukan Langsung":
#         countPnl[p[1]]['pagu'] += int(p[3])
#         countPnl[p[1]]['paket'] += 1
#
#     if p[4] == "Seleksi":
#         countSlk[p[1]]['pagu'] += int(p[3])
#         countSlk[p[1]]['paket'] += 1
#
#     if p[4] == "Tender":
#         countTdr[p[1]]['pagu'] += int(p[3])
#         countTdr[p[1]]['paket'] += 1
#
#     if p[4] == "Tender Cepat":
#         countTdc[p[1]]['pagu'] += int(p[3])
#         countTdc[p[1]]['paket'] += 1
#
# print(countDik)
# print(countEpc)
# print(countPdl)
# print(countPnl)
# print(countSlk)
# print(countTdr)
# print(countTdc)
#
for p in dtSelf:
    if countSwa.get(p[1], 'None') == 'None':
        countSwa[p[1]] = {'paket': 0, 'pagu': 0}

    countSwa[p[1]]['pagu'] += int(p[3])
    countSwa[p[1]]['paket'] += 1

# print(countSwa)
#
# recapKeys = []
#
# for p in dtRecap:
#     recapKeys.append(p[1])
#
# for key in recapKeys:
#     countTotal[key] = {
#         'paket':
#             (countDik[key]['paket'] if countDik[key] else 0) +
#             (countEpc[key]['paket'] if countEpc[key] else 0) +
#             (countPdl[key]['paket'] if countPdl[key] else 0) +
#             (countPnl[key]['paket'] if countPnl[key] else 0) +
#             (countSlk[key]['paket'] if countSlk[key] else 0) +
#             (countTdr[key]['paket'] if countTdr[key] else 0) +
#             (countTdc[key]['paket'] if countTdc[key] else 0) +
#             (countSwa[key]['paket'] if countSwa[key] else 0),
#         'pagu':
#             (countDik[key]['paket'] if countDik[key] else 0) +
#             (countEpc[key]['paket'] if countEpc[key] else 0) +
#             (countPdl[key]['paket'] if countPdl[key] else 0) +
#             (countPnl[key]['paket'] if countPnl[key] else 0) +
#             (countSlk[key]['paket'] if countSlk[key] else 0) +
#             (countTdr[key]['paket'] if countTdr[key] else 0) +
#             (countTdc[key]['paket'] if countTdc[key] else 0) +
#             (countSwa[key]['paket'] if countSwa[key] else 0)
#     }
#
# csv = [
#     [
#         "No.",
#         "Satuan Kerja",
#         "Paket Pengadaan Langsung",
#         "Nilai Pengadaan Langsung",
#         "Paket E-Purchasing",
#         "Nilai E-Purchasing",
#         "Paket Seleksi",
#         "Nilai Seleksi",
#         "Paket Tender",
#         "Nilai Tender",
#         "Paket Tender Cepat",
#         "Nilai Tender Cepat",
#         "Paket Dikecualikan",
#         "Nilai Dikecualikan",
#         "Paket Penunjukan Langsung",
#         "Nlai Penunjukan Langsung",
#         "Swakelola",
#         "Terumumkan",
#     ],
# ]
#
# detailRekap = []

# rekapKeys.forEach((key) = > {
#     detailRekap.push([
#         key,
#         countPdl[key] ? countPdl[key].paket: 0,
# countPdl[key] ? countPdl[key].pagu: 0,
# countEpc[key] ? countEpc[key].paket: 0,
# countEpc[key] ? countEpc[key].pagu: 0,
# countSlk[key] ? countSlk[key].paket: 0,
# countSlk[key] ? countSlk[key].pagu: 0,
# countTdr[key] ? countTdr[key].paket: 0,
# countTdr[key] ? countTdr[key].pagu: 0,
# countTdc[key] ? countTdc[key].paket: 0,
# countTdc[key] ? countTdc[key].pagu: 0,
# countDik[key] ? countDik[key].paket: 0,
# countDik[key] ? countDik[key].pagu: 0,
# countPnl[key] ? countPnl[key].paket: 0,
# countPnl[key] ? countPnl[key].pagu: 0,
# countSwakelola[key] ? countSwakelola[key].pagu: 0,
# countTotal[key] ? countTotal[key].pagu: 0,
# ]);
# });
#
# let
# hashMap = {};
# detailRekap.forEach(function(arr)
# {
#     hashMap[arr.join("|")] = arr;
# });
# let
# filter = Object.keys(hashMap).map(function(k)
# {
# return hashMap[k];
# });
#
# const
# total = filter;
#
# let
# i = 1;
# filter.forEach(myFunction);
# function
# myFunction(item)
# {
#     item.unshift(i);
# i + +;
# }
#
# countPdlTotal = {
#     paket: 0,
#     pagu: 0,
# };
# countEpcTotal = {
#     paket: 0,
#     pagu: 0,
# };
# countSlkTotal = {
#     paket: 0,
#     pagu: 0,
# };
# countTdrTotal = {
#     paket: 0,
#     pagu: 0,
# };
# countTdcTotal = {
#     paket: 0,
#     pagu: 0,
# };
# countDikTotal = {
#     paket: 0,
#     pagu: 0,
# };
# countPnlTotal = {
#     paket: 0,
#     pagu: 0,
# };
# countSwakelolaTotal = {
#     pagu: 0,
# };
# countTotalTotal = {
#     pagu: 0,
# };
#
# total.slice(1).forEach((pin) = > {
# countPdlTotal.paket += pin[2];
# countPdlTotal.pagu += pin[3];
# countEpcTotal.paket += pin[4];
# countEpcTotal.pagu += pin[5];
# countSlkTotal.paket += pin[6];
# countSlkTotal.pagu += pin[7];
# countTdrTotal.paket += pin[8];
# countTdrTotal.pagu += pin[9];
# countTdcTotal.paket += pin[10];
# countTdcTotal.pagu += pin[11];
# countDikTotal.paket += pin[12];
# countDikTotal.pagu += pin[13];
# countPnlTotal.paket += pin[14];
# countPnlTotal.pagu += pin[15];
# countSwakelolaTotal.pagu += pin[16];
# countTotalTotal.pagu += pin[17];
# });
#
# filter.push([
# "",
# "TOTAL",
# countPdlTotal.paket,
# countPdlTotal.pagu,
# countEpcTotal.paket,
# countEpcTotal.pagu,
# countSlkTotal.paket,
# countSlkTotal.pagu,
# countTdrTotal.paket,
# countTdrTotal.pagu,
# countTdcTotal.paket,
# countTdcTotal.pagu,
# countDikTotal.paket,
# countDikTotal.pagu,
# countPnlTotal.paket,
# countPnlTotal.pagu,
# countSwakelolaTotal.pagu,
# countTotalTotal.pagu,
# ]);
#
# const
# hasil = csv.concat(filter);
#
# var
# wb = utils.book_new(), \
# ws = utils.aoa_to_sheet(hasil);
# utils.book_append_sheet(wb, ws, "MySheet1");
# writeFileXLSX(wb, "Rekap.xlsx");
# } catch(e)
# {
#     console.log(e);
# }
# };
#
# const
# makeAPICall2 = async () = > {
# try {
# if (mtd == = "default") {
# return;
# }
#
# let
# response = await axios.get(
# "https://lpse.or.id/sirup/?mod=dataruppenyediakldi&idKldi=D170&tahun=" +
# year
# );
# const
# penyedia2 = response.data.aaData;
# response = await axios.get(
# "https://lpse.or.id/sirup/?mod=dataruppenyediaswakelolaallrekapkldi&idKldi=D170&tahun=" +
# year
# );
# const
# pdsw2 = response.data.aaData;
#
# response = await axios.get(
# "https://lpse.or.id/sirup/?mod=datarupswakelolakldi&idKldi=D170&tahun=" +
# year
# );
# const
# swa2 = response.data.aaData;
#
# let
# tmpTP = [];
# let
# tmpTPds = [];
# let
# tmpSwa = [];
#
# if (mtd === "slk")
# {
# let
# tenderPenyedia = [];
# let
# tenderPds = [];
#
# penyedia2.forEach((pen) = > {
#     let
# edit = pen[7].replace(
#        / January | February | March | April | May | June | July | August | September | October | November | December /,
# function(matched)
# {
# return mapObj[matched];
# }
# );
# if (pen[4] === "Seleksi") {
# tenderPenyedia.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[5],
# pen[6],
# edit,
# ]);
# }
# });
# tmpTP = tenderPenyedia.sort();
#
# pdsw2.forEach((pen) = > {
# let edit = pen[4].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[5] === "Seleksi") {
# tenderPds.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[6],
# pen[9],
# edit,
# ]);
# }
# });
# tmpTPds = tenderPds.sort();
# }
#
# if (mtd == = "tdr") {
# let tenderPenyedia =[];
# let tenderPds =[];
#
# penyedia2.forEach((pen) = > {
# let edit = pen[7].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[4] === "Tender") {
# tenderPenyedia.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[5],
# pen[6],
# edit,
# ]);
# }
# });
# tmpTP = tenderPenyedia.sort();
#
# pdsw2.forEach((pen) = > {
# let edit = pen[4].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[5] === "Tender") {
# tenderPds.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[6],
# pen[9],
# edit,
# ]);
# }
# });
# tmpTPds = tenderPds.sort();
# }
#
# if (mtd == = "tdc") {
# let tenderPenyedia =[];
# let tenderPds =[];
#
# penyedia2.forEach((pen) = > {
# let edit = pen[7].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[4] === "Tender Cepat") {
# tenderPenyedia.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[5],
# pen[6],
# edit,
# ]);
# }
# });
# tmpTP = tenderPenyedia.sort();
#
# pdsw2.forEach((pen) = > {
# let edit = pen[4].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[5] === "Tender Cepat") {
# tenderPds.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[6],
# pen[9],
# edit,
# ]);
# }
# });
# tmpTPds = tenderPds.sort();
# }
#
# if (mtd == = "pgl") {
# let pglPenyedia =[];
# let pglPds =[];
#
# penyedia2.forEach((pen) = > {
# let edit = pen[7].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[4] === "Pengadaan Langsung") {
# pglPenyedia.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[5],
# pen[6],
# edit,
# ]);
# }
# });
# tmpTP = pglPenyedia.sort();
# pdsw2.forEach((pen) = > {
# let edit = pen[4].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[5] === "Pengadaan Langsung") {
# pglPds.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[6],
# pen[9],
# edit,
# ]);
# }
# });
# tmpTPds = pglPds.sort();
# }
#
# if (mtd == = "ep") {
# let epPenyedia =[];
# let epPds =[];
#
# penyedia2.forEach((pen) = > {
# let edit = pen[7].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[4] === "e-Purchasing") {
# epPenyedia.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[5],
# pen[6],
# edit,
# ]);
# }
# });
# tmpTP = epPenyedia.sort();
# pdsw2.forEach((pen) = > {
# let edit = pen[4].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[5] === "e-Purchasing") {
# epPds.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[6],
# pen[9],
# edit,
# ]);
# }
# });
# tmpTPds = epPds.sort();
# }
#
# if (mtd == = "dk") {
# let dkPenyedia =[];
# let dkPds =[];
#
# penyedia2.forEach((pen) = > {
# let edit = pen[7].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[4] === "Dikecualikan") {
# dkPenyedia.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[5],
# pen[6],
# edit,
# ]);
# }
# });
# tmpTP = dkPenyedia.sort();
# pdsw2.forEach((pen) = > {
# let edit = pen[4].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[5] === "Dikecualikan") {
# dkPds.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[6],
# pen[9],
# edit,
# ]);
# }
# });
# tmpTPds = dkPds.sort();
# }
#
# if (mtd == = "pnl") {
# let pnlPenyedia =[];
# let pnlPds =[];
#
# penyedia2.forEach((pen) = > {
# let edit = pen[7].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[4] === "Penunjukan Langsung") {
# pnlPenyedia.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[5],
# pen[6],
# edit,
# ]);
# }
# });
# tmpTP = pnlPenyedia.sort();
# pdsw2.forEach((pen) = > {
# let edit = pen[4].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[5] === "Penunjukan Langsung") {
# pnlPds.push([
# pen[1],
# pen[2],
# parseInt(pen[3], 0),
# pen[6],
# pen[9],
# edit,
# ]);
# }
# });
# tmpTPds = pnlPds.sort();
# }
#
# if (mtd == = "swa") {
# let swaDetail =[];
# let edit2 = "";
#
# swa2.forEach((pen) = > {
# let edit = pen[6].replace(
# / January | February | March | April | May | June | July | August | September | October | November | December /,
# function (matched) {
# return mapObj[matched];
# }
# );
# if (pen[7].charAt(pen[7].length - 1) === ",") {
# edit2 = pen[7].slice(0, -1);
# } else {
# edit2 = pen[7];
# }
# swaDetail.push([
# pen[1],
# edit2,
# pen[2],
# parseInt(pen[3], 0),
# pen[4],
# pen[5],
# edit,
# ]);
# });
# tmpSwa = swaDetail.sort();
# }
#
# let
# tmpJoin = [];
# let
# header = [];
#
# if (tmpTP.length === 0 | | tmpTPds == = 0)
# {
#     tmpJoin = tmpSwa.sort();
# header = [
#     [
#         "No.",
#         "Satuan Kerja",
#         "Kegiatan",
#         "Paket",
#         "Pagu",
#         "Sumber Dana",
#         "Kode RUP",
#         "Waktu Pemilihan",
#     ],
# ];
# } else {
#     let
# kolek = tmpTP.concat(tmpTPds);
# tmpJoin = kolek.sort();
# header = [
#     [
#         "No.",
#         "Satuan Kerja",
#         "Paket",
#         "Pagu",
#         "Sumber Dana",
#         "Kode RUP",
#         "Waktu Pemilihan",
#     ],
# ];
# }
#
# let
# i = 1;
# tmpJoin.forEach(myFunction);
# function
# myFunction(item)
# {
#     item.unshift(i);
# i + +;
# }
#
# let
# gabung = header.concat(tmpJoin);
#
# var
# wb = utils.book_new(), \
# ws = utils.aoa_to_sheet(gabung);
# utils.book_append_sheet(wb, ws, "MySheet1");
# writeFileXLSX(wb, "Detail.xlsx");
# } catch(e)
# {
#     console.log(e);
# }
# };
