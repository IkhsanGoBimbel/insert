import { Injectable } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import * as path from 'path';
import * as fs from 'fs';
import { PrismaClient } from '@prisma/client';
import * as bcrypt from 'bcrypt';
@Injectable()
export class AppService {
  private readonly prisma: PrismaClient;

  constructor() {
    this.prisma = new PrismaClient();
  }

  getHello(): string {
    return 'Hello World!';
  }

  async readExcelFile() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'Data T_Siswa_5_Nov.xlsx');

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      let datas = [];
      const nohp = data.map((item) => item.c_HP1);
      const noreg = data.map((item) => item.c_NoRegistrasi);
      // const result = await this.prisma.t_siswa.findMany({
      //   // where : {
      //   //   c_nomor_hp : {
      //   //     in : nohp
      //   //   }
      //   // },
      //   where : {
      //     c_no_register : {
      //       in : noreg
      //     }
      //   },
      //   select : {
      //     c_email : true,
      //     c_is_aktif : true,
      //     c_nomor_hp : true,
      //     c_password : true,
      //     c_nama_lengkap : true,
      //     c_created_at : true,
      //     c_id_device : true,
      //     c_is_login : true,
      //     c_last_update_password : true,
      //     c_no_register: true
      //   }
      // })

      // const resultWithConvertedNoRegister = result?.map((item) => ({
      //   ...item,
      //   c_no_register: item?.c_no_register.toString(),
      // }));

      // const hapus = await this.prisma.t_siswa.deleteMany({
      //   where : {
      //     c_nomor_hp : {
      //       in : nohp
      //     }
      //   }
      // })
      const saltRounds = 10;
      const hashedData = await Promise.all(
        data.map(async (item) => {
          const hashedPassword = await bcrypt.hash(
            item.c_NoRegistrasi,
            saltRounds,
          );

          return {
            c_no_register: +item.c_NoRegistrasi,
            c_nama_lengkap: item.c_NamaLengkap,
            c_email: `${item.c_NoRegistrasi}@gmail.com`,
            c_nomor_hp: item.c_HP1,
            c_is_aktif: 'aktif',
            c_password: hashedPassword,
          };
        }),
      );

      const insertTable = await this.prisma.t_siswa.createMany({
        data: hashedData,
      });
      //  console.log(hashedData.length)
      //  console.log(insertTable)
      // return hashedData
      // return resultWithConvertedNoRegister
      return insertTable;
    } catch (error) {
      throw new Error('Error reading Excel file');
    }
  }

  async checkpass(username, pass) {
    try {
      const user = await this.prisma.t_siswa.findFirst({
        where: {
          c_no_register: username,
        },
      });
      const isPasswordValid = await bcrypt.compare(pass, user.c_password);
      return isPasswordValid;
    } catch (error) {
      return error;
    }
  }

  async insertpass(username, pass) {
    try {
      const saltRounds = 10;
      const hashedPassword = await bcrypt.hash(pass, saltRounds);
      console.log(username);
      console.log(hashedPassword);
      const user = await this.prisma.t_siswa.update({
        where: {
          c_no_register: username,
        },
        data: {
          c_password: hashedPassword,
        },
      });
      // return user
    } catch (error) {
      return error;
    }
  }

  async insertProdukAktif() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'noreg3.xlsx');

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the c_tanggal_akhir : '2023-10-22'headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const mappedData = await Promise.all(
        data.map(async (item) => {
          const tanggal_awal = new Date('2023-10-18');
          const tanggal_akhir = new Date('2023-10-22');
          return {
            c_no_register: +item.c_NoRegistrasi,
            c_id_produk: 40627,
            c_status: 'Aktif',
            c_tanggal_awal: tanggal_awal.toISOString(),
            c_tanggal_akhir: tanggal_akhir.toISOString(),
          };
        }),
      );

      const insertTable = await this.prisma.t_produk_aktif.createMany({
        data: mappedData,
      });
      // console.log(mappedData.length)
      return insertTable;
    } catch (error) {
      // console.log(error)
      return error;
    }
  }

  async insertProdukSiswa() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'noreg3.xlsx');

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the c_tanggal_akhir : '2023-10-22'headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const mappedData = await Promise.all(
        data.map(async (item) => {
          const tanggal_daftar = new Date('2023-10-17');

          return {
            c_id_pembelian: +item.c_IdPembelian,
            c_no_register: +item.c_NoRegistrasi,
            c_tanggal_daftar: tanggal_daftar.toISOString(),
            c_id_kelas: 0,
            c_tahun_ajaran: '2023/2024',
            c_id_dikdasken: 0,
            c_nama_lengkap: item.c_NamaLengkap,
            c_id_gedung: 1,
            c_id_komar: 1,
            c_id_kota: 1,
            c_id_sekolah: +item.c_IdSekolah,
            c_id_sekolah_kelas: 14,
            c_tingkat_sekolah_kelas: '12 SMA IPA',
            c_id_jenis_kelas: 11,
            c_kapasitas_max: 25,
            c_status_bayar: 'LUNAS',
            c_id_bundling: 798577,
            c_kerjasama: 'N',
            // c_id_kurikulum : 15
          };
        }),
      );

      const insertTable = await this.prisma.t_produk_siswa.createMany({
        data: mappedData,
      });

      return insertTable;
      // return mappedData
    } catch (error) {
      console.log(error);
      return error;
    }
  }

  async readExcelFiles() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(
        __dirname,
        '..',
        'data_hasil_jawaban_23_malam.xlsx',
      );

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      // const uniqueIdSoalSet = new Set();

      // data.forEach((item) => {
      //   uniqueIdSoalSet.add(item['ID Soal']);
      // });
      // const uniqueIdSoalArray = Array.from(uniqueIdSoalSet);

      // return uniqueIdSoalSet

      // const workbookBaru = new ExcelJS.Workbook();
      // const worksheetBaru = workbookBaru.addWorksheet('Data_coba');

      // // Menentukan header kolom
      // worksheetBaru.columns = [
      //   { header: 'No Register', key: 'noRegister' },
      //   { header: 'Kode Paket', key: 'kodePaket' },
      //   { header: 'ID Soal', key: 'idSoal' },
      //   { header: 'ID Kelompok Ujian', key: 'idKelompokUjian' },
      //   { header: 'Nama Kelompok Ujian', key: 'namaKelompokUjian' },
      //   { header: 'Tipe Soal', key: 'tipeSoal' },
      //   { header: 'Tingkat Kesulitan', key: 'tingkatKesulitan' },
      //   { header: 'Jawaban Siswa', key: 'jawabanSiswa' },
      // ];

      // // console.log(worksheet_baru)

      // data.forEach((items) => {
      //   let noreg = items.c_no_register;
      //   if (items['c_detil_jawaban']) {
      //     let detilJawaban = JSON.parse(items['c_detil_jawaban']);
      //     if (detilJawaban.detil) {
      //       detilJawaban.detil.forEach((datanya) => {
      //         worksheetBaru.addRow({
      //           noRegister: noreg,
      //           kodePaket: datanya.kodePaket,
      //           idSoal: datanya.idSoal,
      //           idKelompokUjian: datanya.idKelompokUjian,
      //           namaKelompokUjian: datanya.namaKelompokUjian,
      //           tipeSoal: datanya.tipeSoal,
      //           tingkatKesulitan: datanya.tingkatKesulitan,
      //           jawabanSiswa: datanya.jawabanSiswa,
      //         });
      //       });
      //     }
      //   }
      // });

      // const filePathBaru = 'Datanya_baru.xlsx';
      // await workbookBaru.xlsx.writeFile(filePathBaru);

      // return filePathBaru;

      //=====================================================================
      const workbookBaru = new ExcelJS.Workbook();
      const worksheetBaru = workbookBaru.addWorksheet('Data_baru');

      // // // Menentukan header kolom
      worksheetBaru.columns = [
        { header: 'No Register', key: 'noRegister' },
        { header: 'Kode Paket', key: 'kodePaket' },
        { header: 'Nama Kelompok Ujian', key: 'namaKelompokUjian' },
        { header: 'Benar', key: 'benar' },
        { header: 'Salah', key: 'salah' },
        { header: 'Kosong', key: 'kosong' },
      ];

      data.forEach((items) => {
        let noreg = items.c_no_register;
        let kodePaket = items.c_kode_paket;
        if (items['c_detil_hasil']) {
          let detilHasil = JSON.parse(items['c_detil_hasil']);
          if (detilHasil.hasil) {
            detilHasil.hasil.forEach((datanya) => {
              worksheetBaru.addRow({
                noRegister: noreg,
                kodePaket: kodePaket,
                namaKelompokUjian: datanya.namaKelompokUjian,
                benar: datanya.benar,
                salah: datanya.salah,
                kosong: datanya.kosong,
              });
            });
          }
        }
      });

      const filePathBaru = 'Data_hasil_23_malam.xlsx';
      await workbookBaru.xlsx.writeFile(filePathBaru);
      return data;
    } catch (error) {
      throw new Error('Error reading Excel file');
    }
  }

  async bacaExcel() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'Data_T_Produk_Siswa.xlsx');

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      return data;
    } catch (error) {
      throw new Error('Error reading Excel file');
    }
  }

  async bacaExcel_bcrypt() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'Data T_Siswa_7pagi_2.xlsx');

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const saltRounds = 10;
      const hashedData = await Promise.all(
        data.map(async (item) => {
          const hashedPassword = await bcrypt.hash(item.c_password, saltRounds);

          return {
            c_no_register: +item.c_no_register,
            c_nama_lengkap: item.c_nama_lengkap,
            c_email: item.c_email,
            c_nomor_hp: item.c_nomor_hp,
            c_is_aktif: 'aktif',
            c_password: hashedPassword,
          };
        }),
      );

      // return hashedData
      const insertData = await this.prisma.t_siswa.createMany({
        data: hashedData,
      });
      return insertData;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async bacaExcel_cari_t_siswa() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'Data T_Siswa_7pagi_2.xlsx');

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const nohp = data.map((item) => item.c_nomor_hp);
      const noreg = data.map((item) => item.c_no_register);

      const findSiswa = await this.prisma.t_siswa.findMany({
        where: {
          c_no_register: {
            in: noreg,
          },
        },
      });

      const siswaNoreg = findSiswa.map((item) => {
        return {
          c_no_register: item.c_no_register.toString().padStart(12, '0'),
          c_email: item.c_email,
          c_nomor_hp: item.c_nomor_hp,
          c_nama_lengkap: item.c_nama_lengkap,
        };
      });

      return siswaNoreg;
    } catch (error) {
      throw new Error('Error reading Excel file');
    }
  }

  async bacaExcel_insert_t_produk_siswa() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(
        __dirname,
        '..',
        'data t_produk_siswa_Dummy.xlsx',
      );

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const datanya = data.map((item) => {
        return {
          c_id_pembelian: +item.c_id_pembelian,
          c_no_register: +item.c_no_register,
          c_tanggal_daftar: item.c_tanggal_daftar,
          c_id_kelas: +item.c_id_kelas,
          c_tahun_ajaran: item.c_tahun_ajaran,
          c_id_dikdasken: +item.c_id_dikdasken,
          c_nama_lengkap: item.c_nama_lengkap,
          c_id_gedung: +item.c_id_gedung,
          c_id_komar: +item.c_id_komar,
          c_id_kota: +item.c_id_kota,
          c_id_sekolah: +item.c_id_sekolah,
          c_id_sekolah_kelas: +item.c_id_sekolah_kelas,
          c_tingkat_sekolah_kelas: item.c_tingkat_sekolah_kelas,
          c_id_jenis_kelas: +item.c_id_jenis_kelas,
          c_kapasitas_max: +item.c_kapasitas_max,
          c_status_bayar: item.c_status_bayar,
          c_id_bundling: +item.c_id_bundling,
          c_kerjasama: item.c_kerjasama,
        };
      });
      const insertData = await this.prisma.t_produk_siswa.createMany({
        data: datanya,
      });
      return insertData;
      // return datanya
      // return datanya
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async bacaExcel_cari_produk_aktif() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(
        __dirname,
        '..',
        'Data T_produk_siswa_7pagi_1.xlsx',
      );

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      // const produk_mix = [
      //   {
      //     c_id_produk_mix: 13119,
      //     c_id_produk: [
      //       32796, 33233, 33423, 40792, 40793, 33619, 33630, 33695, 33708,
      //       40794, 40795, 40796, 40797, 40798, 40799, 40800, 40801, 40802,
      //       33295, 33296, 33164, 40803, 40804, 33520, 40805, 40806, 40807,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13112,
      //     c_id_produk: [
      //       32796, 33238, 33449, 40808, 40809, 33619, 33655, 33695, 33708,
      //       40810, 40811, 40812, 40813, 40814, 40815, 40816, 40817, 40818,
      //       33346, 33347, 33190, 40819, 40820, 33559, 40821, 40822, 40823,
      //     ],
      //   },
      // ];
      // const produk_mix = [
      //   {
      //     c_id_produk_mix: 13119,
      //     c_id_produk: [
      //      40803
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13120,
      //     c_id_produk: [
      //      40819
      //     ],
      //   },
      // ];
      const produk_mix = [
        {
          c_id_produk_mix: 13163,
          c_id_produk: [40870],
        },
        {
          c_id_produk_mix: 13164,
          c_id_produk: [40871],
        },
        {
          c_id_produk_mix: 13165,
          c_id_produk: [40872],
        },
        {
          c_id_produk_mix: 13166,
          c_id_produk: [40873],
        },
        {
          c_id_produk_mix: 13182,
          c_id_produk: [40866],
        },
        {
          c_id_produk_mix: 13183,
          c_id_produk: [40867],
        },
        {
          c_id_produk_mix: 13184,
          c_id_produk: [40868],
        },
      ];

      const result = data.flatMap((pembelian) => {
        const produkMixData = produk_mix.find(
          (mix) => mix.c_id_produk_mix === pembelian.id_produk_mix,
        );
        return produkMixData
          ? produkMixData.c_id_produk.map((id_produk) => ({
              c_no_register: pembelian.c_no_register,
              c_id_produk: id_produk,
            }))
          : [];
      });

      const workbookBaru = new ExcelJS.Workbook();
      const worksheetBaru = workbookBaru.addWorksheet('Data_coba');

      // Menentukan header kolom
      worksheetBaru.columns = [
        { header: 'c_no_register', key: 'c_no_register' },
        { header: 'c_id_produk', key: 'c_id_produk' },
        { header: 'c_status', key: 'c_status' },
        { header: 'c_tanggal_awal', key: 'c_tanggal_awal' },
        { header: 'c_tanggal_akhir', key: 'c_tanggal_akhir' },
      ];

      // console.log(worksheet_baru)

      result.forEach((items) => {
        worksheetBaru.addRow({
          c_no_register: items.c_no_register,
          c_id_produk: items.c_id_produk,
          c_status: 'Aktif',
          c_tanggal_awal: '2023-10-27',
          c_tanggal_akhir: '2024-10-27',
        });
      });

      const filePathBaru = 't_produk-aktif-7pagi_1Fix.xlsx';
      await workbookBaru.xlsx.writeFile(filePathBaru);

      return filePathBaru;
    } catch (error) {
      throw new Error('Error reading Excel file');
    }
  }

  async bacaExcel_insert_t_produk_aktif() {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(
        __dirname,
        '..',
        't_produk-aktif-7pagi_1Fix.xlsx',
      );

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const datanya = data.map((item, index) => {
        const c_idnya = 6002670;
        return {
          c_no_register: +item.c_no_register,
          c_id_produk: +item.c_id_produk,
          c_status: item.c_status,
          c_tanggal_awal: new Date(item.c_tanggal_awal),
          c_tanggal_akhir: new Date(item.c_tanggal_akhir),
          c_id: c_idnya + (index + 1),
        };
      });

      // return datanya
      const insertData = await this.prisma.t_produk_aktif.createMany({
        data: datanya,
        skipDuplicates: true,
      });

      return insertData;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async uploadbacaExcel(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }
      return data;
    } catch (error) {
      throw new Error('Error reading Excel file');
    }
  }

  async uploadbcryptExcelCariT_Siswa(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const noreg = data.map((item) => item.c_no_register);

      const findSiswa = await this.prisma.t_siswa.findMany({
        where: {
          c_no_register: {
            in: noreg,
          },
        },
      });
      console.log(findSiswa.length);
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }
      const siswaNoreg = findSiswa.map((item) => {
        return {
          c_no_register: item.c_no_register.toString().padStart(12, '0'),
          c_email: item.c_email,
          c_nomor_hp: item.c_nomor_hp,
          c_nama_lengkap: item.c_nama_lengkap,
        };
      });

      return siswaNoreg;
    } catch (error) {
      console.log(error);
      return error;
    }
  }

  async uploadExcel_Hapus_T_Siswa(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const noreg = data.map((item) => item.c_no_register);

      const deleteSiswa = await this.prisma.t_siswa.deleteMany({
        where: {
          c_no_register: {
            in: noreg,
          },
        },
      });
      console.log(deleteSiswa);
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }
      return deleteSiswa;
    } catch (error) {
      console.log(error);
      return error;
    }
  }

  async uploadbcryptExcel(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const saltRounds = 10;
      const hashedData = await Promise.all(
        data.map(async (item, index) => {
          console.log('count bcrypt', index);
          const hashedPassword = await bcrypt.hash(item.c_password, saltRounds);
          const noregstring = item.c_no_register.toString();
          return {
            c_no_register: +item.c_no_register,
            c_nama_lengkap: item.c_nama_lengkap,
            c_email: !item.c_email
              ? `${item.c_noregister}@gmail.com`
              : item.c_email.toString(),
            c_nomor_hp: !item.c_nomor_hp
              ? `0824123456${index}`
              : item.c_nomor_hp,
            c_is_aktif: 'aktif',
            c_password: hashedPassword,
          };
        }),
      );

      const insertData = await this.prisma.t_siswa.createMany({
        data: hashedData,
        skipDuplicates: true,
      });
      // let count;
      // hashedData.map(async (item, index) => {
      //   console.log('update', index);
      //   const update = await this.prisma.t_siswa.upsert({
      //     where: {
      //       c_no_register: item.c_no_register,
      //     },
      //     create: {
      //       c_no_register: item.c_no_register,
      //       c_password: item.c_password,
      //       c_email: item.c_email,
      //       c_is_aktif: item.c_is_aktif,
      //       c_nama_lengkap: item.c_nama_lengkap,
      //       c_nomor_hp: item.c_nomor_hp,
      //       c_created_at: new Date(),
      //     },
      //     update: {},
      //   });
      // });

      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }
      return insertData;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async uploadT_produkSiswaRead(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });
      const datanya = data.map((item) => {
        return {
          c_id_pembelian: +item.c_id_pembelian,
          c_no_register: +item.c_no_register,
          c_tanggal_daftar: item.c_tanggal_daftar,
          c_id_kelas: +item.c_id_kelas,
          c_tahun_ajaran: item.c_tahun_ajaran,
          c_id_dikdasken: +item.c_id_dikdasken,
          c_nama_lengkap: item.c_nama_lengkap,
          c_id_gedung: +item.c_id_gedung,
          c_id_komar: +item.c_id_komar,
          c_id_kota: +item.c_id_kota,
          c_id_sekolah: +item.c_id_sekolah,
          c_id_sekolah_kelas: +item.c_id_sekolah_kelas,
          c_tingkat_sekolah_kelas: item.c_tingkat_sekolah_kelas,
          c_id_jenis_kelas: +item.c_id_jenis_kelas,
          c_kapasitas_max: +item.c_kapasitas_max,
          c_status_bayar: item.c_status_bayar,
          c_id_bundling: +item.c_id_bundling,
          c_kerjasama: item.c_kerjasama,
        };
      });

      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }
      return datanya;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async uploadT_produkSiswaInsert(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const datanya = data.map((item) => {
        return {
          c_id_pembelian: +item.c_id_pembelian,
          c_no_register: +item.c_no_register,
          c_tanggal_daftar: new Date(item.c_tanggal_daftar),
          c_id_kelas: +item.c_id_kelas,
          c_tahun_ajaran: item.c_tahun_ajaran,
          c_id_dikdasken: +item.c_id_dikdasken,
          c_nama_lengkap: item.c_nama_lengkap,
          c_id_gedung: +item.c_id_gedung,
          c_id_komar: +item.c_id_komar,
          c_id_kota: +item.c_id_kota,
          c_id_sekolah: +item.c_id_sekolah,
          c_id_sekolah_kelas: +item.c_id_sekolah_kelas,
          c_tingkat_sekolah_kelas: item.c_tingkat_sekolah_kelas,
          c_id_jenis_kelas: +item.c_id_jenis_kelas,
          c_kapasitas_max: +item.c_kapasitas_max,
          c_status_bayar: item.c_status_bayar,
          c_id_bundling: +item.c_id_bundling,
          c_kerjasama: item.c_kerjasama,
        };
      });

      const insertData = await this.prisma.t_produk_siswa.createMany({
        data: datanya,
        skipDuplicates: true,
      });

      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }
      return insertData;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async uploadT_produkSiswaFind(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const datanya = data.map((item) => {
        return {
          c_id_pembelian: +item.c_id_pembelian,
          c_no_register: +item.c_no_register,
          c_tanggal_daftar: item.c_tanggal_daftar,
          c_id_kelas: +item.c_id_kelas,
          c_tahun_ajaran: item.c_tahun_ajaran,
          c_id_dikdasken: +item.c_id_dikdasken,
          c_nama_lengkap: item.c_nama_lengkap,
          c_id_gedung: +item.c_id_gedung,
          c_id_komar: +item.c_id_komar,
          c_id_kota: +item.c_id_kota,
          c_id_sekolah: +item.c_id_sekolah,
          c_id_sekolah_kelas: +item.c_id_sekolah_kelas,
          c_tingkat_sekolah_kelas: item.c_tingkat_sekolah_kelas,
          c_id_jenis_kelas: +item.c_id_jenis_kelas,
          c_kapasitas_max: +item.c_kapasitas_max,
          c_status_bayar: item.c_status_bayar,
          c_id_bundling: +item.c_id_bundling,
          c_kerjasama: item.c_kerjasama,
        };
      });

      const noregnya = datanya.map((item) => item.c_no_register);
      const id_bundling = datanya.map((item) => item.c_id_bundling);
      const find = await this.prisma.t_produk_siswa.findMany({
        where: {
          c_no_register: {
            in: noregnya,
          },
          // c_id_bundling: {
          //   in: id_bundling,
          // },
        },
      });

      const siswaNoreg = find.map((item) => {
        return {
          c_no_register: item.c_no_register.toString().padStart(12, '0'),
        };
      });

      console.log(siswaNoreg.length);
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }
      return siswaNoreg;
    } catch (error) {
      console.log(error);
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async uploadT_produkSiswa_delete(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const datanya = data.map((item) => {
        return {
          c_id_pembelian: +item.c_id_pembelian,
          c_no_register: +item.c_no_register,
          c_tanggal_daftar: item.c_tanggal_daftar,
          c_id_kelas: +item.c_id_kelas,
          c_tahun_ajaran: item.c_tahun_ajaran,
          c_id_dikdasken: +item.c_id_dikdasken,
          c_nama_lengkap: item.c_nama_lengkap,
          c_id_gedung: +item.c_id_gedung,
          c_id_komar: +item.c_id_komar,
          c_id_kota: +item.c_id_kota,
          c_id_sekolah: +item.c_id_sekolah,
          c_id_sekolah_kelas: +item.c_id_sekolah_kelas,
          c_tingkat_sekolah_kelas: item.c_tingkat_sekolah_kelas,
          c_id_jenis_kelas: +item.c_id_jenis_kelas,
          c_kapasitas_max: +item.c_kapasitas_max,
          c_status_bayar: item.c_status_bayar,
          c_id_bundling: +item.c_id_bundling,
          c_kerjasama: item.c_kerjasama,
        };
      });

      const datanoreg = datanya.map((item) => item.c_no_register);
      const databundling = datanya.map((item) => item.c_id_bundling);

      const find = await this.prisma.t_produk_siswa.deleteMany({
        where: {
          c_id_bundling: {
            in: databundling,
          },
          c_no_register: {
            in: datanoreg,
          },
        },
      });

      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }

      return find;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async uploadT_produkSiswaBuildProdukAktif(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      //pgts UAT GO EXPERT
      // const produk_mix = [
      //   {
      //     c_id_produk_mix: 13111,
      //     c_id_produk: [
      //       40807, 40793, 33708, 33619, 33296, 33295, 40802, 33423, 40797,
      //       33630, 32796, 40805, 40804, 40799, 40798, 33520, 33233, 33695,
      //       33164, 40803, 40801, 40800, 40796, 40792, 40806, 40795, 40794,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13112,
      //     c_id_produk: [
      //       40813, 40811, 33449, 40822, 40819, 40818, 40815, 33619, 40812,
      //       33708, 33190, 40823, 40821, 40820, 33238, 33695, 33655, 33559,
      //       33347, 40816, 40814, 40817, 40810, 32796, 40809, 40808, 33346,
      //     ],
      //   },
      // ];

      //kebanjahe
      // const produk_mix = [
      //   { c_id_produk_mix: 13229, c_id_produk: [40930] },
      //   { c_id_produk_mix: 13231, c_id_produk: [40932] },
      // ];

      //Magelang
      // const produk_mix = [
      //   { c_id_produk_mix: 13175, c_id_produk: [40879] },
      //   { c_id_produk_mix: 13176, c_id_produk: [40880] },
      //   { c_id_produk_mix: 13180, c_id_produk: [40884] },
      //   { c_id_produk_mix: 13181, c_id_produk: [40885] },
      // ];

      //10 november bimker
      const produk_mix = [
        { c_id_produk_mix: 13175, c_id_produk: [40879] },
        { c_id_produk_mix: 13176, c_id_produk: [40880] },
        { c_id_produk_mix: 13180, c_id_produk: [40884] },
        { c_id_produk_mix: 13181, c_id_produk: [40885] },
      ];

      //UAT MOBILE TOBK
      // const produk_mix = [
      //   {
      //     c_id_produk_mix: 13111,
      //     c_id_produk: [40803],
      //   },
      //   {
      //     c_id_produk_mix: 13112,
      //     c_id_produk: [40819],
      //   },
      // ];

      //12 ribu tobk
      // const produk_mix = [
      //   {
      //     c_id_produk_mix: 7363,
      //     c_id_produk: [
      //       17581, 19950, 23279, 23327, 23374, 23411, 23789, 23801, 23876,
      //       30841, 31041, 23357, 23781, 30843, 31246, 17369, 17598, 19007,
      //       23456, 23522, 23550, 23785, 30009, 30507, 30830, 30975, 17413,
      //       19009, 23289, 23313, 23739, 23756, 23877, 31233, 16934, 17531,
      //       17688, 23744, 23762, 23769, 23772, 31056, 31207, 17470, 17546,
      //       23388, 23764, 23776, 30357, 30368, 30965, 17641, 18943, 23255,
      //       23455, 23738, 23748, 23752, 30351, 23341, 23793, 30459, 30985,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 7749,
      //     c_id_produk: [
      //       18053, 19255, 23436, 23478, 23727, 23754, 23765, 30973, 18161,
      //       21539, 23287, 23409, 23730, 23731, 23746, 23750, 23779, 23799,
      //       30509, 30878, 18188, 23791, 30008, 30770, 31038, 31054, 21338,
      //       21637, 21712, 23267, 23728, 23742, 23872, 30453, 17890, 18021,
      //       18072, 21709, 23339, 23729, 23774, 23787, 30011, 30409, 31231,
      //       21282, 23277, 23325, 23520, 23548, 23735, 23766, 23873, 31019,
      //       31205, 18038, 19414, 23353, 23452, 23758, 30366, 31244, 18224,
      //       19207, 23231, 23242, 23253, 23386, 23726, 23759, 23767, 23768,
      //       23804,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 9048,
      //     c_id_produk: [
      //       19986, 23480, 23550, 23614, 30368, 30866, 31207, 31246, 18099,
      //       23269, 23369, 23605, 23676, 17917, 18070, 18172, 23455, 23575,
      //       18036, 23313, 23327, 30975, 19781, 23289, 23576, 30351, 30854,
      //       30872, 18051, 18212, 19950, 23341, 23456, 30010, 31040, 31220,
      //       17892, 21635, 23357, 23374, 23509, 23647, 30357, 30459, 17974,
      //       18019, 23255, 23279, 23388, 23411, 23522, 30514,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 9147,
      //     c_id_produk: [
      //       21633, 23281, 30874, 19956, 23315, 23524, 23534, 23540, 23546,
      //       23552, 30378, 30465, 30469, 30659, 30853, 31209, 18034, 18112,
      //       23329, 23413, 23482, 23233, 23511, 30406, 31222, 17894, 21608,
      //       23257, 23390, 30885, 31248, 18017, 18067, 18097, 18115, 21738,
      //       23271, 23376, 23450, 30370, 23244, 30868, 17972, 18049, 18105,
      //       21979, 23344, 23432, 23443, 30000, 30379,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12106,
      //     c_id_produk: [
      //       33446, 33695, 32796, 32832, 33346, 33708, 39210, 33776, 33050,
      //       33551, 32833, 32915, 33049, 33550, 33552, 39080, 32649, 32689,
      //       32978, 33447, 33449, 33520, 33619, 33655, 38525, 33238, 33345,
      //       33347, 32758, 33190, 38736, 40286,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12107,
      //     c_id_produk: [
      //       32821, 32892, 33708, 39163, 39834, 33295, 33296, 33695, 32647,
      //       32950, 33024, 33164, 33523, 38516, 33023, 33233, 33294, 33630,
      //       33744, 32664, 32733, 32822, 33422, 33522, 32796, 33521, 38444,
      //       38733, 39030, 33520, 33619, 33423, 38737,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12108,
      //     c_id_produk: [
      //       32708, 32777, 32950, 33069, 33674, 32650, 32853, 33295, 33366,
      //       33524, 33576, 33804, 39104, 40268, 32852, 33467, 32796, 38730,
      //       32933, 33257, 33619, 33297, 33468, 33695, 33708, 33577, 38522,
      //       39236, 33070, 33210, 33520, 40288,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12112,
      //     c_id_produk: [
      //       33024, 33296, 33523, 33619, 32647, 32951, 33423, 33521, 33522,
      //       38516, 38733, 39156, 33023, 33294, 33422, 33695, 39110, 39240,
      //       32664, 38444, 32822, 33164, 33520, 39834, 32733, 32796, 32892,
      //       33233, 33708, 33745, 32821, 33295, 33630, 38737,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12113,
      //     c_id_produk: [
      //       32758, 32833, 32916, 33446, 38736, 32689, 32832, 33238, 33551,
      //       33552, 33619, 33708, 38525, 33520, 32979, 33050, 39205, 33346,
      //       33447, 33550, 33695, 32796, 33345, 33449, 33655, 32649, 33049,
      //       33777, 39132, 40286, 33190, 33347, 40168,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12115,
      //     c_id_produk: [
      //       32796, 32853, 32933, 33069, 33210, 33468, 33805, 38730, 32650,
      //       33070, 33695, 38522, 33576, 39122, 39170, 33520, 40268, 32777,
      //       32951, 33257, 33295, 33524, 32852, 33366, 33467, 33577, 33619,
      //       33674, 33708, 40288, 32708, 33297,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12117,
      //     c_id_produk: [
      //       32726, 33405, 33406, 38663, 32885, 33407, 33502, 33226, 33278,
      //       33619, 38944, 33010, 33503, 32794, 33279, 33623, 33704, 32808,
      //       33009, 33157, 33504, 32807, 33691, 38392, 32657, 33277, 33721,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12119,
      //     c_id_produk: [
      //       33622, 33691, 33720, 32725, 32805, 32656, 33275, 33403, 33404,
      //       32884, 33007, 33008, 33156, 33274, 39800, 33501, 33619, 33704,
      //       38668, 40258, 33225, 32794, 32806, 33276, 33499, 33500, 38360,
      //       38905,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12121,
      //     c_id_produk: [
      //       32727, 33619, 33691, 38401, 32794, 33012, 33704, 33278, 33409,
      //       33725, 32658, 33158, 32809, 32810, 33011, 33227, 33279, 33505,
      //       33506, 33624, 32886, 40259, 33280, 33503, 38664, 33408, 38913,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12123,
      //     c_id_produk: [
      //       32660, 32814, 33229, 33284, 33286, 33413, 32813, 33619, 38436,
      //       32795, 38990, 33285, 38673, 40260, 33510, 33707, 33706, 40279,
      //       32729, 33015, 33016, 33160, 33626, 32888, 33414, 33511, 33692,
      //       33731,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12125,
      //     c_id_produk: [
      //       33289, 32661, 32816, 33735, 32889, 33161, 33692, 38432, 32730,
      //       33238, 33287, 33416, 33417, 33705, 33512, 32795, 33018, 33288,
      //       33513, 33415, 33619, 32815, 33017, 33514, 33627, 38689, 39001,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12127,
      //     c_id_produk: [
      //       32952, 33120, 33520, 33523, 39010, 39122, 39170, 32647, 32892,
      //       33024, 33233, 33522, 33695, 38737, 32733, 33164, 33294, 33296,
      //       32664, 32821, 33708, 38516, 39139, 32822, 33121, 33295, 33619,
      //       33630, 33746, 39834, 32796, 33521, 38444, 40265, 33023, 33422,
      //       33423, 33805, 38733,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12130,
      //     c_id_produk: [
      //       32821, 32892, 33294, 33520, 33708, 38733, 39834, 32796, 33164,
      //       32822, 33295, 33521, 33523, 33619, 32733, 33233, 33296, 33695,
      //       39260, 33024, 33630, 32664, 33124, 33423, 33747, 38444, 40265,
      //       32647, 33023, 32953, 33422, 33522, 39013,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12145,
      //     c_id_produk: [
      //       32658, 33280, 33619, 32810, 32794, 32886, 33158, 33278, 33279,
      //       33408, 40259, 33505, 33624, 33725, 32727, 33012, 33506, 38913,
      //       33011, 33704, 33409, 38664, 32809, 33227, 33503, 33691, 38401,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12152,
      //     c_id_produk: [
      //       33016, 33160, 33284, 33413, 33414, 33619, 33706, 32660, 33707,
      //       33733, 32729, 32813, 33626, 33286, 33692, 38673, 32814, 33511,
      //       40260, 32888, 33015, 33510, 38436, 32795, 33229, 33285, 39825,
      //       40279,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12153,
      //     c_id_produk: [
      //       32818, 32890, 33516, 40296, 32817, 33019, 33290, 33418, 33020,
      //       33230, 33628, 40517, 32662, 32795, 33291, 33292, 39288, 32731,
      //       33419, 33515, 33723, 40572, 33162, 33619, 39316, 33510, 33692,
      //       33705,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12157,
      //     c_id_produk: [
      //       32730, 33416, 33514, 33627, 33415, 33417, 33513, 33737, 38689,
      //       33017, 33287, 33289, 32816, 32889, 33018, 33161, 33238, 33288,
      //       33619, 33692, 32795, 33705, 32661, 33512, 38984, 32815, 38432,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12160,
      //     c_id_produk: [
      //       32954, 33296, 33523, 33748, 38733, 32822, 32892, 33023, 33024,
      //       33121, 39018, 32664, 33120, 33233, 33295, 33695, 33708, 39122,
      //       39149, 32733, 32821, 33422, 33521, 40265, 32647, 33294, 33520,
      //       33630, 32796, 33423, 38444, 39170, 33522, 33619, 33164, 33805,
      //       38516, 38737, 39834,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12162,
      //     c_id_produk: [
      //       33294, 33296, 33619, 33745, 39834, 32796, 32821, 32822, 32951,
      //       33521, 38737, 39156, 32733, 33233, 33522, 33695, 39110, 39240,
      //       33023, 33423, 38444, 38733, 32664, 33024, 33630, 32892, 33422,
      //       38516, 32647, 33164, 33520, 33523, 33708, 33295,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12165,
      //     c_id_produk: [
      //       32916, 33190, 39058, 39192, 40285, 33778, 32796, 32832, 33049,
      //       33346, 33446, 33551, 33619, 32649, 33125, 33447, 33552, 33708,
      //       33345, 33347, 33449, 33695, 38525, 32689, 32980, 33520, 33550,
      //       32758, 33126, 33238, 33655, 38736, 32833, 33050,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12166,
      //     c_id_produk: [
      //       33552, 32916, 33050, 33238, 33347, 33520, 33550, 33619, 33695,
      //       32649, 32689, 32796, 32833, 33345, 33446, 32832, 33049, 33551,
      //       39062, 39197, 40285, 33449, 33655, 32758, 32915, 33346, 38525,
      //       33447, 33708, 32981, 33190, 33779, 38736,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12169,
      //     c_id_produk: [
      //       32916, 33345, 33520, 39132, 40168, 32796, 33049, 33238, 33346,
      //       33550, 33449, 33695, 33708, 40285, 33190, 33347, 33552, 33446,
      //       32689, 33050, 33447, 33655, 39205, 32649, 33777, 32758, 32832,
      //       32833, 32979, 33551, 33619, 38736,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12173,
      //     c_id_produk: [
      //       32708, 32954, 33257, 33297, 33577, 33674, 33708, 33808, 38730,
      //       33069, 33131, 33210, 33619, 39226, 33295, 33695, 32650, 32777,
      //       33130, 33520, 33524, 32853, 33366, 38522, 39093, 32796, 32933,
      //       33468, 33576, 40268, 40288, 32852, 33070, 33467,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12174,
      //     c_id_produk: [
      //       33674, 40268, 32777, 32951, 33366, 33295, 33577, 33619, 38730,
      //       32853, 33468, 33520, 32796, 32852, 33805, 38522, 32650, 33576,
      //       33070, 33467, 33524, 39122, 39170, 32708, 33257, 33708, 32933,
      //       33069, 33210, 33297, 33695, 40288,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12175,
      //     c_id_produk: [
      //       32709, 33470, 40586, 32651, 33527, 33578, 33367, 33368, 33525,
      //       33809, 40582, 32778, 33369, 33696, 40256, 40302, 32854, 33165,
      //       33469, 33619, 40253, 40571, 32797, 33005, 33258, 40242, 40569,
      //       33071, 33526, 33675, 32934, 33711,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12178,
      //     c_id_produk: [
      //       32652, 32855, 40243, 33072, 33258, 33810, 40257, 40574, 33006,
      //       33191, 33368, 33554, 33696, 40570, 32935, 33471, 33472, 40316,
      //       40583, 32710, 32797, 33370, 33675, 40247, 33553, 33619, 33371,
      //       33579, 40573, 32779, 33555, 33711, 40254, 40303,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12209,
      //     c_id_produk: [
      //       32808, 33226, 33623, 38392, 33503, 33504, 33405, 33407, 33691,
      //       32657, 33704, 32794, 33010, 33277, 33279, 33502, 32885, 33406,
      //       33619, 33157, 33278, 32807, 33009,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12213,
      //     c_id_produk: [
      //       32726, 33010, 33226, 33277, 33502, 32807, 33406, 33619, 33721,
      //       32794, 33623, 33704, 38663, 33407, 33503, 33504, 33157, 33279,
      //       33405, 38944, 38392, 32657, 32808, 32885, 33009, 33278, 33691,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12215,
      //     c_id_produk: [
      //       32806, 33500, 33008, 33225, 33275, 33499, 33720, 38360, 32805,
      //       33156, 32725, 33007, 33276, 32656, 32884, 33404, 33501, 40307,
      //       32794, 33403, 33619, 33274, 38668, 40258, 33622, 33691, 33704,
      //       38905,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12228,
      //     c_id_produk: [
      //       32794, 32809, 33011, 33408, 33227, 33279, 33503, 33619, 33624,
      //       38913, 32658, 33278, 33725, 32727, 32886, 33158, 33691, 33506,
      //       32810, 33012, 33280, 33409, 33505, 33704, 38401, 38664, 40259,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12229,
      //     c_id_produk: [
      //       32794, 32812, 33013, 33159, 33411, 33508, 39812, 33283, 33014,
      //       33507, 33704, 32728, 33509, 33625, 33228, 33282, 33410, 33619,
      //       39816, 32811, 33412, 33727, 39810, 32659, 32887, 33281, 33691,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12232,
      //     c_id_produk: [
      //       32727, 33012, 32794, 33503, 33011, 33158, 33280, 38401, 32886,
      //       33408, 33505, 32658, 33624, 33691, 38664, 32809, 33278, 33704,
      //       33227, 33279, 33409, 33619, 32810, 33506,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12235,
      //     c_id_produk: [
      //       32886, 33227, 33408, 33506, 38378, 33011, 33691, 33619, 33278,
      //       33624, 38401, 40259, 33409, 33505, 32658, 33012, 33158, 33503,
      //       32794, 32809, 32810, 33279, 33280, 33704,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12240,
      //     c_id_produk: [
      //       32809, 32886, 33704, 32810, 33691, 33503, 33624, 32658, 33158,
      //       33280, 33409, 33505, 33619, 38401, 40259, 33279, 33506, 33725,
      //       38913, 33011, 33227, 33408, 38664, 32727, 32794, 33012, 33278,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12243,
      //     c_id_produk: [
      //       32728, 33228, 33282, 33411, 33412, 33508, 33159, 33281, 33619,
      //       32794, 32887, 33283, 32811, 33014, 32812, 33507, 32659, 33013,
      //       33410, 33509, 33625, 33730, 33691, 33704, 39816, 40318,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12249,
      //     c_id_produk: [
      //       32660, 32729, 40260, 33229, 33286, 33413, 32888, 33284, 33707,
      //       38991, 33511, 33692, 32813, 33160, 33734, 38436, 38673, 40279,
      //       32795, 33015, 33706, 33016, 33510, 32814, 33285, 33414, 33619,
      //       33626, 40513,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12251,
      //     c_id_produk: [
      //       33290, 33418, 33419, 32818, 33019, 33619, 33291, 33516, 40516,
      //       32731, 33230, 33510, 39288, 33162, 33724, 39316, 32662, 33020,
      //       32795, 32817, 32890, 33292, 33515, 33628, 33692, 33705, 40296,
      //       40308,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12259,
      //     c_id_produk: [
      //       33510, 33628, 32817, 32890, 33020, 32662, 33230, 33291, 33292,
      //       33515, 33619, 33692, 39288, 40296, 32818, 33290, 33418, 33516,
      //       32795, 33162, 33019, 33419, 33705,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12265,
      //     c_id_produk: [
      //       33160, 33284, 33511, 38436, 40260, 33510, 38673, 40279, 32814,
      //       32888, 33286, 33706, 33707, 33731, 38990, 32660, 33414, 33619,
      //       33015, 33285, 32813, 33229, 33413, 33626, 33692, 32795, 33016,
      //       32729,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12266,
      //     c_id_produk: [
      //       33230, 33419, 33619, 33705, 33721, 39822, 32795, 33510, 40292,
      //       32662, 32818, 33020, 33290, 33292, 33418, 33731, 33628, 39288,
      //       32731, 33692, 38944, 32890, 33019, 33162, 33291, 33516, 40296,
      //       40309, 40567, 32817, 33515,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12267,
      //     c_id_produk: [
      //       32888, 33511, 33692, 38436, 38990, 40260, 33619, 32795, 32813,
      //       33015, 33016, 33626, 32660, 33160, 33706, 33229, 33414, 33510,
      //       33707, 33731, 38673, 32729, 33284, 32814, 33285, 33286, 40279,
      //       33413,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12268,
      //     c_id_produk: [
      //       39288, 32662, 33290, 33291, 33628, 33619, 33721, 38944, 32731,
      //       32795, 33019, 33230, 33419, 33516, 32818, 33162, 33510, 32890,
      //       33020, 33692, 40567, 32817, 33292, 33705, 40296, 33418, 33515,
      //       38967, 39316,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12276,
      //     c_id_produk: [
      //       33017, 33627, 32795, 33287, 33416, 33513, 32815, 33705, 33417,
      //       33161, 33514, 33619, 32889, 33288, 33289, 33415, 38432, 32816,
      //       33018, 33512, 33692, 32661, 33232,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12280,
      //     c_id_produk: [
      //       33512, 33514, 33692, 38689, 39001, 32661, 32889, 33161, 33289,
      //       33417, 33619, 38432, 32816, 33287, 33288, 32730, 32795, 33018,
      //       33415, 33017, 33705, 32815, 33238, 33416, 33513, 33627, 33735,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12284,
      //     c_id_produk: [
      //       32891, 33022, 33291, 32663, 33292, 33705, 33629, 32732, 32820,
      //       33420, 33518, 33741, 32819, 33293, 33517, 33231, 33519, 33619,
      //       33692, 32795, 33021, 33163, 33419,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12287,
      //     c_id_produk: [
      //       33692, 33629, 32891, 33021, 33022, 33743, 32663, 32795, 39288,
      //       33231, 33619, 33291, 33292, 33419, 33518, 32819, 32732, 33163,
      //       33293, 33420, 33517, 33705, 40296,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12288,
      //     c_id_produk: [
      //       32663, 32732, 33021, 40320, 33518, 33692, 33163, 33519, 32819,
      //       33517, 33629, 39288, 32891, 33419, 32795, 33022, 33291, 33293,
      //       39301, 40296, 33292, 33420, 33705, 33743, 32820, 33231, 33619,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12289,
      //     c_id_produk: [
      //       32820, 33163, 33421, 33518, 33619, 33629, 32795, 33517, 33419,
      //       33705, 39273, 40319, 32891, 33022, 32663, 33292, 33293, 32819,
      //       33231, 33739, 39301, 32732, 33021, 33291, 33420, 33519, 33692,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12291,
      //     c_id_produk: [
      //       33424, 33529, 33695, 33708, 39841, 40293, 32647, 32665, 33296,
      //       38468, 33520, 33619, 33631, 33026, 33122, 33123, 33234, 33298,
      //       33423, 39122, 39165, 39827, 32952, 33805, 32823, 32893, 33166,
      //       33523, 39037, 39170, 32824, 33025, 33749, 32734, 32796, 33528,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12295,
      //     c_id_produk: [
      //       33708, 38444, 32796, 33024, 33522, 39260, 33423, 33619, 33520,
      //       33630, 32664, 32892, 32953, 33023, 33294, 33422, 33521, 32647,
      //       32733, 32821, 32822, 33233, 33295, 33296, 33523, 33695, 38733,
      //       39834, 33164, 33747, 39013,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12296,
      //     c_id_produk: [
      //       33631, 33708, 32953, 33234, 33523, 39171, 39827, 40293, 33166,
      //       33423, 33424, 32796, 33528, 33298, 33520, 33750, 32823, 32893,
      //       33619, 32665, 33025, 33026, 38468, 32647, 32734, 32824, 33296,
      //       33529, 33695, 39040, 39841,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12299,
      //     c_id_produk: [
      //       33295, 33520, 32979, 33023, 33296, 33423, 32796, 33024, 33233,
      //       33294, 33695, 39834, 33521, 33522, 33619, 33708, 32664, 32892,
      //       33164, 33630, 38444, 32822, 33422, 32821,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12303,
      //     c_id_produk: [
      //       32647, 33521, 33695, 33708, 39752, 32821, 33522, 33630, 39834,
      //       32892, 33233, 33164, 33295, 33523, 38448, 33024, 33619, 39753,
      //       40265, 32822, 33023, 33294, 33296, 33422, 32796, 32951, 33423,
      //       38444, 32664, 33520,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12320,
      //     c_id_produk: [
      //       32796, 33331, 33619, 33648, 33694, 40311, 33520, 38448, 40283,
      //       32750, 33296, 33710, 40244, 32682, 32909, 33237, 40249, 40264,
      //       32971, 33332, 32826, 33545, 40263, 33042, 33183, 33440, 33767,
      //       38444, 40310,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12322,
      //     c_id_produk: [
      //       33185, 33296, 32973, 33044, 33335, 33769, 40284, 32752, 32911,
      //       33694, 40245, 40266, 40312, 32828, 33442, 33546, 40313, 33237,
      //       40250, 32684, 33336, 33619, 33710, 40241, 32796, 33520, 33650,
      //       40267,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12330,
      //     c_id_produk: [
      //       33520, 33521, 33023, 33164, 33233, 33708, 32821, 33295, 33423,
      //       38516, 39110, 39156, 39240, 39834, 32796, 33296, 33422, 33522,
      //       33695, 38444, 32951, 33294, 32733, 33024, 33523, 32647, 32822,
      //       32892, 33630, 38733, 32664, 33619, 33745, 38737,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12331,
      //     c_id_produk: [
      //       32796, 32824, 32951, 33296, 32665, 32734, 33025, 33234, 33619,
      //       33708, 33773, 39184, 32893, 33298, 33520, 33528, 33529, 40293,
      //       33424, 33695, 33166, 33423, 38468, 39049, 39841, 33026, 33523,
      //       33631, 39827, 32647, 39838, 40304, 32823,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12334,
      //     c_id_produk: [
      //       32796, 33025, 33026, 33523, 33775, 39185, 33296, 33424, 33528,
      //       32823, 32893, 33234, 39053, 39827, 32665, 32734, 33166, 33298,
      //       33619, 33631, 33708, 32950, 33529, 38468, 32647, 33423, 33695,
      //       39838, 32824, 39841, 33520,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12340,
      //     c_id_produk: [
      //       33345, 33655, 39857, 33447, 38748, 32758, 33049, 33190, 33446,
      //       33695, 38481, 38525, 32649, 33050, 33619, 33777, 33784, 32916,
      //       33346, 33449, 33551, 32832, 32833, 33520, 33558, 33559, 33708,
      //       33782, 32981, 33238, 33550, 38736, 39132, 39205, 32689, 32796,
      //       33129, 33347,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12341,
      //     c_id_produk: [
      //       33448, 33552, 33708, 39887, 32690, 32834, 33051, 33557, 39861,
      //       39863, 32649, 32917, 33449, 33520, 38488, 40285, 32759, 32984,
      //       32796, 33656, 32835, 33239, 33348, 33695, 39862, 33192, 33347,
      //       33783, 33052, 33346, 33450, 33556, 33619,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12388,
      //     c_id_produk: [
      //       33469, 32709, 32854, 33369, 33526, 33578, 33711, 40256, 33258,
      //       33368, 40586, 32651, 33165, 33525, 33527, 33619, 33675, 32797,
      //       32934, 33367, 33071, 33696, 40253, 32778, 33005, 40569, 40571,
      //       33470, 40242, 40302, 40582,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12393,
      //     c_id_produk: [
      //       33191, 33370, 33696, 32779, 32797, 33006, 40584, 32855, 33371,
      //       33553, 33711, 40256, 40583, 32710, 33368, 33472, 33555, 40568,
      //       32935, 33072, 32652, 33259, 33471, 33579, 33676, 40302, 40572,
      //       33554, 33619,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12709,
      //     c_id_produk: [
      //       32817, 33418, 33692, 32890, 33290, 33291, 33292, 33516, 32662,
      //       33020, 33019, 33419, 33705, 32795, 33162, 33510, 33515, 33619,
      //       33628, 32818, 33230,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12726,
      //     c_id_produk: [
      //       32795, 33419, 33628, 33418, 33516, 37947, 33019, 33162, 38011,
      //       32662, 32731, 33020, 33291, 32817, 33230, 33518, 33692, 33708,
      //       32818, 32890, 33290, 33292, 33515,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12737,
      //     c_id_produk: [
      //       33233, 33522, 33630, 32822, 33422, 32796, 33164, 33295, 33423,
      //       33294, 33297, 33809, 32664, 33024, 33521, 33695, 32647, 32733,
      //       32821, 32951, 33520, 33523, 33619, 32892, 33023, 33708,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12740,
      //     c_id_produk: [
      //       32727, 33412, 33704, 33011, 33280, 33283, 32658, 33624, 33691,
      //       32794, 33012, 33227, 33505, 33506, 33282, 33619, 32809, 32810,
      //       32886, 33158, 33508, 33725, 33408,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12745,
      //     c_id_produk: [
      //       32709, 33469, 33527, 33578, 38330, 33167, 33258, 33368, 33675,
      //       32797, 33470, 33071, 33369, 33525, 33696, 33711, 32934, 32651,
      //       32778, 33526, 33619, 38328, 32854, 33367,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12746,
      //     c_id_produk: [
      //       32797, 33471, 33579, 33696, 33619, 38331, 32779, 33193, 33676,
      //       33368, 33370, 33371, 32652, 32710, 33554, 38332, 32855, 33472,
      //       33555, 33711, 33072, 33553, 32935, 33259,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12781,
      //     c_id_produk: [32933, 33708, 33520, 33809, 32708, 33297, 39570],
      //   },
      //   {
      //     c_id_produk_mix: 12815,
      //     c_id_produk: [
      //       33422, 33423, 33294, 32664, 33023, 33297, 39712, 33164, 33024,
      //       33520, 33809,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12819,
      //     c_id_produk: [33024, 32664, 33745, 33023, 39715, 33520, 33296, 33423],
      //   },
      //   {
      //     c_id_produk_mix: 12823,
      //     c_id_produk: [
      //       32854, 32934, 33526, 33527, 33696, 32709, 33525, 33619, 32651,
      //       32778, 33165, 33258, 33469, 33578, 33675, 33005, 33809, 32797,
      //       33367, 33368, 33369, 33711, 33071, 33470,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12877,
      //     c_id_produk: [32660, 33286, 39768, 33413, 33414, 33510, 38419],
      //   },
      //   {
      //     c_id_produk_mix: 12934,
      //     c_id_produk: [
      //       33023, 32951, 33024, 33294, 33423, 33708, 38444, 39842, 40188,
      //       32664, 32796, 33164, 33809, 33630, 33695, 32892, 33233, 33297,
      //       33520, 33422, 33619,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12953,
      //     c_id_produk: [
      //       32892, 33233, 33522, 33695, 38444, 39106, 33023, 33295, 33422,
      //       33521, 32664, 32796, 33423, 33520, 33619, 38733, 33745, 33164,
      //       33630, 33708, 32733, 32821, 32822, 33024, 33294, 33297, 33523,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13036,
      //     c_id_produk: [33408, 40618, 32886, 33279, 33407, 39815, 33503, 33704],
      //   },
      //   {
      //     c_id_produk_mix: 13037,
      //     c_id_produk: [32888, 33692, 40619, 33413, 33286, 33414, 33510],
      //   },
      //   {
      //     c_id_produk_mix: 12348,
      //     c_id_produk: [
      //       33551, 32832, 33049, 33346, 32689, 33050, 33447, 33552, 33190,
      //       33619, 33238, 33347, 33695, 40168, 40287, 32649, 32796, 32833,
      //       32979, 33345, 33446, 33550, 33655, 33708, 32916, 33449, 33520,
      //       38484, 38525, 40248, 40285,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 12725,
      //     c_id_produk: [
      //       32817, 32890, 33020, 33162, 33628, 32731, 33292, 33692, 33291,
      //       33230, 33516, 33619, 32662, 32795, 33290, 33418, 33419, 33515,
      //       33518, 37947, 33019, 33705, 32818,
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13018,
      //     c_id_produk: [33422, 33520, 40594, 33297, 33423],
      //   },
      //   { c_id_produk_mix: 11729, c_id_produk: [31796, 23524] },
      //   { c_id_produk_mix: 12981, c_id_produk: [33423, 33520, 40361, 33297] },
      //   { c_id_produk_mix: 13012, c_id_produk: [33520, 33297, 33423, 40566] },
      //   {
      //     c_id_produk_mix: 12814,
      //     c_id_produk: [33520, 39711, 33422, 33423, 33297, 33809],
      //   },
      //   { c_id_produk_mix: 12903, c_id_produk: [33423, 33520, 33297, 39792] },
      // ];

      // const produk_mix = [
      //   {
      //     c_id_produk_mix: 13073,
      //     c_id_produk: [
      //       40708
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13074,
      //     c_id_produk: [
      //       40709
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13077,
      //     c_id_produk: [
      //       40710
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13125,
      //     c_id_produk: [
      //       40832
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13126,
      //     c_id_produk: [
      //       40833
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13127,
      //     c_id_produk: [
      //       40834
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13128,
      //     c_id_produk: [
      //       40835
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13129,
      //     c_id_produk: [
      //       40836
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13130,
      //     c_id_produk: [
      //       40837
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13131,
      //     c_id_produk: [
      //       40838
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13141,
      //     c_id_produk: [
      //       40848
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13142,
      //     c_id_produk: [
      //       40849
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13143,
      //     c_id_produk: [
      //       40850
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13152,
      //     c_id_produk: [
      //       40859
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13153,
      //     c_id_produk: [
      //       40860
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13154,
      //     c_id_produk: [
      //       40861
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13155,
      //     c_id_produk: [
      //       40862
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13160,
      //     c_id_produk: [
      //       40867
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13161,
      //     c_id_produk: [
      //       40868
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13162,
      //     c_id_produk: [
      //       40869
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13196,
      //     c_id_produk: [
      //       40897
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13197,
      //     c_id_produk: [
      //       40898
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13198,
      //     c_id_produk: [
      //       40899
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13216,
      //     c_id_produk: [
      //       40917
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13217,
      //     c_id_produk: [
      //       40918
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 13218,
      //     c_id_produk: [
      //       40919
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 14263,
      //     c_id_produk: [
      //       41963
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 14264,
      //     c_id_produk: [
      //       41964
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 14265,
      //     c_id_produk: [
      //       41965
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 14266,
      //     c_id_produk: [
      //       41966
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 15282,
      //     c_id_produk: [
      //       42471
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 15283,
      //     c_id_produk: [
      //       42472
      //     ],
      //   },
      //   {
      //     c_id_produk_mix: 15284,
      //     c_id_produk: [
      //       42473
      //     ],
      //   },
      // ];

      const result = data.flatMap((pembelian) => {
        const produkMixData = produk_mix.find(
          (mix) => mix.c_id_produk_mix === pembelian.id_produk_mix,
        );
        return produkMixData
          ? produkMixData.c_id_produk.map((id_produk) => ({
              c_no_register: pembelian.c_no_register,
              c_id_produk: id_produk,
            }))
          : [];
      });

      const workbookBaru = new ExcelJS.Workbook();
      const worksheetBaru = workbookBaru.addWorksheet('Data_coba');

      // Menentukan header kolom
      worksheetBaru.columns = [
        { header: 'c_no_register', key: 'c_no_register' },
        { header: 'c_id_produk', key: 'c_id_produk' },
        { header: 'c_status', key: 'c_status' },
        { header: 'c_tanggal_awal', key: 'c_tanggal_awal' },
        { header: 'c_tanggal_akhir', key: 'c_tanggal_akhir' },
      ];

      // console.log(worksheet_baru)

      result.forEach((items) => {
        worksheetBaru.addRow({
          c_no_register: items.c_no_register,
          c_id_produk: items.c_id_produk,
          c_status: 'Aktif',
          c_tanggal_awal: '2023-11-03',
          c_tanggal_akhir: '2023-11-12',
        });
      });

      const filePathBaru = `t_produk_aktif_${file.filename}.xlsx`;
      await workbookBaru.xlsx.writeFile(filePathBaru);

      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }
      return filePathBaru;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async uploadT_produkAktif_find(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const datanya = data.map((item) => {
        return {
          c_id_pembelian: +item.c_id_pembelian,
          c_no_register: +item.c_no_register,
          c_tanggal_daftar: item.c_tanggal_daftar,
          c_id_kelas: +item.c_id_kelas,
          c_tahun_ajaran: item.c_tahun_ajaran,
          c_id_dikdasken: +item.c_id_dikdasken,
          c_nama_lengkap: item.c_nama_lengkap,
          c_id_gedung: +item.c_id_gedung,
          c_id_komar: +item.c_id_komar,
          c_id_kota: +item.c_id_kota,
          c_id_sekolah: +item.c_id_sekolah,
          c_id_sekolah_kelas: +item.c_id_sekolah_kelas,
          c_tingkat_sekolah_kelas: item.c_tingkat_sekolah_kelas,
          c_id_jenis_kelas: +item.c_id_jenis_kelas,
          c_kapasitas_max: +item.c_kapasitas_max,
          c_status_bayar: item.c_status_bayar,
          c_id_bundling: +item.c_id_bundling,
          c_kerjasama: item.c_kerjasama,
        };
      });

      const datanoreg = datanya.map((item) => item.c_no_register);

      const find = await this.prisma.t_produk_aktif.findMany({
        where: {
          c_no_register: {
            in: datanoreg,
          },
        },
      });

      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }

      return find;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async uploadT_produkAktif_delete(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const datanya = data.map((item) => {
        return {
          c_id_pembelian: +item.c_id_pembelian,
          c_no_register: +item.c_no_register,
          c_tanggal_daftar: item.c_tanggal_daftar,
          c_id_kelas: +item.c_id_kelas,
          c_tahun_ajaran: item.c_tahun_ajaran,
          c_id_dikdasken: +item.c_id_dikdasken,
          c_nama_lengkap: item.c_nama_lengkap,
          c_id_gedung: +item.c_id_gedung,
          c_id_komar: +item.c_id_komar,
          c_id_kota: +item.c_id_kota,
          c_id_sekolah: +item.c_id_sekolah,
          c_id_sekolah_kelas: +item.c_id_sekolah_kelas,
          c_tingkat_sekolah_kelas: item.c_tingkat_sekolah_kelas,
          c_id_jenis_kelas: +item.c_id_jenis_kelas,
          c_kapasitas_max: +item.c_kapasitas_max,
          c_status_bayar: item.c_status_bayar,
          c_id_bundling: +item.c_id_bundling,
          c_kerjasama: item.c_kerjasama,
        };
      });

      const datanoreg = datanya.map((item) => item.c_no_register);

      const find = await this.prisma.t_produk_aktif.deleteMany({
        where: {
          c_no_register: {
            in: datanoreg,
          },
        },
      });

      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath); // Menghapus file jika ada
        console.log('File berhasil dihapus.');
      } else {
        console.log('File tidak ditemukan, tidak ada yang dihapus.');
      }

      return find;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async bacaExcel_table_t_produk_aktif(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const c_id = await this.prisma
        .$queryRaw`SELECT MAX(c_id) from t_produk_aktif`;

      const datanya = data.map((item, index) => {
        const c_idnya = c_id[0].max;
        return {
          c_no_register: +item.c_no_register,
          c_id_produk: +item.c_id_produk,
          c_status: item.c_status,
          c_tanggal_awal: new Date(item.c_tanggal_awal),
          c_tanggal_akhir: new Date(item.c_tanggal_akhir),
          c_id: c_idnya + (index + 1),
        };
      });

      return datanya;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async bacaExcel_insert_t_produk_aktifnya(file: any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..', 'excel', file.filename);

      if (!fs.existsSync(filePath)) {
        throw new Error('File not found');
      }
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Assuming you want to read the first worksheet.

      const data = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          // Header row, you can process the headers here.
        } else {
          // Data rows
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            // Map column headers to keys
            const columnHeader: any = worksheet
              .getRow(1)
              .getCell(colNumber).value;
            rowData[columnHeader] = cell.value;
          });
          data.push(rowData);
        }
      });

      const c_id = await this.prisma
        .$queryRaw`SELECT MAX(c_id) from t_produk_aktif`;

      const datanya = data.map((item, index) => {
        const c_idnya = c_id[0].max;
        return {
          c_no_register: +item.c_no_register,
          c_id_produk: +item.c_id_produk,
          c_status: item.c_status,
          c_tanggal_awal: new Date(item.c_tanggal_awal),
          c_tanggal_akhir: new Date(item.c_tanggal_akhir),
          c_id: c_idnya + (index + 1),
        };
      });

      const insertData = await this.prisma.t_produk_aktif.createMany({
        data: datanya,
        skipDuplicates: true,
      });

      return insertData;
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }

  async bcrypt_saja(password){
    try {
      const saltRounds = 10
      const hashedPassword = await bcrypt.hash(
        password,
        saltRounds,
      );
      console.log(hashedPassword)
      return hashedPassword
    } catch (error) {
      console.log(error);
      throw new Error('Error reading Excel file');
    }
  }
}
