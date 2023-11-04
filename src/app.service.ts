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
      console.log(error)
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
      const filePath = path.join(__dirname, '..', 'data t_produk_siswa_Dummy.xlsx');

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
          c_id_produk: [
            40870
          ],
        },
        {
          c_id_produk_mix: 13164,
          c_id_produk: [
            40871
          ],
        },
        {
          c_id_produk_mix: 13165,
          c_id_produk: [
            40872
          ],
        },
        {
          c_id_produk_mix: 13166,
          c_id_produk: [
            40873
          ],
        },
        {
          c_id_produk_mix: 13182,
          c_id_produk: [
           40866
          ],
        },
        {
          c_id_produk_mix: 13183,
          c_id_produk: [
           40867
          ],
        },
        {
          c_id_produk_mix: 13184,
          c_id_produk: [
           40868
          ],
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
        const c_idnya = 6002670
        return {
          c_no_register: +item.c_no_register,
          c_id_produk : +item.c_id_produk,
          c_status : item.c_status,
          c_tanggal_awal : new Date(item.c_tanggal_awal),
          c_tanggal_akhir : new Date(item.c_tanggal_akhir),
          c_id : c_idnya + (index+1)
        };
      });

      // return datanya
      const insertData = await this.prisma.t_produk_aktif.createMany({
        data: datanya,
        skipDuplicates: true
      });

      return insertData
    }
    catch (error){
      console.log(error)
      throw new Error('Error reading Excel file');
    }
  }

  async uploadbacaExcel(file:any) {
    const workbook = new ExcelJS.Workbook();
    try {
      const filePath = path.join(__dirname, '..','excel', file.filename);

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

}
