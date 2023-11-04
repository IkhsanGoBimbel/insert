import { Body, Controller, Get, HttpException, Post, UploadedFile, UseInterceptors } from '@nestjs/common';
import { AppService } from './app.service';
import { FileInterceptor } from '@nestjs/platform-express';
import { diskStorage } from 'multer';
import { Express } from 'express';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get()
  getHello(): string {
    return this.appService.getHello();
  }

  // @Post('excel')
  // getdataExcel() {
  //   return this.appService.readExcelFile()
  // }

  @Post('excel1')
  getdataExcel() {
    return this.appService.readExcelFiles()
  }



  @Post('insert/pass')
  insertPassword(@Body() body:any) {
    return this.appService.insertpass(+body.username, body.password)
  }

  @Post('user')
  checkpass(@Body() body:any) {
    return this.appService.checkpass(+body.username, body.password)
  }

  @Post('excel/produk-aktif')
  getdataExcelProduk() {
    return this.appService.insertProdukAktif()
  }

  @Post('excel/produk-siswa')
  getdataExcelProdukSiswa() {
    return this.appService.insertProdukSiswa()
  }

  //Excel
  @Post('baca-excel')
  getdataExcelnya() {
    return this.appService.bacaExcel()
  }

  @Post('bcrypt-excel')
  bcryptExcel() {
    return this.appService.bacaExcel_bcrypt()
  }

  @Post('excel-t-siswa')
  findSiswa() {
    return this.appService.bacaExcel_cari_t_siswa()
  }

  @Post('insert-t-produk-siswa')
  find_t_produk_siswa() {
    return this.appService.bacaExcel_insert_t_produk_siswa()
  }

  @Post('cari_produk_siswa')
  find_produk_siswa() {
    return this.appService.bacaExcel_cari_produk_aktif()
  }

  @Post('insert-t-produk-aktif')
  find_t_produk_aktif() {
    return this.appService.bacaExcel_insert_t_produk_aktif()
  }


  @Post('excel/t_siswa/read')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcelt_siswa(@UploadedFile() file){
    return this.appService.uploadbacaExcel(file)
  }

  @Post('excel/t_produk_siswa/read')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcelt_produk_siswa(@UploadedFile() file){
    return this.appService.uploadbacaExcel(file)
  }

  @Post('excel/t_produk_aktif/read')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcelt_produk_aktif(@UploadedFile() file){
    return this.appService.uploadbacaExcel(file)
  }

  @Post('excel/t_siswa/find')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcel_find_t_siswa(@UploadedFile() file){
    return this.appService.uploadbcryptExcelCariT_Siswa(file)
  }

  @Post('excel/t_siswa/insert')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcel_insert_t_siswa(@UploadedFile() file){
    return this.appService.uploadbcryptExcel(file)
  }

  @Post('excel/t_produk_siswa/readtable')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcel_read_table(@UploadedFile() file){
    return this.appService.uploadT_produkSiswaRead(file)
  }

  @Post('excel/t_produk_siswa/insert')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcel_insert_table(@UploadedFile() file){
    return this.appService.uploadT_produkSiswaInsert(file)
  }

  @Post('excel/t_produk_siswa/build_produk_aktif')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcel_build_produk_aktif(@UploadedFile() file){
    return this.appService.uploadT_produkSiswaBuildProdukAktif(file)
  }

  @Post('excel/t_produk_aktif/readtable')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcel_read_table_produk_aktif(@UploadedFile() file){
    return this.appService.bacaExcel_table_t_produk_aktif(file)
  }

  @Post('excel/t_produk_aktif/insert')
  @UseInterceptors(FileInterceptorWithDest('./excel'))
  async readExcel_insert_t_produk_aktif(@UploadedFile() file){
    return this.appService.bacaExcel_insert_t_produk_aktifnya(file)
  }

}

export function FileInterceptorWithDest(destination: string) {
  return FileInterceptor('excel', {
    storage: diskStorage({
      destination: destination,
      filename: (req, file, cb) => {
        const table = req.originalUrl.split('/')
        const date = new Date()
        const tanggal = date.getDate()
        const bulan = date.getMonth()
        const tahun = date.getFullYear()
        const jam = date.getHours()
        const menit = date.getMinutes()
        const uniqueSuffix = file.originalname;
        return cb(null, file.fieldname + '-' + table[2]+'-'+ tanggal + '-' + bulan + '-' + tahun + '-' + jam + '-' + menit + '-' + uniqueSuffix);
      },
    }),
    fileFilter: (req, file, cb) => {
      if (!file.originalname.match(/\.(xlsx)$/)) {
        return cb(new HttpException('Invalid file type', 403), false);
      }
      cb(null, true);
    },
  });
}

