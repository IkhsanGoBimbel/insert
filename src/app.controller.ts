import { Body, Controller, Get, Post } from '@nestjs/common';
import { AppService } from './app.service';

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


}
