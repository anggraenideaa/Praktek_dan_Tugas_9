<?php
include 'koneksi.php';
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;



// // Fungsi untuk menyimpan data siswa ke dalam file Excel
// function exportToExcel($data)
// {
//     $spreadsheet = new Spreadsheet();
//     $sheet = $spreadsheet->getActiveSheet();
//     $sheet->setCellValue('A1', 'NPM');
//     $sheet->setCellValue('B1', 'NAMA LENGKAP');
//     $sheet->setCellValue('C1', 'JENIS KELAMIN');
//     $sheet->setCellValue('D1', 'NISN');
//     $sheet->setCellValue('E1', 'NIK');
//     $sheet->setCellValue('F1', 'TANGGAL LAHIR');
//     $sheet->setCellValue('G1', 'TEMPAT LAHIR');
//     $sheet->setCellValue('H1', 'AGAMA');
//     $sheet->setCellValue('I1', 'ALAMAAT JALAN');
//     $sheet->setCellValue('J1', 'RT');
//     $sheet->setCellValue('K1', 'RW');
//     $sheet->setCellValue('L1', 'NAMA DUSUN');
//     $sheet->setCellValue('M1', 'KELURAHAN');
//     $sheet->setCellValue('N1', 'KECAMATAN');
//     $sheet->setCellValue('O1', 'KODE POS');
//     $sheet->setCellValue('P1', 'NOMOR HP');
//     $sheet->setCellValue('Q1', 'EMAIL');

//     // Mendapatkan baris terakhir di file Excel
//     $lastRow = $sheet->getHighestRow();

//     // Menyisipkan data siswa baru ke dalam file Excel
//     $sheet->setCellValue('A' . $lastRow, $data['npm']);
//     $sheet->setCellValue('B' . $lastRow, $data['nama_lengkap']);
//     $sheet->setCellValue('C' . $lastRow, $data['jenis_kelamin']);
//     $sheet->setCellValue('D' . $lastRow, $data['nisn']);
//     $sheet->setCellValue('E' . $lastRow, $data['nik']);
//     $sheet->setCellValue('F' . $lastRow, $data['tanggal_lahir']);
//     $sheet->setCellValue('G' . $lastRow, $data['tempat_lahir']);
//     $sheet->setCellValue('H' . $lastRow, $data['agama']);
//     $sheet->setCellValue('I' . $lastRow, $data['alamat_jalan']);
//     $sheet->setCellValue('J' . $lastRow, $data['rt']);
//     $sheet->setCellValue('K' . $lastRow, $data['rw']);
//     $sheet->setCellValue('L' . $lastRow, $data['nama_dusun']);
//     $sheet->setCellValue('M' . $lastRow, $data['nama_kelurahan']);
//     $sheet->setCellValue('N' . $lastRow, $data['kecamatan']);
//     $sheet->setCellValue('O' . $lastRow, $data['kode_pos']);
//     $sheet->setCellValue('P' . $lastRow, $data['nomor_hp']);
//     $sheet->setCellValue('Q' . $lastRow, $data['email_pribadi']);

//     // Menyimpan file Excel
//     $writer = new Xlsx($spreadsheet);
//     $writer->save('Daftar Siswa.xlsx');
// }

// Cek apakah ada data yang dikirimkan melalui formulir
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Mendapatkan data dari formulir
    $npm = $_POST['npm'];
    $nama_lengkap = $_POST['nama_lengkap'];
    $jenis_kelamin = $_POST['jenis_kelamin'];
    $nisn = $_POST['nisn'];
    $nik = $_POST['nik'];
    $tanggal_lahir = $_POST['tanggal_lahir'];
    $tempat_lahir = $_POST['tempat_lahir'];
    $agama = $_POST['agama'];
    $alamat_jalan = $_POST['alamat_jalan'];
    $rt = $_POST['rt'];
    $rw = $_POST['rw'];
    $nama_dusun = $_POST['nama_dusun'];
    $nama_kelurahan = $_POST['nama_kelurahan'];
    $kecamatan = $_POST['kecamatan'];
    $kode_pos = $_POST['kode_pos'];
    $nomor_hp = $_POST['nomor_hp'];
    $email_pribadi = $_POST['email_pribadi'];

    $sql = "INSERT INTO data_peserta_didik (npm, nama_lengkap, jenis_kelamin, nisn, nik, tanggal_lahir, tempat_lahir, agama, alamat_jalan, rt, rw, nama_dusun, kelurahan, kecamatan, kode_pos, nomor_hp, email)
VALUES ('$npm', '$nama_lengkap', '$jenis_kelamin', '$nisn', '$nik', '$tanggal_lahir', '$tempat_lahir', '$agama', '$alamat_jalan', '$rt', '$rw', '$nama_dusun', '$nama_kelurahan', '$kecamatan', '$kode_pos', '$nomor_hp', '$email_pribadi')";

    mysqli_query($koneksi, $sql);

    $host = "localhost";
    $user = "root";
    $password = "";
    $database = "db_siswa";

    $koneksi = mysqli_connect($host, $user, $password, $database);


    $spreadsheet = new Spreadsheet();
    $sheet          = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', 'NPM');
    $sheet->setCellValue('B1', 'NAMA LENGKAP');
    $sheet->setCellValue('C1', 'JENIS KELAMIN');
    $sheet->setCellValue('D1', 'NISN');
    $sheet->setCellValue('E1', 'NIK');
    $sheet->setCellValue('F1', 'TANGGAL LAHIR');
    $sheet->setCellValue('G1', 'TEMPAT LAHIR');
    $sheet->setCellValue('H1', 'AGAMA');
    $sheet->setCellValue('I1', 'ALAMAAT JALAN');
    $sheet->setCellValue('J1', 'RT');
    $sheet->setCellValue('K1', 'RW');
    $sheet->setCellValue('L1', 'NAMA DUSUN');
    $sheet->setCellValue('M1', 'KELURAHAN');
    $sheet->setCellValue('N1', 'KECAMATAN');
    $sheet->setCellValue('O1', 'KODE POS');
    $sheet->setCellValue('P1', 'NOMOR HP');
    $sheet->setCellValue('Q1', 'EMAIL');

    $query = mysqli_query($koneksi, "select * from data_peserta_didik");
    $i = 2;
    $no = 1;
    while ($row = mysqli_fetch_array($query)) {
        $sheet->setCellValue('A' . $i, $row['npm']);
        $sheet->setCellValue('B' . $i, $row['nama_lengkap']);
        $sheet->setCellValue('C' . $i, $row['jenis_kelamin']);
        $sheet->setCellValue('D' . $i, $row['nisn']);
        $sheet->setCellValue('E' . $i, $row['nik']);
        $sheet->setCellValue('F' . $i, $row['tanggal_lahir']);
        $sheet->setCellValue('G' . $i, $row['tempat_lahir']);
        $sheet->setCellValue('H' . $i, $row['agama']);
        $sheet->setCellValue('I' . $i, $row['alamat_jalan']);
        $sheet->setCellValue('J' . $i, $row['rt']);
        $sheet->setCellValue('K' . $i, $row['rw']);
        $sheet->setCellValue('L' . $i, $row['nama_dusun']);
        $sheet->setCellValue('M' . $i, $row['kelurahan']);
        $sheet->setCellValue('N' . $i, $row['kecamatan']);
        $sheet->setCellValue('O' . $i, $row['kode_pos']);
        $sheet->setCellValue('P' . $i, $row['nomor_hp']);
        $sheet->setCellValue('Q' . $i, $row['email']);
        $i++;
    }
    $styleArray = [
        'borders' => [
            'allBorders' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            ],
        ],
    ];

    //     // Menyimpan file Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save('Daftar Siswa.xlsx');
}



?>


<!doctype html>
<html lang="en">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>FORM REGISTRASI</title>
    <!-- css -->
    <link rel="stylesheet" href="style.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-KK94CHFLLe+nY2dmCWGMq91rCGa5gtU4mk92HdvYe+M/SXH301p5ILy+dN9+nJOZ" crossorigin="anonymous">
</head>

<body>

    <!-- form start -->
    <form class="form" method="post" action="" style="padding-top: 2rem;">

        <div class="judulform">
            <h5>Data Pribadi</h5>
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 1. </label>
            <label class="judul-label" for=""> NPM </label>
            <input class="input" type="text" name="npm">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 2. </label>
            <label class="judul-label" for=""> Nama Lengkap </label>
            <input class="input" type="text" name="nama_lengkap">
        </div>
        <div class="label mb-3" style="display: flex;">
            <label for="" class="nomor-label"> 3. </label>
            <label class="judul-label" for="" style="width: 18rem;"> Jenis Kelamin </label>
            <div style="display: flex;">
                <div style="padding-left: 8px;"> <input class="radiobtn" type="radio" name="jenis_kelamin" value="L"> Laki-Laki</div>

                <div style="padding-left: 10rem;"> <input class="radiobtn" type="radio" name="jenis_kelamin" value="P"> Perempuan </div>
            </div>
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 4. </label>
            <label class="judul-label" for=""> NISN </label>
            <input class="input" type="number" name="nisn">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 5. </label>
            <label class="judul-label" for=""> NIK </label>
            <input class="input" type="number" name="nik">
        </div>
        <div class="label mb-3" style="display: flex;">
            <label for="" class="nomor-label"> 6. </label>
            <label class="judul-label" for="" style="width: 18.6rem;"> Tanggal Lahir </label>
            <input type="date" name="tanggal_lahir">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 7. </label>
            <label class="judul-label" for="" style="width: 18rem;"> Tempat Lahir </label>
            <input type="text" name="tempat_lahir">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 8. </label>
            <label class="judul-label" for=""> Agama</label>
            <input class="input" type="text" name="agama" style="width: 6rem ;">
            <label for="">01)Islam</label>
            <label for="">02)Kristem/Protestan</label>
            <label for="">03)Katholik</label>
            <label for="">04)Hindhu</label>
            <label for="">05)Budha</label>
            <label for="">06)Kong Hu Chu</label>
            <label for="">99)Lainnya</label>
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 9. </label>
            <label class="judul-label" for=""> Alamat Jalan </label>
            <input class="input" type="text" name="alamat_jalan">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 10. </label>
            <label class="judul-label" for=""> RT </label>
            <input class="input" type="number" name="rt" style="width: 6rem ;">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 11. </label>
            <label class="judul-label" for=""> RW </label>
            <input class="input" type="number" name="rw" style="width: 6rem ;">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 12. </label>
            <label class="judul-label" for=""> Nama Dusun </label>
            <input class="input" type="text" name="nama_dusun">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 13. </label>
            <label class="judul-label" for=""> Nama Kelurahan/Desa </label>
            <input class="input" type="text" name="nama_kelurahan">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 14. </label>
            <label class="judul-label" for=""> Kecamatan </label>
            <input class="input" type="text" name="kecamatan">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 15. </label>
            <label class="judul-label" for=""> Kode Pos </label>
            <input class="input" type="number" name="kode_pos" style="width: 6rem ;">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 16. </label>
            <label class="judul-label" for=""> Nomor HP </label>
            <input class="input" type="number" name="nomor_hp">
        </div>
        <div class="label mb-3">
            <label for="" class="nomor-label"> 17. </label>
            <label class="judul-label" for=""> Email Pribadi </label>
            <input class="input" type="text" name="email_pribadi">
        </div>
        <input type="submit" value="Simpan">

        <button type="submit" class="btnreport"> Export To Excel</button>
    </form>

    <!-- form end -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ENjdO4Dr2bkBIFxQpeoTz1HIcje39Wm4jDKdf19U8gI4ddQ3GYNS7NTKfAdVQSZe" crossorigin="anonymous"></script>
</body>

</html>