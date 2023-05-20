<?php
include('koneksi.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Cek apakah ada data yang dikirimkan melalui formulir
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Mendapatkan data dari formulir
    $nama = $_POST['nama'];
    $jenis_kelamin = $_POST['jenis_kelamin'];
    $tempat_lahir = $_POST['tempat_lahir'];
    $tanggal_lahir = $_POST['tanggal_lahir'];
    $agama = $_POST['agama'];
    $nik = $_POST['nik'];
    $no_hp = $_POST['no_hp'];
    $email = $_POST['email'];
    $alamat = $_POST['alamat'];

    $host = "localhost";
    $user = "root";
    $password = "";
    $database = "db_siswa";

    $koneksi = mysqli_connect($host, $user, $password, $database);

    $sql = "INSERT INTO tb_siswabaru (nama, jenis_kelamin, tempat_lahir, tanggal_lahir, agama, nik, no_hp, email, alamat)
            VALUES ('$nama', '$jenis_kelamin', '$tempat_lahir', '$tanggal_lahir', '$agama', '$nik', '$no_hp', '$email', '$alamat')";

    mysqli_query($koneksi, $sql);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', 'No');
    $sheet->setCellValue('B1', 'Nama');
    $sheet->setCellValue('C1', 'Jenis Kelamin');
    $sheet->setCellValue('D1', 'Tempat Lahir');
    $sheet->setCellValue('E1', 'Tanggal Lahir');
    $sheet->setCellValue('F1', 'Agama');
    $sheet->setCellValue('G1', 'NIK');
    $sheet->setCellValue('H1', 'No HP');
    $sheet->setCellValue('I1', 'Email');
    $sheet->setCellValue('J1', 'Alamat');

    $query = mysqli_query($koneksi, "SELECT * FROM tb_siswabaru");
    $i = 2;
    $no = 1;
    while ($row = mysqli_fetch_array($query)) {
        $sheet->setCellValue('A' . $i, $no);
        $sheet->setCellValue('B' . $i, $row['nama']);
        $sheet->setCellValue('C' . $i, $row['jenis_kelamin']);
        $sheet->setCellValue('D' . $i, $row['tempat_lahir']);
        $sheet->setCellValue('E' . $i, $row['tanggal_lahir']);
        $sheet->setCellValue('F' . $i, $row['agama']);
        $sheet->setCellValue('G' . $i, $row['nik']);
        $sheet->setCellValue('H' . $i, $row['no_hp']);
        $sheet->setCellValue('I' . $i, $row['email']);
        $sheet->setCellValue('J' . $i, $row['alamat']);
        $i++;
        $no++;
    }

    $styleArray = [
        'borders' => [
            'allBorders' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            ],
        ],
    ];

    // Menyimpan file Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save('Pendaftaran Siswa.xlsx');
}
?>
<!doctype html>
<html lang="en">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Form Siswa</title>
    <!-- css -->
    <link rel="stylesheet" href="style.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-KK94CHFLLe+nY2dmCWGMq91rCGa5gtU4mk92HdvYe+M/SXH301p5ILy+dN9+nJOZ" crossorigin="anonymous">
</head>

<body>

    <!-- form start -->
    <form class="form" method="post" action="" style="padding-top: 2rem;">

        <div class="judulform">
            <h5>Formulir Pendaftaran Siswa Baru</h5>
        </div>

        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="nama" class="col-form-label">1. Nama Lengkap</label>
            </div>
            <div class="col">
                <input type="text" id="nama" name="nama" class="form-control" required>
            </div>
        </div>
        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="jenis_kelamin" class="col-form-label">2. Jenis Kelamin</label>
            </div>
            <div class="col">
                <div style="display: flex;">
                    <div style="padding-left: 8px;">
                        <input class="radiobtn" type="radio" name="jenis_kelamin" value="L"> Laki-Laki
                    </div>
                    <div style="padding-left: 10rem;"> 
                        <input class="radiobtn" type="radio" name="jenis_kelamin" value="P"> Perempuan 
                    </div>
                </div>
            </div>
        </div>

        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="tempat_lahir" class="col-form-label">3. Tempat Lahir</label>
            </div>
            <div class="col">
                <input type="text" id="tempat_lahir" name="tempat_lahir" class="form-control" required>
            </div>
        </div>

        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="tanggal_lahir" class="col-form-label">4. Tanggal Lahir</label>
            </div>
            <div class="col">
                <input type="date" class="form-control" id="tanggal_lahir" name="tanggal_lahir" required>
            </div>
        </div>

        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="agama" class="col-form-label">5. Agama</label>
            </div>
            <div class="col">
                <input type="text" id="agama" name="agama" class="form-control" required>
            </div>
        </div>

        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="nik" class="col-form-label">6. NIK </label>
            </div>
            <div class="col">
                <input type="text" id="nik" name="nik" class="form-control" required>
            </div>
        </div>

        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="no_hp" class="col-form-label">7. No HP</label>
            </div>
            <div class="col">
                <input type="text" id="no_hp" name="no_hp" class="form-control" required>
            </div>
        </div>

        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="email" class="col-form-label">8. Email</label>
            </div>
            <div class="col">
                <input type="text" id="email" name="email" class="form-control" required>
            </div>
        </div>

        <div class="row g-3 align-items-center">
            <div class="col-3">
                <label for="alamat" class="col-form-label">9. Alamat</label>
            </div>
            <div class="col">
                <input type="text" id="alamat" name="alamat" class="form-control" required>
            </div>
        </div>

        <input type="submit" value="Submit" class="btn btn-primary">
        <input type="submit" value="Export To Excel" class="btn btn-warning btnreport">
    </form>

    <!-- form end -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ENjdO4Dr2bkBIFxQpeoTz1HIcje39Wm4jDKdf19U8gI4ddQ3GYNS7NTKfAdVQSZe" crossorigin="anonymous"></script>
</body>

</html>