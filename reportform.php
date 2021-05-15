<?php
// Menyambungkan dengan koneksi1.php
include "koneksi1.php";
require 'vendor/autoload.php'; 
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
// Membuat object dengan nama $spreadsheet dengan menggunakan Spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
// Nama kolom
$sheet->setCellValue('A1', 'Jenis Pendaftaran');
$sheet->setCellValue('B1', 'Tanggal Masuk Sekolah');
$sheet->setCellValue('C1', 'NIS');
$sheet->setCellValue('D1', 'Nomor Peserta Ujian');
$sheet->setCellValue('E1', 'Pernah PAUD');
$sheet->setCellValue('F1', 'Pernah TK');
$sheet->setCellValue('G1', 'No. Seri SKHUN Sebelumnya');
$sheet->setCellValue('H1', 'No. Seri Ijazah Sebelumnya');
$sheet->setCellValue('I1', 'Hobi');
$sheet->setCellValue('J1', 'Cita-cita');
$sheet->setCellValue('K1', 'Nama Lengkap');
$sheet->setCellValue('L1', 'Jenis Kelamin');
$sheet->setCellValue('M1', 'NISN');
$sheet->setCellValue('N1', 'NIK');
$sheet->setCellValue('O1', 'Tempat Lahir');
$sheet->setCellValue('P1', 'Tanggal Lahir');
$sheet->setCellValue('Q1', 'Agama');
$sheet->setCellValue('R1', 'Berkebutuhan Khusus');
$sheet->setCellValue('S1', 'Alamat Jalan');
$sheet->setCellValue('T1', 'RT');
$sheet->setCellValue('U1', 'RW');
$sheet->setCellValue('V1', 'Dusun');
$sheet->setCellValue('W1', 'Kelurahan/Desa');
$sheet->setCellValue('X1', 'Kecamatan');
$sheet->setCellValue('Y1', 'Kode Pos');
$sheet->setCellValue('Z1', 'Tempat Tinggal');
$sheet->setCellValue('AA1', 'Moda Transportasi');
$sheet->setCellValue('AB1', 'No. HP');
$sheet->setCellValue('AC1', 'No. Telepon');
$sheet->setCellValue('AD1', 'E-mail Pribadi');
$sheet->setCellValue('AE1', 'Penerima KPS/PKH/KIP');
$sheet->setCellValue('AF1', 'No. KPS/PKH/KIP');
$sheet->setCellValue('AG1', 'Kewarganegaraan');

// Mengambil data pada database dan menuliskan pada excel
$query = mysqli_query($koneksi,"select * from pesertadidik");
$i = 2;
while ($row = mysqli_fetch_array($query)){
	$sheet->setCellValue('A'.$i, $row['jenisdaftar']);
	$sheet->setCellValue('B'.$i, $row['tglmasuk']);
	$sheet->setCellValue('C'.$i, $row['nis']);
	$sheet->setCellValue('D'.$i, $row['noujian']);
	$sheet->setCellValue('E'.$i, $row['paud']);
	$sheet->setCellValue('F'.$i, $row['tk']);
	$sheet->setCellValue('G'.$i, $row['noskhun']);
	$sheet->setCellValue('H'.$i, $row['noijazah']);
	$sheet->setCellValue('I'.$i, $row['hobi']);
	$sheet->setCellValue('J'.$i, $row['citacita']);
    $sheet->setCellValue('K'.$i, $row['namalengkap']);
	$sheet->setCellValue('L'.$i, $row['jeniskelamin']);
	$sheet->setCellValue('M'.$i, $row['nisn']);
	$sheet->setCellValue('N'.$i, $row['nik']);
	$sheet->setCellValue('O'.$i, $row['tempatlahir']);
	$sheet->setCellValue('P'.$i, $row['tgllahir']);
	$sheet->setCellValue('Q'.$i, $row['agama']);
	$sheet->setCellValue('R'.$i, $row['kebutuhankhusus']);
	$sheet->setCellValue('S'.$i, $row['alamatjalan']);
	$sheet->setCellValue('T'.$i, $row['rt']);
	$sheet->setCellValue('U'.$i, $row['rw']);
	$sheet->setCellValue('V'.$i, $row['dusun']);
	$sheet->setCellValue('W'.$i, $row['kelurahan']);
	$sheet->setCellValue('X'.$i, $row['kecamatan']);
	$sheet->setCellValue('Y'.$i, $row['kodepos']);
	$sheet->setCellValue('Z'.$i, $row['tinggal']);
	$sheet->setCellValue('AA'.$i, $row['transportasi']);
	$sheet->setCellValue('AB'.$i, $row['nohp']);
	$sheet->setCellValue('AC'.$i, $row['notelp']);
	$sheet->setCellValue('AD'.$i, $row['email']);
	$sheet->setCellValue('AE'.$i, $row['penerimakps']);
	$sheet->setCellValue('AF'.$i, $row['nokps']);
	$sheet->setCellValue('AG'.$i, $row['kwn']);	
	$i++;
}

$styleArray = [
			'borders' => [
				'allBorders' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
		];
$i = $i - 1;
$sheet->getStyle('A1:Y'.$i)->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Report Pendaftaran Siswa.xlsx'); // Menyimpan file excel dengan nama Report Pendaftaran Siswa Baru.xlsx
?>