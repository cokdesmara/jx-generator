<?php
require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Shared\Date as XlsDate;
use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

// Read filter: only load INSERT FORM sheet
class InsertFormFilter implements IReadFilter
{
    public function readCell($columnAddress, $row, $worksheetName = ''): bool
    {
        return $worksheetName === 'INSERT FORM';
    }
}

header('Content-Type: application/json');
if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['error' => 'Gunakan method POST']);
    exit;
}
if (empty($_FILES['acuan'])) {
    http_response_code(400);
    echo json_encode(['error' => 'File acuan belum dipilih']);
    exit;
}
$f = $_FILES['acuan'];
if ($f['error'] !== UPLOAD_ERR_OK) {
    http_response_code(400);
    echo json_encode(['error' => 'Upload gagal (kode ' . $f['error'] . ')']);
    exit;
}

$INSROWS = [15,19,25,29,96,97,98,99,100,101,127,128,129,130,131,132,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,202,204,205,206,207,208,209,210,211,213,215,216,217,219,221,222,223];

// Hanya baca nilai cache (oldCalculatedValue) - jauh lebih cepat, tidak evaluasi formula
function cellCached($ws, $r, $c)
{
    try {
        $cell = $ws->getCell(Coordinate::stringFromColumnIndex($c) . $r);
        $v = $cell->getOldCalculatedValue();
        if ($v === null || $v === '') {
            // Fallback ke value mentah (bukan formula)
            $v = $cell->getValue();
        }
        return $v;
    } catch (\Throwable $e) {
        return null;
    }
}

function cellDateCached($ws, $r, $c)
{
    $v = cellCached($ws, $r, $c);
    if ($v instanceof \DateTimeInterface) {
        return $v->format('Y-m-d');
    }
    if (is_numeric($v)) {
        try {
            return XlsDate::excelToDateTimeObject((float)$v)->format('Y-m-d');
        } catch (\Throwable $e) {
        }
    }
    return null;
}

try {
    $reader = IOFactory::createReaderForFile($f['tmp_name']);
    $reader->setReadFilter(new InsertFormFilter());
    $reader->setReadDataOnly(true);
    $ss = $reader->load($f['tmp_name']);
    if (null === $ss->getSheetByName('INSERT FORM')) {
        throw new Exception('Sheet "INSERT FORM" tidak ditemukan ??? file ini bukan Meal Charge?');
    }
    $wsI = $ss->getSheetByName('INSERT FORM');

    $dates = [];
    for ($c = 3; $c <= 36; $c++) {
        $d = cellDateCached($wsI, 9, $c);
        if ($d !== null) {
            $has = false;
            foreach ($INSROWS as $r) {
                $val = cellCached($wsI, $r, $c);
                if ($val !== null && $val !== '') {
                    $has = true;
                    break;
                }
            }
            $dates[] = ['date' => $d, 'has' => $has];
        }
    }
    if (!$dates) {
        throw new Exception('Tidak ada tanggal di INSERT FORM baris 9.');
    }
    echo json_encode(['ok' => true, 'dates' => $dates]);
} catch (\Throwable $e) {
    http_response_code(500);
    echo json_encode(['error' => $e->getMessage()]);
}
