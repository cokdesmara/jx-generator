<?php
require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Shared\Date as XlsDate;

// Template checklist (base64) di-embed agar tersedia di environment serverless Vercel.
const TEMPLATE_B64 = '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';

// ---- Konstanta mapping (port dari Generate-Checklist.ps1) ----
$J_FNL = 19; $Y_FNL = 29; $J_SPML = 127; $Y_SPML = 195; $FC_NRML = 202; $CC_NRML = 213;
$FC_SPML = range(204, 211); $CC_SPML = [215, 216, 217];
$DMAP = ["D12"=>9,"D13"=>10,"D14"=>10,"D16"=>11,"D17"=>13,"D18"=>15,"D19"=>17,"D20"=>23,"D21"=>19,"D22"=>24,"D25"=>28,"D26"=>28,"D27"=>30,"D28"=>32,"D29"=>34];
$EMAP = ["E16"=>12,"E17"=>14,"E18"=>16,"E19"=>18,"E21"=>20,"E22"=>25,"E25"=>29,"E26"=>29,"E27"=>31,"E28"=>33,"E29"=>35];
$YDMAP = ["D32"=>14,"D35"=>10,"D36"=>12];
$YEINS = ["E35"=>168,"E36"=>169];
$JBLOCK = [96,97,98,99,100,101];
$Y_SPMLC = [170,171,172,173,174,175,176,177,178,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194];
$BBML_ROW = 196;
$INSROWS = [15,19,25,29,96,97,98,99,100,101,127,128,129,130,131,132,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,202,204,205,206,207,208,209,210,211,213,215,216,217,219,221,222,223];
$JNROWS = [9,10,11,12,13,14,15,16,17,18,19,20,22,23,24,25,28,29,30,31,32,33,34,35,37];

function cellValue($ws, $r, $c)
{
    try {
        $cell = $ws->getCell(Coordinate::stringFromColumnIndex($c) . $r);
    } catch (\Throwable $e) {
        return null;
    }
    // Untuk sheet non-INSERT FORM, gunakan nilai cache saja (lebih cepat)
    // Hanya INSERT FORM yang perlu fallback ke cache karena rumusnya gagal dievaluasi
    $v = null;
    try {
        $v = $cell->getOldCalculatedValue();
    } catch (\Throwable $e) {
    }
    // Jika cache kosong, baru coba calculated value
    if ($v === null || $v === '') {
        try {
            $v = $cell->getCalculatedValue();
        } catch (\Throwable $e) {
            $v = null;
        }
    }
    $isErr = is_string($v) && preg_match('/^#(VALUE|REF|DIV|NAME|NUM|NULL|N\/A)/', $v);
    if ($isErr) {
        $v = null;
    }
    return $v;
}

function cellNum($ws, $r, $c)
{
    $v = cellValue($ws, $r, $c);
    return is_numeric($v) ? (float)$v : null;
}

function cellDate($ws, $r, $c)
{
    $v = cellValue($ws, $r, $c);
    if ($v instanceof \DateTimeInterface) {
        return $v->format('Y-m-d');
    }
    if (is_numeric($v)) {
        try {
            return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject((float)$v)->format('Y-m-d');
        } catch (\Throwable $e) {
        }
    }
    return null;
}

function codeOf($label)
{
    $s = trim((string)$label);
    foreach (['FC ', 'CC ', 'ME '] as $p) {
        if (str_starts_with($s, $p)) {
            $s = trim(substr($s, strlen($p)));
            break;
        }
    }
    if (str_contains($s, '/')) {
        $s = trim(explode('/', $s)[0]);
    }
    return $s;
}

function getLabel($ws, $r, $c)
{
    try {
        $v = $ws->getCell(Coordinate::stringFromColumnIndex($c) . $r)->getCalculatedValue();
        if ($v !== null && $v !== '') return (string)$v;
        $v = $ws->getCell(Coordinate::stringFromColumnIndex($c) . $r)->getOldCalculatedValue();
        if ($v !== null && $v !== '') return (string)$v;
        $v = $ws->getCell(Coordinate::stringFromColumnIndex($c) . $r)->getValue();
        if ($v !== null && $v !== '') return (string)$v;
    } catch (\Throwable $e) {
    }
    return '';
}

// ---- Terima request ----
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
$date = isset($_POST['date']) ? trim($_POST['date']) : '';
if (!preg_match('/^\d{4}-\d{2}-\d{2}$/', $date)) {
    http_response_code(400);
    echo json_encode(['error' => 'Tanggal tidak valid']);
    exit;
}
$f = $_FILES['acuan'];
if ($f['error'] !== UPLOAD_ERR_OK) {
    http_response_code(400);
    echo json_encode(['error' => 'Upload gagal (kode ' . $f['error'] . ')']);
    exit;
}

try {
    $reader = IOFactory::createReaderForFile($f['tmp_name']);
    $reader->setLoadSheetsOnly(['INSERT FORM', 'J NRML', 'Y NRML', 'FCML Tech Crew', 'CCML Cabin Crew', 'Y SPML ']);
    $reader->setReadDataOnly(true);
    $ss = $reader->load($f['tmp_name']);
    if (null === $ss->getSheetByName('INSERT FORM')) {
        throw new Exception('Sheet "INSERT FORM" tidak ditemukan ??? file ini bukan Meal Charge?');
    }
    $wsI = $ss->getSheetByName('INSERT FORM');

    // Cari kolom tanggal
    $col = null;
    for ($c = 3; $c <= 36; $c++) {
        if (cellDate($wsI, 9, $c) === $date) {
            $col = $c;
            break;
        }
    }
    if ($col === null) {
        throw new Exception('Tanggal ' . $date . ' tidak ada di INSERT FORM baris 9.');
    }

    // Baca INSERT FORM
    $ins = [];
    foreach ($INSROWS as $r) {
        $ins[$r] = cellNum($wsI, $r, $col);
    }

    // Baca J NRML
    $jn = []; $jcol = null;
    if (null !== $ss->getSheetByName('J NRML')) {
        $wsJ = $ss->getSheetByName('J NRML');
        for ($c = 11; $c <= 44; $c++) {
            if (cellDate($wsJ, 8, $c) === $date) {
                $jcol = $c;
                break;
            }
        }
        if ($jcol !== null) {
            foreach ($JNROWS as $r) {
                $jn[$r] = cellNum($wsJ, $r, $jcol);
            }
        }
    }

    // Baca Y NRML
    $yn = [];
    if ($jcol !== null && null !== $ss->getSheetByName('Y NRML')) {
        $wsY = $ss->getSheetByName('Y NRML');
        foreach ([10, 12, 14] as $r) {
            $yn[$r] = cellNum($wsY, $r, $jcol);
        }
    }

    // Baca label dari sheet sumber (INSERT FORM col B hanya berisi formula referensi)
    $lbl = [];
    // JBLOCK (96-101) -> J NRML col F (6), rows [14,16,18,20,22,33]
    $jLabelMap = [96=>14, 97=>16, 98=>18, 99=>20, 100=>22, 101=>33];
    foreach ($jLabelMap as $r => $srcRow) {
        $lbl[$r] = getLabel($wsJ, $srcRow, 6);
    }
    // FC_SPML (204-211) -> FCML Tech Crew col E (5), rows 18-25
    $wsFC = $ss->getSheetByName('FCML Tech Crew');
    foreach ($FC_SPML as $idx => $r) {
        $lbl[$r] = getLabel($wsFC, 18 + $idx, 5);
    }
    // CC_SPML (215-217) -> CCML Cabin Crew col D (4), rows 23-25
    $wsCC = $ss->getSheetByName('CCML Cabin Crew');
    foreach ($CC_SPML as $idx => $r) {
        $lbl[$r] = getLabel($wsCC, 23 + $idx, 4);
    }
    // Y_SPMLC (170-194) -> Y SPML  col E (5), rows 15-39
    $wsYSPML = $ss->getSheetByName('Y SPML ');
    foreach ($Y_SPMLC as $idx => $r) {
        $lbl[$r] = getLabel($wsYSPML, 15 + $idx, 5);
    }
    // BBML_ROW (196) - formula, keep empty or read from INSERT FORM
    $lbl[$BBML_ROW] = '';

    // ---- Bangun nilai header ----
    $dObj = new DateTime($date);
    $headers = [];
    $vcells = [];
    $dcells = [];
    $ecells = [];

    $headers['B2'] = 'DATE : ' . strtoupper($dObj->format('j M Y'));
    $headers['C2'] = 'PRIOR TO FLIGHT : 24 HOURS';
    $j = $ins[$J_FNL] ?? null;
    $y = $ins[$Y_FNL] ?? null;
    if ($j !== null && $j != 0) {
        $headers['B3'] = 'J : ' . $j;
    }
    if ($y !== null && $y != 0) {
        $headers['C3'] = 'Y : ' . $y;
    }

    $crewText = function ($tag, $nrmlRow, $coderows) use ($ins, $lbl) {
        $parts = [];
        $nn = $ins[$nrmlRow] ?? null;
        if ($nn !== null && $nn != 0) {
            $parts[] = $nn . 'NRML';
        }
        $codes = [];
        foreach ($coderows as $r) {
            $n = $ins[$r] ?? null;
            if ($n !== null && $n != 0) {
                $codes[] = ['C' => codeOf($lbl[$r] ?? ''), 'N' => $n];
            }
        }
        usort($codes, function ($a, $b) {
            return strcmp($a['C'], $b['C']);
        });
        foreach ($codes as $cd) {
            $parts[] = $cd['N'] . $cd['C'];
        }
        if (!$parts) {
            return null;
        }
        return $tag . ' : ' . implode(',', $parts);
    };

    $t4 = $crewText('FC', $FC_NRML, $FC_SPML);
    if ($t4 !== null) {
        $headers['B4'] = $t4;
    }
    $t4 = $crewText('CC', $CC_NRML, $CC_SPML);
    if ($t4 !== null) {
        $headers['C4'] = $t4;
    }

    $jb = [];
    foreach ([96, 97, 101, 98, 99] as $r) {
        $n = $ins[$r] ?? null;
        if ($n !== null && $n != 0) {
            $jb[] = $n . ' ' . codeOf($lbl[$r] ?? '');
        }
    }
    $n100 = $ins[100] ?? null;
    if ($n100 !== null && $n100 != 0) {
        $jb[] = $n100 . ' ' . codeOf($lbl[100] ?? '');
    }
    if ($jb) {
        $headers['B6'] = implode(', ', $jb);
    }
    $headers['E7'] = 'FC';

    $spmlText = function ($rows) use ($ins, $lbl, $BBML_ROW) {
        $codes = [];
        foreach ($rows as $r) {
            $n = $ins[$r] ?? null;
            if ($n !== null && $n != 0) {
                $codes[] = ['C' => codeOf($lbl[$r] ?? ''), 'N' => $n];
            }
        }
        usort($codes, function ($a, $b) {
            return strcmp($a['C'], $b['C']);
        });
        $out = [];
        foreach ($codes as $cd) {
            $out[] = $cd['N'] . ' ' . $cd['C'];
        }
        $b = $ins[$BBML_ROW] ?? null;
        if ($b !== null && $b != 0) {
            $out[] = $b . ' BBML';
        }
        return $out;
    };

    $yb = [];
    $m1 = $ins[168] ?? null;
    if ($m1 !== null && $m1 != 0) {
        $yb[] = $m1 . ' MBF1';
    }
    $m2 = $ins[169] ?? null;
    if ($m2 !== null && $m2 != 0) {
        $yb[] = $m2 . ' MCK1';
    }
    $yb = array_merge($yb, $spmlText($Y_SPMLC));
    if ($yb) {
        $headers['B9'] = implode(', ', $yb);
    }
    $f5 = $spmlText($Y_SPMLC);
    if ($f5) {
        $headers['F5'] = implode(', ', $f5);
    }

    $spmlRaw = function ($rows) use ($ins, $lbl) {
        $codes = [];
        foreach ($rows as $r) {
            $n = $ins[$r] ?? null;
            if ($n !== null && $n != 0) {
                $codes[] = ['C' => codeOf($lbl[$r] ?? ''), 'N' => $n];
            }
        }
        usort($codes, function ($a, $b) {
            return strcmp($a['C'], $b['C']);
        });
        $out = [];
        foreach ($codes as $cd) {
            $out[] = $cd['N'] . ' ' . $cd['C'];
        }
        return $out;
    };
    $f7 = $spmlRaw($FC_SPML);
    if ($f7) {
        $headers['F7'] = implode(', ', $f7);
    }
    $f9 = $spmlRaw($CC_SPML);
    if ($f9) {
        $headers['F9'] = implode(', ', $f9);
    }

    foreach ($DMAP as $k => $r) {
        $n = $jn[$r] ?? null;
        if ($n !== null && $n != 0) {
            $dcells[$k] = $n;
        }
    }
    foreach ($YDMAP as $k => $r) {
        $n = $yn[$r] ?? null;
        if ($n !== null && $n != 0) {
            $dcells[$k] = $n;
        }
    }
    foreach ($EMAP as $k => $r) {
        $n = $jn[$r] ?? null;
        if ($n !== null && $n != 0) {
            $ecells[$k] = $n;
        }
    }
    foreach ($YEINS as $k => $r) {
        $n = $ins[$r] ?? null;
        if ($n !== null && $n != 0) {
            $ecells[$k] = $n;
        }
    }
    $t0 = $ins[$J_SPML] ?? null;
    if ($t0 !== null && $t0 != 0) {
        $vcells['F3'] = $t0;
    }

    // ---- Muat template + isi ----
    $tplTmp = tempnam(sys_get_temp_dir(), 'tpl');
    file_put_contents($tplTmp, base64_decode(TEMPLATE_B64));
    $out = IOFactory::load($tplTmp);
    @unlink($tplTmp);
    if (null === $out->getSheetByName('Meal Number Check List')) {
        throw new Exception('Template tidak valid.');
    }
    $wo = $out->getSheetByName('Meal Number Check List');

    // Bersihkan D/E baris 12-49 (lewati sel merge)
    for ($r = 12; $r <= 49; $r++) {
        foreach ([4, 5] as $c) {
            $coord = Coordinate::stringFromColumnIndex($c) . $r;
            $cell = $wo->getCell($coord);
            $mr = $cell->getMergeRange();
            if ($mr !== null && $mr !== $coord) {
                continue;
            }
            $cell->setValue(null);
        }
    }
    foreach ($headers as $k => $v) {
        $wo->setCellValue($k, (string)$v);
        $wo->getStyle($k)->getAlignment()->setWrapText(true);
    }
    foreach ($vcells as $k => $v) {
        $wo->setCellValue($k, $v);
    }
    foreach ($dcells as $k => $v) {
        $wo->setCellValue($k, $v);
    }
    foreach ($ecells as $k => $v) {
        $wo->setCellValue($k, $v);
    }

    // ---- Stream hasil sebagai download ----
    $stamp = date('YmdHis');
    $fname = 'v2 Meal Number Check List JX 2026D - ' . $dObj->format('d-m-Y') . ' - ' . $stamp . '.xlsx';
    while (ob_get_level()) {
        ob_end_clean();
    }
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . $fname . '"');
    header('Cache-Control: no-store');
    $writer = IOFactory::createWriter($out, 'Xlsx');
    $writer->save('php://output');
    exit;
} catch (\Throwable $e) {
    http_response_code(500);
    echo json_encode(['error' => $e->getMessage()]);
    exit;
}
