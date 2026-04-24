<?php

declare(strict_types=1);

error_reporting(E_ALL);
ini_set('display_errors', '0');
ini_set('log_errors', '1');
ini_set('max_execution_time', '300');
ini_set('memory_limit', '1024M');

header('Content-Type: application/json; charset=utf-8');

$autoloadPath = __DIR__ . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';

if (!file_exists($autoloadPath)) {
    jsonError(
        'vendor/autoload.php not found. Run: composer require phpoffice/phpspreadsheet',
        500
    );
}

require $autoloadPath;

use PhpOffice\PhpSpreadsheet\IOFactory;

function jsonResponse(array $payload, int $statusCode = 200): void
{
    http_response_code($statusCode);
    echo json_encode($payload, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
    exit;
}

function jsonError(string $message, int $statusCode = 400, array $details = []): void
{
    $payload = [
        'success' => false,
        'error' => $message
    ];

    if (!empty($details)) {
        $payload['details'] = $details;
    }

    jsonResponse($payload, $statusCode);
}

function normalizeToUSFormat($text): string
{
    if ($text === null) {
        return '';
    }

    $text = trim((string)$text);

    if ($text === '') {
        return '';
    }

    if (function_exists('mb_strtoupper')) {
        $text = mb_strtoupper($text, 'UTF-8');
    } else {
        $text = strtoupper($text);
    }

    $accents = [
        'Š'=>'S','š'=>'S','Ž'=>'Z','ž'=>'Z',
        'À'=>'A','Á'=>'A','Â'=>'A','Ã'=>'A','Ä'=>'A','Å'=>'A',
        'Æ'=>'A','Ç'=>'C',
        'È'=>'E','É'=>'E','Ê'=>'E','Ë'=>'E',
        'Ì'=>'I','Í'=>'I','Î'=>'I','Ï'=>'I',
        'Ñ'=>'N',
        'Ò'=>'O','Ó'=>'O','Ô'=>'O','Õ'=>'O','Ö'=>'O','Ø'=>'O',
        'Ù'=>'U','Ú'=>'U','Û'=>'U','Ü'=>'U',
        'Ý'=>'Y','Þ'=>'B','ß'=>'SS'
    ];

    $text = strtr($text, $accents);

    if (function_exists('iconv')) {
        $converted = @iconv('UTF-8', 'ASCII//TRANSLIT//IGNORE', $text);

        if ($converted !== false && $converted !== '') {
            $text = $converted;
        }
    }

    $text = preg_replace('/\s+/', ' ', $text);

    return trim((string)$text);
}

function getUpload(string $fieldName): array
{
    if (!isset($_FILES[$fieldName])) {
        jsonError("Missing uploaded file field: {$fieldName}");
    }

    $file = $_FILES[$fieldName];

    if (!is_array($file) || ($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
        $errorCode = $file['error'] ?? UPLOAD_ERR_NO_FILE;
        jsonError("Upload failed for {$fieldName}.", 400, ['upload_error_code' => $errorCode]);
    }

    if (!is_uploaded_file($file['tmp_name'])) {
        jsonError("Invalid upload received for {$fieldName}.");
    }

    $originalName = (string)($file['name'] ?? 'uploaded_file');
    $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
    $allowedExtensions = ['xlsx', 'xlsm', 'xls'];

    if (!in_array($extension, $allowedExtensions, true)) {
        jsonError("Invalid file type for {$fieldName}. Please upload an XLSX, XLSM, or XLS file.");
    }

    return [
        'tmp_name' => $file['tmp_name'],
        'name' => $originalName,
        'extension' => $extension
    ];
}

function resolvePythonPath(): string
{
    $commands = [];

    if (stripos(PHP_OS_FAMILY, 'Windows') !== false) {
        $commands = [
            'where python 2>&1',
            'where py 2>&1',
            'where python3 2>&1'
        ];
    } else {
        $commands = [
            'command -v python3 2>&1',
            'command -v python 2>&1'
        ];
    }

    foreach ($commands as $command) {
        $output = [];
        $returnCode = 1;

        @exec($command, $output, $returnCode);

        if ($returnCode === 0 && !empty($output)) {
            foreach ($output as $possiblePython) {
                $possiblePython = trim((string)$possiblePython);

                if ($possiblePython !== '' && file_exists($possiblePython)) {
                    return $possiblePython;
                }
            }
        }
    }

    throw new Exception(
        'Python executable not found. Install Python and make sure it is added to PATH. Test with: python --version and where python.'
    );
}

function makeTemporaryExcelPath(string $prefix, string $extension): string
{
    $extension = strtolower($extension);
    $extension = in_array($extension, ['xlsx', 'xlsm', 'xls'], true) ? $extension : 'xlsx';

    $base = rtrim(sys_get_temp_dir(), DIRECTORY_SEPARATOR)
        . DIRECTORY_SEPARATOR
        . $prefix
        . '_'
        . bin2hex(random_bytes(12));

    return $base . '.' . $extension;
}

function decryptExcelFileWithPython(string $inputFile, string $password, string $outputFile): array
{
    if (!function_exists('exec')) {
        throw new Exception('PHP exec() is disabled. Enable exec() in php.ini to decrypt password-protected Excel files.');
    }

    $pythonPath = resolvePythonPath();
    $decryptScript = __DIR__ . DIRECTORY_SEPARATOR . 'decrypt.py';

    if (!file_exists($decryptScript)) {
        throw new Exception('decrypt.py was not found in the project folder.');
    }

    if (!file_exists($inputFile)) {
        throw new Exception('Input Excel file was not found in the temporary upload location.');
    }

    $cmd = escapeshellarg($pythonPath) . ' '
        . escapeshellarg($decryptScript) . ' '
        . escapeshellarg($inputFile) . ' '
        . escapeshellarg($password) . ' '
        . escapeshellarg($outputFile) . ' 2>&1';

    $output = [];
    $returnCode = 1;

    exec($cmd, $output, $returnCode);

    return [
        'returnCode' => $returnCode,
        'output' => $output
    ];
}

function createReaderByExtension(string $extension)
{
    $extension = strtolower($extension);

    if ($extension === 'xlsx' || $extension === 'xlsm') {
        return IOFactory::createReader('Xlsx');
    }

    if ($extension === 'xls') {
        return IOFactory::createReader('Xls');
    }

    throw new Exception('Unsupported Excel extension: ' . $extension);
}

function loadSpreadsheetDirectly(string $filePath, string $originalName)
{
    $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));

    if (($extension === 'xlsx' || $extension === 'xlsm') && !class_exists('ZipArchive')) {
        throw new Exception(
            'PHP ZIP / ZipArchive extension is missing. Open C:\\xampp\\php\\php.ini, enable extension=zip, confirm extension_dir="C:\\xampp\\php\\ext", then restart Apache.'
        );
    }

    try {
        $readerType = IOFactory::identify($filePath);
        $reader = IOFactory::createReader($readerType);
        $reader->setReadDataOnly(true);

        return $reader->load($filePath);
    } catch (Throwable $identifyError) {
        $reader = createReaderByExtension($extension);
        $reader->setReadDataOnly(true);

        return $reader->load($filePath);
    }
}

function loadSpreadsheetOptionalPassword(array $upload, string $excelPassword)
{
    $filePath = $upload['tmp_name'];
    $originalName = $upload['name'];
    $extension = $upload['extension'];
    $decryptedFile = null;

    try {
        return loadSpreadsheetDirectly($filePath, $originalName);
    } catch (Throwable $originalError) {
        if ($excelPassword === '') {
            throw new Exception(
                'The Excel file could not be opened normally and no password was provided. Original error: ' . $originalError->getMessage()
            );
        }

        $decryptedFile = makeTemporaryExcelPath('decrypted_excel', $extension);
        $decryptResult = decryptExcelFileWithPython($filePath, $excelPassword, $decryptedFile);

        if ($decryptResult['returnCode'] !== 0) {
            if ($decryptedFile !== null && file_exists($decryptedFile)) {
                @unlink($decryptedFile);
            }

            $pythonOutput = trim(implode("\n", $decryptResult['output']));
            $pythonOutput = $pythonOutput !== '' ? $pythonOutput : 'No Python output was returned.';

            throw new Exception(
                'Unable to open/decrypt Excel file. Original reader error: '
                . $originalError->getMessage()
                . ' Python decrypt error: '
                . $pythonOutput
            );
        }

        if (!file_exists($decryptedFile) || filesize($decryptedFile) === 0) {
            if ($decryptedFile !== null && file_exists($decryptedFile)) {
                @unlink($decryptedFile);
            }

            throw new Exception('Python finished, but the decrypted Excel file was not created or is empty.');
        }

        try {
            $spreadsheet = loadSpreadsheetDirectly($decryptedFile, $originalName);
            @unlink($decryptedFile);

            return $spreadsheet;
        } catch (Throwable $decryptedLoadError) {
            if ($decryptedFile !== null && file_exists($decryptedFile)) {
                @unlink($decryptedFile);
            }

            throw new Exception(
                'Decryption completed, but the decrypted Excel file could not be loaded. Reader error: '
                . $decryptedLoadError->getMessage()
            );
        }
    }
}

function buildEmailLookup(array $emailUpload, string $excelPassword): array
{
    $spreadsheet = loadSpreadsheetOptionalPassword($emailUpload, $excelPassword);
    $lookup = [];

    foreach ($spreadsheet->getAllSheets() as $sheet) {
        $sheetName = $sheet->getTitle();
        $rows = $sheet->toArray(null, true, true, true);

        foreach ($rows as $rowNumber => $row) {
            $firstName = normalizeToUSFormat($row['C'] ?? '');
            $lastName = normalizeToUSFormat($row['D'] ?? '');
            $email = strtolower(trim((string)($row['B'] ?? '')));

            if ($email !== '' && strpos($email, '@') !== false) {
                $key = trim($firstName . ' ' . $lastName);

                if ($key !== '') {
                    $lookup[$key] = [
                        'email' => $email,
                        'email_source_file' => $emailUpload['name'],
                        'email_source_sheet' => $sheetName,
                        'email_source_row' => (int)$rowNumber,
                        'email_source_columns' => 'B=email, C=first_name, D=last_name',
                        'email_source_location' => $emailUpload['name'] . ' → ' . $sheetName . ' → Row ' . $rowNumber . ' → Column B'
                    ];
                }
            }
        }
    }

    return $lookup;
}

function processConfidentialRecords(array $confidentialUpload, string $excelPassword, array $emailLookup): array
{
    $spreadsheet = loadSpreadsheetOptionalPassword($confidentialUpload, $excelPassword);
    $sheet = $spreadsheet->getActiveSheet();
    $rows = $sheet->toArray(null, true, true, true);

    $matched = [];
    $nonMatched = [];
    $totalConfidentialRecords = 0;

    foreach ($rows as $row) {
        $empNoRaw = trim((string)($row['A'] ?? ''));

        if (
            strcasecmp($empNoRaw, 'Empno') === 0 ||
            strcasecmp($empNoRaw, 'Emp No') === 0 ||
            strcasecmp($empNoRaw, 'Employee No') === 0 ||
            strcasecmp($empNoRaw, 'Employee Number') === 0
        ) {
            continue;
        }

        $empNo = $empNoRaw;
        $lastName = normalizeToUSFormat($row['B'] ?? '');
        $firstName = normalizeToUSFormat($row['C'] ?? '');

        if ($empNo === '' && $lastName === '' && $firstName === '') {
            continue;
        }

        $totalConfidentialRecords++;
        $searchKey = trim($firstName . ' ' . $lastName);
        $fullName = trim($firstName . ' ' . $lastName);

        if ($searchKey !== '' && isset($emailLookup[$searchKey])) {
            $source = $emailLookup[$searchKey];

            $matched[] = [
                'empno' => $empNo,
                'first_name' => $firstName,
                'last_name' => $lastName,
                'full_name' => $fullName,
                'email' => $source['email'],
                'status' => 'Matched',
                'email_source_file' => $source['email_source_file'],
                'email_source_sheet' => $source['email_source_sheet'],
                'email_source_row' => $source['email_source_row'],
                'email_source_columns' => $source['email_source_columns'],
                'email_source_location' => $source['email_source_location']
            ];
        } else {
            $nonMatched[] = [
                'empno' => $empNo,
                'first_name' => $firstName,
                'last_name' => $lastName,
                'full_name' => $fullName,
                'email' => '',
                'status' => 'No Email Found',
                'email_source_file' => '',
                'email_source_sheet' => '',
                'email_source_row' => '',
                'email_source_columns' => '',
                'email_source_location' => 'No matching email source found'
            ];
        }
    }

    return [
        'total_confidential_records' => $totalConfidentialRecords,
        'matched' => $matched,
        'non_matched' => $nonMatched
    ];
}

try {
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
        jsonError('Invalid request method. Use POST only.', 405);
    }

    $emailUpload = getUpload('email_file');
    $confidentialUpload = getUpload('confidential_file');
    $excelPassword = trim((string)($_POST['excel_password'] ?? ''));

    $emailLookup = buildEmailLookup($emailUpload, $excelPassword);
    $confidentialResult = processConfidentialRecords($confidentialUpload, $excelPassword, $emailLookup);

    $matched = $confidentialResult['matched'];
    $nonMatched = $confidentialResult['non_matched'];

    jsonResponse([
        'success' => true,
        'summary' => [
            'total_confidential_records' => $confidentialResult['total_confidential_records'],
            'total_email_users_loaded' => count($emailLookup),
            'matched_count' => count($matched),
            'non_matched_count' => count($nonMatched)
        ],
        'matched' => $matched,
        'non_matched' => $nonMatched
    ]);
} catch (Throwable $error) {
    jsonError($error->getMessage(), 500);
}
