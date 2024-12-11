<?php
// Start Generation Here
try {
    // Check if the export format is set
    if (isset($_POST['format'])) {
        $format = $_POST['format'];
        $data = isset($_SESSION['data']) ? $_SESSION['data'] : [];
        $selected_columns = isset($_SESSION['selected_columns']) ? $_SESSION['selected_columns'] : [];

        // Prepare data for export
        $exportData = [];
        foreach ($data as $row) {
            $exportRow = [];
            foreach ($selected_columns as $column) {
                $exportRow[$column] = isset($row[$column]) ? $row[$column] : '';
            }
            $exportData[] = $exportRow;
        }

        // Export to CSV
        if ($format === 'csv') {
            header('Content-Type: text/csv');
            header('Content-Disposition: attachment; filename="export.csv"');
            $output = fopen('php://output', 'w');
            if ($output === false) {
                throw new Exception('Failed to open output stream for CSV export.');
            }
            fputcsv($output, $selected_columns); // Add headers
            foreach ($exportData as $exportRow) {
                fputcsv($output, $exportRow);
            }
            fclose($output);
            exit;
        }

        // Export to TXT
        if ($format === 'txt') {
            header('Content-Type: text/plain');
            header('Content-Disposition: attachment; filename="export.txt"');
            $output = fopen('php://output', 'w');
            if ($output === false) {
                throw new Exception('Failed to open output stream for TXT export.');
            }
            foreach ($exportData as $exportRow) {
                fwrite($output, implode("\t", $exportRow) . PHP_EOL);
            }
            fclose($output);
            exit;
        }
    }
} catch (Exception $e) {
    // Handle exceptions (logging, user feedback, etc.)
    error_log($e->getMessage());
    echo 'An error occurred during the export process.';
    exit;
}
