<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Carbon;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class ConvertJsonToExcel extends Command
{
    protected $signature = 'report:convert-json';
    protected $description = 'Converts report.json from trello to Excel using optimized memory handling';

    // Move configuration to constants or config files for easier maintenance
    private const DISK = 'local';
    private const INPUT_FILE = 'report.json';
    private const OUTPUT_FILE = 'report_export.xlsx';

    public function handle()
    {
        $disk = Storage::disk(self::DISK);

        if (!$disk->exists(self::INPUT_FILE)) {
            $this->error("❌ File not found: " . self::INPUT_FILE);
            return 1;
        }

        $this->info("📥 Reading JSON data...");
        $rawContent = $disk->get(self::INPUT_FILE);
        $decoded = json_decode($rawContent, true);

        if (!isset($decoded['cards']) || empty($decoded['cards'])) {
            $this->error("❌ Invalid JSON structure: 'cards' key missing or empty.");
            return 1;
        }

        // 1. Transform and Filter Data in one pass
        $this->info("⚙️ Filtering and Transforming data...");
        $reportData = $this->transformData($decoded['cards']);

        if (empty($reportData)) {
            $this->warn("⚠️ No data matches the filter criteria.");
            return 0;
        }

        // 2. Initialize Spreadsheet
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        
        $headers = $this->getHeaders();
        $this->setupHeaderStyle($sheet, $headers);

        // 3. Populate Rows using Progress Bar
        $bar = $this->output->createProgressBar(count($reportData));
        $bar->start();

        foreach ($reportData as $index => $rowValues) {
            $rowIndex = $index + 2; // Start at row 2 (below header)
            
            // fromArray is significantly faster than setCellValueByColumnAndRow in loops
            $sheet->fromArray(array_values($rowValues), null, "A{$rowIndex}");
            
            $bar->advance();
        }

        $bar->finish();
        $this->newLine();

        // 4. Auto-size columns
        foreach (range('A', $sheet->getHighestColumn()) as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        // 5. Streamed Save (Memory Efficient)
        $this->info("💾 Saving to storage...");
        $tempPath = tempnam(sys_get_temp_dir(), 'laravel_xlsx');
        $writer = new Xlsx($spreadsheet);
        $writer->save($tempPath);

        $disk->put(self::OUTPUT_FILE, fopen($tempPath, 'r+'));
        unlink($tempPath);

        $this->info("✅ Done! Saved as: " . self::OUTPUT_FILE);
        return 0;
    }

    private function transformData(array $cards): array
    {
        $excludeLists = ['67e234f5b56118741d87d616', '6826ec0930b95b46166354f8'];

        return collect($cards)
            ->filter(fn($card) => is_null($card['dateClosed']) && !in_array($card['idList'], $excludeLists))
            ->map(function ($card) {
                // Centralized Date Parsing logic
                $parseDate = fn($date) => $date ? Carbon::parse($date)->format('d/m/y H:i') : '-';

                $dateCreated = $parseDate($card['start']);
                $lastUpdated = $parseDate($card['dateLastActivity']);
                $slaDueDate = $parseDate($card['dateCompleted']);
                if($dateCreated == '-' && $slaDueDate == '-')
                {
                    $dateCreated && $slaDueDate = $lastUpdated;
                }
                if($dateCreated == '-' && $slaDueDate != '-')
                {
                    $dateCreated = $slaDueDate;
                }

                return [
                    'id'               => $card['id'],
                    'date_created'     => $dateCreated,
                    'subject'          => 'Perangkat Head Office',
                    'from'             => 'Samira Tasyaa',
                    'from_email'       => 'samiratasyaa@gmail.com',
                    'priority'         => 'normal',
                    'department'       => 'PT. Centria Integrity Advisory (CIA)',
                    'help_topic'       => 'Aplikasi / TRUST (ManRisk)',
                    'source'           => 'web',
                    'current_status'   => 'closed',
                    'last_updated'     => $lastUpdated,
                    'sla_due_date'     => $slaDueDate,
                    'sla_plan'         => 'Default SLA',
                    'due_date'         => $card['due'] ? $parseDate($card['due']) : "Yes",
                    'closed_date'      => 'Default SLA',
                    'overdue'          => ($card['dueComplete'] ?? false) ? "No" : "Yes",
                    'merged'           => 'No',
                    'linked'           => 'Yes',
                    'answered'         => 'Agent CIA',
                    'agent_assigned'   => 2,
                    'team_assigned'    => null,
                    'thread_count'     => null,
                    'reopen_count'     => null,
                    'attachment_count' => null,
                    'task_count'       => null
                ];
            })
            ->values()
            ->all();
    }

    private function getHeaders(): array
    {
        return [
            'Ticket Number', 'Date Created', 'Subject', 'From', 'From Email',
            'Priority', 'Department', 'Help Topic', 'Source', 'Current Status',
            'Last Updated', 'SLA Due Date', 'SLA Plan', 'Due Date', 'Closed Date',
            'Overdue', 'Merged', 'Linked', 'Answered', 'Agent Assigned',
            'Team Assigned', 'Thread Count', 'Reopen Count', 'Attachment Count', 'Task Count'
        ];
    }

    private function setupHeaderStyle($sheet, $headers)
    {
        $sheet->fromArray($headers, null, 'A1');
        
        $highestCol = $sheet->getHighestColumn();
        $headerRange = "A1:{$highestCol}1";

        $sheet->getStyle($headerRange)->applyFromArray([
            'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '4472C4']
            ],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
        ]);
    }
}