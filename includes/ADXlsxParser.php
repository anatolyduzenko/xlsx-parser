<?

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 *
 * XLSX Parser Class 
 *
 */

class ADXlsxParser 
{

	private $post;
	
	function __construct()
	{
		
	}

	function run ()
	{
		$html =  "<h2>Welocme to Xlsx Parser!</h2><hr>";

		// File upload form
		$html .= "<form enctype='multipart/form-data' method='post' action=".admin_url( 'admin-post.php' ).">";
		$html .= "<input type='hidden' name='action' value='xlsx_parser_process_form_data'>";
		$html .= "<input type='file' name='xlsx_file' accept='.xls, .xlsx' required/><br>";
		$html .= "<label>Download result</label><input type='checkbox' name='export_file' /><hr>";
		$html .= "<input type='submit' class='button button-primary' name='submit' value='Analyze File'/>";
		$html .= "</form>";
		$html .= "<div id='parse-analysis-result'>".$_SESSION['result']."</div>";

		echo $html;
	}

	function analyze_file()
	{
		if ( ! function_exists( 'wp_handle_upload' ) ) {
		    require_once( ABSPATH . 'wp-admin/includes/file.php' );
		}
		if ( ! empty( $_POST ) ) {
			if(!empty($_FILES['xlsx_file'])) {
				$_SESSION['xlsx_parsed_file'] = $_FILES['xlsx_file']['name'];
				add_filter('upload_dir', 'xlsx_parser_upload_folder');
				$movefile = wp_handle_upload( $_FILES['xlsx_file'], ['test_form' => false] );
				remove_filter('upload_dir', 'xlsx_parser_upload_folder');
				if(isset($movefile['error'])) {
					$_SESSION['result'] = "<span class='error'>".$movefile['error']."</span>";
				} else {
					$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
					$reader->setReadDataOnly(true);
					$spreadsheet = $reader->load($movefile['file']);

					// Reading is here 
					$worksheet = $spreadsheet->getActiveSheet();
					$highestRow = $worksheet->getHighestRow();
					$highestColumn = $worksheet->getHighestDataColumn();
					$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
					$data = [];
					for ($row = 0; $row <= $highestRow; ++$row) {
						$rdata = [];
						for ($col = 0; $col <= $highestColumnIndex; ++$col) { // Hidden column
							$cell = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
							if(!is_null($cell) && $cell !== "#NULL!") 
								$rdata[] = $cell;
						}
						$data[] = $rdata;
					}

					$xdata = [];
					$html = "<table class='xlsx-analysis-result'>";
					$html .= "<tr><th>Дата</th><th>Сумма</th><th>Уч</th><th>Контрагент</th><th>Описание</th><th>Доп.Описание</th></tr>";
					$xdata[] = [
						"Дата",
						"Сумма",
						"Уч-к",
						"Контрагент",
						"Описание",
						"Доп.Описание"
					];
					foreach ($data as $row) {
						// if(count($row) == 9) {
							if($row[5] == "76.09.1"&& in_array($row[4], ["50.01", "51"])) {
								// Date, Amount, Item, Employer, DocDescr, DocDescr2,
								// Item 
								$re = '/\(№*.+\)/m';
								preg_match_all($re, $row[3], $matches, PREG_SET_ORDER, 0);
								$partition = str_replace("(№", "", str_replace(")", "", $matches[0][0]));
								$re1 = '/(\p{Cyrillic}+\ (.)\.(.)\.)\ /u';
								preg_match_all($re1, $row[3], $matches1, PREG_SET_ORDER, 0);
								$participant = trim($matches1[0][0]);
								$html .= "<tr><td>".$row[0]."</td><td>".number_format($row[6], 2)."</td><td>".$partition."</td><td>".$participant."</td><td>".$row[1]."</td><td>".$row[2]."</td></tr>";
								$total += $row[6];
								$xdata[] = [
									$row[0],
									number_format($row[6], 2),
									$partition,
									$participant,
									$row[1],
									$row[2],
								];
							}
						// }
					}
					$xdata[] = ['Итого', number_format($total, 2)];

					$html .= "<tr><th>Итого</th><th>".number_format($total, 2)."</th><th colspan='4'></th></tr>";
					$html .= "</table>";
					$_SESSION['result'] = $html;
					if($_POST["export_file"]) {
						// Export Excel
						$spreadsheet = new Spreadsheet;
				        $spreadsheet->getActiveSheet()->fromArray($xdata, NULL, 'A1');
				        foreach(range('A','ZZ') as $columnID) {
				            $spreadsheet->getActiveSheet()->getColumnDimension($columnID)
				                ->setAutoSize(true);
				        }

				        $cellStyle = [ '0' => 
				            [
				                'font' => [
				                    'size' => 11,
				                    'bold' =>true
				                ],
				                'borders' => [
				                    'allBorders' => [
				                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				                        'color' => ['argb' => 'FF000000']
				                    ]
				                ]
				            ],
				            '1' => [
				                'alignment' => [
				                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER
				                ],
				            ],
				            '2' => [
				                'borders' => [
				                    'left' => [
				                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				                        'color' => ['argb' => 'FF000000']
				                    ],
				                    'right' => [
				                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				                        'color' => ['argb' => 'FF000000']
				                    ]
				                ]
				            ],
				            '3' => [
				                'numberFormat' => [
				                    'formatCode' => '###0.00'
				                ]
				            ],
				            '4' => [
				                'numberFormat' => [
				                    'formatCode' => \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT
				                ]
				            ],
				            '5' => [
				                'font' => [
				                    'size' => 11,
				                ],
				                'borders' => [
				                    'allBorders' => [
				                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				                        'color' => ['argb' => 'FF000000']
				                    ]
				                ]
				            ],
				        ];
				        $highestCol = $spreadsheet->getActiveSheet()->getHighestColumn(2);
				        $highestRow = $spreadsheet->getActiveSheet()->getHighestRow();
				        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestCol);
				        // Header
				        $spreadsheet->getActiveSheet()->getStyle("A1:".$highestCol."1")
				            ->applyFromArray($cellStyle[0]);
				        $spreadsheet->getActiveSheet()->getStyle("A2:".$highestCol."2")
				            ->applyFromArray($cellStyle[1]);
				        //Amounts
				        $spreadsheet->getActiveSheet()->getStyle("G3:".$highestCol.$highestRow)
				                ->applyFromArray($cellStyle[1])
				                ->applyFromArray($cellStyle[3]);
				        // Total row
				        // $spreadsheet->getActiveSheet()->mergeCells("A".$highestRow.":F".$highestRow);
				         $spreadsheet->getActiveSheet()->getStyle("A".$highestRow.":".$highestCol.$highestRow)
				            ->applyFromArray($cellStyle[5]);
				        foreach(range('A', $highestCol) as $columnID) {
				            $spreadsheet->getActiveSheet()->getStyle($columnID."1:".$columnID.$highestRow)
				                ->applyFromArray($cellStyle[2]);
				        }
				        $spreadsheet->getActiveSheet()->setTitle('Документы '.$participant);

				        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
				        if(!is_dir(plugin_dir_path( __FILE__ ).'../files'.DIRECTORY_SEPARATOR."export")) {
				        	mkdir(plugin_dir_path( __FILE__ ).'../files'.DIRECTORY_SEPARATOR."export");
				        }
				        $writer->save(plugin_dir_path( __FILE__ ).'../files'.DIRECTORY_SEPARATOR."export".DIRECTORY_SEPARATOR."Документы_{$participant}.xlsx");
					}
				}
			} else {
				$_SESSION['result'] = "<span class='error'> The file is required.</span>";
			}
		}
	}
}