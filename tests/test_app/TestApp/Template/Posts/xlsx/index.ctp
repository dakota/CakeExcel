<?php
$worksheet = $this->Spreadsheet->setActiveSheetIndex(0);
$worksheet->setCellValueByColumnAndRow(0, 0, 'Test string');
