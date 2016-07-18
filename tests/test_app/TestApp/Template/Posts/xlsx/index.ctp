<?php
$worksheet = $this->PhpExcel->setActiveSheetIndex(0);
$worksheet->setCellValueByColumnAndRow(0, 0, 'Test string');
