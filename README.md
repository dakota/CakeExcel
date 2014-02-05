CakeExcel
=========

[![Coverage Status](https://coveralls.io/repos/dakota/CakeExcel/badge.png)](https://coveralls.io/r/dakota/CakeExcel)
[![Build Status](https://travis-ci.org/dakota/CakeExcel.png?branch=master)](https://travis-ci.org/dakota/CakeExcel)

A plugin to generate Excel files with CakePHP.

Requirements
------------

* PHP 5.2.8
* CakePHP 2.1+
* [PHPExcel](https://github.com/PHPOffice/PHPExcel)
* Composer


Installation
------------

Add to your composer.json file (dakota/cake-excel)
```
"dakota/cake-excel": ">=1.0"
```

Or run
```
composer require "dakota/cake-excel >=1.0"
```

Usage
-----

In app/Config/bootstrap.php add:
```
CakePlugin::load('CakeExcel', array('bootstrap' => true, 'routes' => true));
```


Add the RequestHandler component to AppController, and map xlsx to the CakeExcel plugin
```
'RequestHandler' => array(
	'viewClassMap' => array(
		'xlsx' => 'CakeExcel.Excel'
	)
),
```

Create a link to the a action with the .xlsx extension
```
$this->Html->link('Excel file', array('ext' => 'xlsx'));
```

Place the view templates in a 'Excel' subdir, for instance `app/View/Invoices/Excel/index.ctp`

Inside your view file you will have access to the PHPExcel library with `$this->PhpExcel`. Please see the [PHPExcel](https://github.com/PHPOffice/PHPExcel) documentation for a guide on how to use PHPExcel.
