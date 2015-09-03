CakeExcel
=========

[![Coverage Status](https://coveralls.io/repos/dakota/CakeExcel/badge.png)](https://coveralls.io/r/dakota/CakeExcel)
[![Build Status](https://travis-ci.org/dakota/CakeExcel.png?branch=3.0)](https://travis-ci.org/dakota/CakeExcel)

A plugin to generate Excel files with CakePHP.

Requirements
------------

* PHP 5.4.16
* CakePHP 3.0+
* [PHPExcel](https://github.com/PHPOffice/PHPExcel)
* Composer


Installation
------------

Run
```
composer require dakota/cake-excel 3.1.x-dev
```

Usage
-----

In app/Config/bootstrap.php add:
```
Plugin::load('CakeExcel', ['bootstrap' => true, 'routes' => true]);
```


Add the RequestHandler component to AppController if not loading the plugin's bootstrap, and map xlsx to the CakeExcel plugin
```
'RequestHandler' => array(
	'viewClassMap' => array(
		'xlsx' => 'CakeExcel.Excel'
	)
),
```

Create a link to the a action with the .xlsx extension
```
$this->Html->link('Excel file', array('_ext' => 'xlsx'));
```

Place the view templates in a 'xlsx' subdir, for instance `src/Template/Invoices/xlsx/index.ctp`, you also need a layout file, `src/Template/Layout/xlsx/default.ctp`

Inside your view file you will have access to the PHPExcel library with `$this->PhpExcel`. Please see the [PHPExcel](https://github.com/PHPOffice/PHPExcel) documentation for a guide on how to use PHPExcel.
