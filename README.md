[![Build Status](https://img.shields.io/travis/dakota/CakeExcel/master.svg?style=flat-square)](https://travis-ci.org/dakota/CakeExcel)
[![Coverage Status](https://img.shields.io/coveralls/dakota/CakeExcel.svg?style=flat-square)](https://coveralls.io/r/dakota/CakeExcel?branch=master)
[![Total Downloads](https://img.shields.io/packagist/dt/dakota/cake-excel.svg?style=flat-square)](https://packagist.org/packages/dakota/cake-excel)
[![Latest Stable Version](https://img.shields.io/packagist/v/dakota/cake-excel.svg?style=flat-square)](https://packagist.org/packages/dakota/cake-excel)

# CakeExcel

A plugin to generate Excel files with CakePHP.

## Requirements

* CakePHP 3.x
* PHP 5.4.16 or greater
* Patience

## Installation

_[Using [Composer](http://getcomposer.org/)]_

```
composer require dakota/cake-excel
```

### Enable plugin

Load the plugin in your app's `config/bootstrap.php` file:

    Plugin::load('CakeExcel', ['bootstrap' => true, 'routes' => true]);

## Installation

Run
```
composer require dakota/cake-excel 3.1.x-dev
```

## Usage

First, you'll want to setup extension parsing for the `xlsx` extension. To do so, you will need to add the following to your `config/routes.php` file:

```php
Router::extensions('xlsx');
```

Next, we'll need to add a viewClassMap entry to your Controller. You can place the following in your AppController:

```php
public $components = [
    'RequestHandler' => [
        'viewClassMap' => [
            'xlsx' => 'CakeExcel.Excel',
        ],
    ]
];
```

Each application *must* have an xlsx layout. The following is a barebones layout that can be placed in `src/Template/Layout/xlsx/default.ctp`:

```php
<?= $this->fetch('content') ?>
```

Finally, you can link to the current page with the .xlsx extension. This assumes you've created an `xlsx/index.ctp` file in your particular controller's template directory:

```php
$this->Html->link('Excel file', ['_ext' => 'xlsx']);
```

Inside your view file you will have access to the PHPExcel library with `$this->PhpExcel`. Please see the [PHPExcel](https://github.com/PHPOffice/PHPExcel) documentation for a guide on how to use PHPExcel.

## License

The MIT License (MIT)

Copyright (c) 2013 Walther Lalk

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
