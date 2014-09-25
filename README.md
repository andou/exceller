# PHP Exceller

PHPExcel Wrapper to easily print excel files

## Installation

Add a dependency on `andou/exceller` to your project's `composer.json` file if you use [Composer](http://getcomposer.org/) to manage the dependencies of your project.
You have to also add the relative repository.

Here is a minimal example of a `composer.json` file that just defines a dependency on `andou/exceller`:

```json
{
    "require": {
        "andou/exceller": "*"
    },
    "repositories": [
    {
      "type": "git",
      "url": "https://github.com/andou/exceller.git"
    }
  ],
}
```    

## Usage Examples

```php
require_once './vendor/autoload.php';

$exceller = new Andou\Exceller\Exceller();

$exceller
        ->setSavePath(__DIR__ . "/output")
        ->setFileName("test")
        ->insertHeaderCell("A", 1, "header")
        ->insertCell("A", 3, 'value')
        ->finalize();

```

