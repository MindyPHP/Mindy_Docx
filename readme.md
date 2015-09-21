Example usage:

Create docx file with variables like `{{ title }}`

```php
$phpDocx = new \Mindy\Docx\Docx();
$phpDocx->render("./template.docx", [
    'title' => 'My super title'
]);
$phpDocx->save("./new_document.docx");
```
