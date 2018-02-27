# yii2 Exportable Gridview
Gridview with **full data export**

This works as normal gridview, but adds a button to export the data loaded into it.
It exports all records found on the active record query. Works with the same data
as the gridview does. Uses [phpspreadsheet](https://github.com/PHPOffice/PhpSpreadsheet)
to generate export file.

## How does it work
Following the same behavior that pjax uses, this widget act as a normal gridview until
a request is made with the url query parameters `export-grid` and `export-container`.
The latter should contain the id of gridview without the hashtag. When both parameters are received, response
is cleared and then spreadsheet generation begins. This allow to generate a file with absolutely all
records found on the gridview.


> This project is currently under development. Any contribution is welcome.


### Installation
```
composer require eseperio/yii2-exportable-gridview @dev
```

### Usage

This widget extends from yii2-gridview but add functionality to export
**all the rows** queried by the DataProvider.


```php
use eseperio\gridview\ExportableGridview as GridView;
<?= GridView::widget([
        'dataProvider' => $dataProvider,
        'filterModel' => $searchModel,
        'columns' => [
            ['class' => 'yii\grid\SerialColumn'],
            'id',
            'title',
            'description',
            'author',
        ],
    ]);
    ?>

```

### Additional configuration
asd

|Name|Type|default|Description|
|----|----|-------|-----------|
|`layout`|string|{summary} {items} {export} {pager}|In addition to default layout this gridview has `{export`} section. This is the place for export button.|
|`fileName`|string|exported.xls|Name to use on the generated filename. If `writerType` value is not set then the writer will be guessed from the extension.|
|`writerType`|string|null| The writer to be used when generating file. See [Spreadsheet writer](https://phpspreadsheet.readthedocs.io/en/develop/topics/reading-and-writing-to-file/). Accepts Xls, Xlsx, Ods, Csv, Html, Tcpdf, Dompdf, Mpdf|
|`exportable`|boolean|true|Whether to enable export for this gridview |
|`exportLinkOptions`|array| `['class'=> 'btn btn-default', 'target'=>'_blank']` |Options for the export link. It also accepts `label` and `encode`|
|`exportColumns`|array|empty|Property to define a different column combination for export only. If empty default columns of gridview will be used|

### Notes
All html tags are removed when exporting.

## Todo
* [ ] Add option to exclude certain columns like ActionColumn.

