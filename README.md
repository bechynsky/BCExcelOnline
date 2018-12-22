# Excel Online integration

This article describe how to store data from BigClown do [Microsoft Excel Online](https://products.office.com/excel).

## Prerequsities

* MQTT broker listening to BigClown dongle.
* Microsoft Id
* Python 3

## Setup online environment

### Excel table

* Login to [Onedrive](https://onedrive.live.com) using your Microsoft Id.

* Open folder or create new one where you want to place new Excel file to store data from BigClown.

![Onderive empty folder](onedrive-01.png)
* Click new on left up corner and choose _New_ - _Excel Workbook_

![New Excel Workbook](new-excel.png)

* It opens new Excel file in online editor.

* Click name of Excel workbook on top of screen and rename it.

![Rename Excel](rename-excel.png)

* Create table header _device_, _sensor_, _sensorInfo_, _measurement_, _value_, _time_.

![Excel table header](excel-table-header.png)

* Create Table from header. Click to eny cell with table header and click _Format as Table_.

![Excel format as Table](excel-format-as-table.png)

* Check _My table has header_.

![Excel create table](excel-create-table.png)

* Excel is ready.

![New table](excel-new-table.png)

* Close browser tab with Excel.

### Flow

* On top left corner click _List of Microsoft Services_ and open _Flow_.

![Open Flow](flow-open.png)

* On top bar on left click _My Flows_ and click _Create from blank_. Click button _Create from blank_.

![Create new Flow from blank](flow-create-blank.png)

* Now you see empty Flow.

![Blank Flow](flow-blank-flow.png)

* Search for _Request_ action.

![Search request action](flow-search-request.png)

* Click _Request - When a HTTP request is received_.

![Request first stepan](flow-request-01.png)

* Now we need to describe JSON object we will send to Flow. Click _Use sample payload to generate schema_ bellow text box.

* Paste sample JSON to text box.

```
{"device": "temperature-button:0", "sensor": "thermometer", "sensorInfo": "0:0", "measurement": "temperature", "value": "27.44", "time": "2018-10-06 20:08:38"}
```

![JSON schema](flow-json-schema.png)

* Click _Done_.

![Fineshed JSON schema](flow-json-schema-finished.png)

* Click _+ New Step_ and _Add an Action_. Search _Excel_.

![Add Excel action](flow-excel-action.png)

* Click _Excel - Insert row_.

![Excel Insert row](flow-excel-insert-row.png)

* Choose Excel file you created.

![Choose Excel file](flow-choose-excel-file.png)

* When you create table it gets name. Choose name of created table.

![Choose Table name](flow-choose-table.png)

* Map columns to properties of JSON.

![Map JSON properties](flow-excel-map-properties.png)

* Click _Save_ on top right.
* Click _My flows_.

![Flow list](flow-save-close.png)

## Local application

* Clone BC2JsonPostPy

```
git clone https://github.com/bechynsky/BC2JsonPostPy.git
```

* Go back to _Flow_ dashboard.
* Click _Edit flow_ symbol ![Edit symbol](flow-edit-symbol.png)
* Click _Request_ action and copy _HTTP POST URL_

![Copy post URL](flow-copy-post-url.png)

* Edit _config.ini_. In _URL_ you must doubled _%_ symbol.

```
[DEFAULT]
URL = https://prod-26.westeurope.logic.azure.com:443/workflows/7bc74b9fce644126a478ae047d835a7e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%%2Ftriggers%%2Fmanual%%2Frun&sv=1.0&sig=ba_mcMZYbxxxxxxxxxD2TnfQie1KFM5o_7eQRKm0A
MQTT_SERVER = localhost
MQTT_PORT = 1883
```

* Run application. 

```
python3 main.py
```

* Sample output. Answer from server _202_ means everything is OK. If you get _404_ you have wrong _URL_.

```
user@user-pc:~/BC2JsonPostPy$ python3 main.py 
Connected with result code 0
{'device': 'temperature-button:0', 'sensor': 'thermometer', 'sensorInfo': '0:0', 'measurement': 'temperature', 'value': b'26.44', 'time': '2018-10-06 20:55:56'}
202
{'device': 'temperature-button:0', 'sensor': 'thermometer', 'sensorInfo': '0:0', 'measurement': 'temperature', 'value': b'26.44', 'time': '2018-10-06 20:55:59'}
202
{'device': 'temperature-button:0', 'sensor': 'thermometer', 'sensorInfo': '0:0', 'measurement': 'temperature', 'value': b'26.44', 'time': '2018-10-06 20:56:03'}
202
{'device': 'kit-lcd-thermostat:0', 'sensor': 'thermometer', 'sensorInfo': 'set-point', 'measurement': 'temperature', 'value': b'20.00', 'time': '2018-10-06 20:59:06'}
202

```

## Open Excel

* Open Excel to see data.

![Excel data](excel-with-data.png)
