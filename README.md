# Batch-test
https://github.com/monitorjbl/excel-streaming-reader

Excel-Streaming-Reader é uma biblioteca que utiliza de todas as funcionalidades do Apache POI com XSSFReader and outras funcionalidades de da Event API e SAX para tornar o processo todo muito mais encapsulado e no final retorna uma Workbook que então possui todas outras funcionalidades para manipular o arquivo.

Essa biblioteca permite abrir arquivos em excel muito grande fazendo streaming e não carregando tudo em memoria (que causa a exception que execede a quantidade de memoria).
