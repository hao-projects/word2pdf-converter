# word2pdf-converter
word2pdf converter in java by soffice command on terminal
## requirements
1. linux environment
2. libreoffice is already installed
3. ports 8100, 8101, 8102, 8103, 8104, 8105, 8106, 8107, 8108, 8109 are available（of course,you can modify ports in the properties file,but you have to run `mvn package` at converter dir then run the jar file in converter/target/ diretory）
## usage
run `java -jar converter.xx.jar` & open localhost:8090 in your browser send your MS word file,then you can get the pdf file
