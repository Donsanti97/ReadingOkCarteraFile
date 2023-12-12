package org.dataTest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

public class MethotsAzureMasterFiles {

    public static void buscarYListarArchivos(String ubicacion) throws IOException {
        Path ruta = Paths.get(ubicacion);

        if (!Files.exists(ruta)) {
            System.out.println("La ubicación no existe. Creando...");
            Files.createDirectories(ruta);
            System.out.println("Ubicación creada: " + ubicacion);
        } else {
            System.out.println("La ubicación ya existe: " + ubicacion);
            listarArchivosEnCarpeta(ruta);
        }
    }

    public static void listarArchivosEnCarpeta(Path carpeta) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(carpeta)) {
            for (Path archivo : stream) {
                if (Files.isRegularFile(archivo)) {
                    System.out.println("Archivo: " + archivo.getFileName());
                }
            }
        }
    }


    public static String getDocument() {
        // Crea un objeto JFileChooser
        JFileChooser fileChooser = new JFileChooser();

        // Configura el directorio inicial en la carpeta de documentos del usuario
        String rutaDocumentos = System.getProperty("user.home") + File.separator + "Documentos";
        fileChooser.setCurrentDirectory(new File(rutaDocumentos));

        // Filtra para mostrar solo archivos de Excel
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Archivos Excel", "xlsx", "xls"));

        // Muestra el diálogo de selección de archivo
        int resultado = fileChooser.showOpenDialog(null);

        if (resultado == JFileChooser.APPROVE_OPTION) {
            File archivoSeleccionado = fileChooser.getSelectedFile();
            String rutaCompleta = archivoSeleccionado.getAbsolutePath();
            return rutaCompleta;
        } else {
            return null; // Si no se seleccionó ningún archivo, retorna null
        }
    }

    public static String getDirectory() {
        // Crea un objeto JFileChooser
        JFileChooser fileChooser = new JFileChooser();

        // Configura el directorio inicial en la carpeta de documentos del usuario
        String rutaDocumentos = System.getProperty("user.home")/* + File.separator + "Documentos"*/;
        fileChooser.setCurrentDirectory(new File(rutaDocumentos));

        // Filtra para mostrar solo archivos de Excel
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        // Muestra el diálogo de selección de archivo
        int resultado = fileChooser.showOpenDialog(null);

        if (resultado == JFileChooser.APPROVE_OPTION) {
            File archivoSeleccionado = fileChooser.getSelectedFile();
            String rutaCompleta = archivoSeleccionado.getAbsolutePath();
            return rutaCompleta;
        } else {
            return null; // Si no se seleccionó ningún archivo, retorna null
        }
    }

    /*-------------------------------------------------------------------------------------------------------------------------------*/
    public static int findSheetIndexInExcelB(String excelAFilePath, String excelBFilePath, String targetSheetName) throws IOException {
        FileInputStream excelAFile = new FileInputStream(excelAFilePath);
        FileInputStream excelBFile = new FileInputStream(excelBFilePath);

        Workbook workbookA = new XSSFWorkbook(excelAFile);
        Workbook workbookB = new XSSFWorkbook(excelBFile);

        int sheetIndexInB = -1;

        for (int i = 0; i < workbookB.getNumberOfSheets(); i++) {
            if (workbookB.getSheetName(i).equals(targetSheetName)) {
                sheetIndexInB = i;
                break;
            }
        }

        List<String> removedSheetNames = new ArrayList<>();

        if (sheetIndexInB != -1) {
            // Elimina las hojas anteriores a la hoja objetivo en Excel B
            for (int i = 0; i < sheetIndexInB; i++) {
                String sheetNameToRemove = workbookB.getSheetName(i);
                removedSheetNames.add(sheetNameToRemove);
            }
        }

        // Cerrar los archivos
        excelAFile.close();
        excelBFile.close();

        return sheetIndexInB;
    }

    public static void runtime() {
        Runtime runtime = Runtime.getRuntime();
        long minRunningMemory = (8L * 1024L * 1024L * 1024L);
        if (runtime.freeMemory() < minRunningMemory) {
            System.gc();
        }
    }
    /*---------------------------------------------------------------------------------------------------------------*/

    public static List<String> getWorkSheet(String filePath, int i) {
        List<String> shetNames = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            int numberOfSheets = workbook.getNumberOfSheets();
            ;
            for (int index = i; index < numberOfSheets; index++) {
                Sheet sheet = workbook.getSheetAt(index);
                shetNames.add(sheet.getSheetName());
            }
            workbook.close();
            fis.close();

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return shetNames;
    }

    public static List<Map<String, String>> getValuebyHeader(String excelFilePath, String sheetName) {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = getHeaders(excelFilePath, sheetName);
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String header = headers.get(cellIndex);
                    String value = "";
                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING) {
                            value = cell.getStringCellValue();
                            break;
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            value = String.valueOf(cell.getNumericCellValue());
                        }
                    }
                    rowData.put(header, value);
                }
                data.add(rowData);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }

    public static List<String> getHeaders(String excelFilePath, String sheetName) {
        List<String> headers = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return headers;
    }

    public static List<String> getHeaders(Sheet sheet) {
        List<String> encabezados = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Aquí puedes especificar en qué fila esperas encontrar los encabezados
            // Por ejemplo, si están en la tercera fila (fila índice 2), puedes usar:
            if (row.getRowNum() == 0) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    encabezados.add(obtenerValorVisibleCelda(cell));
                }
                break; // Terminamos de buscar encabezados una vez que los encontramos
            }
        }

        return encabezados;
    }

    public static List<String> findValueInColumn(Sheet sheet, int columnaBuscada, String valorBuscado) {
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnaBuscada);
            String valorCelda = obtenerValorVisibleCelda(cell);

            if (valorBuscado.equals(valorCelda)) {
                return obtenerValoresFila(row);
            }
        }

        return null; // Valor no encontrado en la columna especificada
    }

    public static List<String> getHeadersMasterfile(Sheet sheet1, Sheet sheet2) throws IOException {
        List<String> headers1 = getHeaders(sheet1);
        String headerFirstFile1 = headers1.get(0);
        List<String> headers2 = getHeaders(sheet2);
        String headerSecondFile = headers2.get(0);

        if (!headerFirstFile1.equals(headerSecondFile)) {
            headers2 = findValueInColumn(sheet1, 0, headers1.get(0));
        }

        return headers2;
    }

    public static List<String> obtenerValoresFila(Row row) {
        List<String> valoresFila = new ArrayList<>();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            valoresFila.add(obtenerValorVisibleCelda(cell));//obtenerValorCelda()
        }
        return valoresFila;
    }

    public static String obtenerValorCelda(Cell cell) {
        String valor = "";
        if (cell != null) {
            try {
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getStringCellValue() + ", CELL: " + cell);
                        valor = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            //valor = cell.getDateCellValue().toString();
                            System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getNumericCellValue() + ", CELL: " + cell);
                            String formatDate = cell.getDateCellValue().toString();
                            SimpleDateFormat formatoEntrada = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                            Date date = formatoEntrada.parse(formatDate);
                            SimpleDateFormat formatoSalida = new SimpleDateFormat("dd/MM/yyyy");
                            //valor = formatoSalida.format(date);
                            valor = cell.toString();
                            System.out.println("VALOR1 " + valor);
                        }else {
                            //valor = Date.toString(cell.getDateCellValue()/*NumericCellValue()*/);
                            System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getNumericCellValue() + ", CELL: " + cell);
                            valor = cell.getStringCellValue();
                            System.out.println("VALOR3 " + valor);
                        }
                        break;
                    case BOOLEAN:
                        System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getBooleanCellValue() + ", CELL: " + cell);
                        valor = Boolean.toString(cell.getBooleanCellValue());
                        break;
                    case FORMULA:
                            System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getCellFormula().toString() + ", CELL: " + cell + " CELLADD: " + cell.getAddress());
                        /*System.err.println("Formato fecha no valido."+ cell.getCellFormula() +" Encabezado "+ cell.getSheet().getSheetName() +" Posición: "+ cell.getAddress() +" puede contener formula o valor cadena de caracteres");
                        valor = evaluarFormula(cell);
                        System.out.println("VALORF: " + valor);
                        FunctionsApachePoi.waitSeconds(20);
                        System.exit(1);*/
                        valor = obtenerValorCeldaString(cell);

                        //break;
                    default:
                        /*valor = obtenerValorCeldaString(cell);*/
                        break;
                }

            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        return valor;
    }

    public static String obtenerValorVisibleCelda(Cell cell) {
        try {
            DataFormatter dataFormatter = new DataFormatter();
            String valor = "";

            // Verificar el tipo de celda
            switch (cell.getCellType()) {
                case STRING:
                    valor = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        valor = dataFormatter.formatCellValue(cell);
                    } else {
                        double numericValue = cell.getNumericCellValue();
                        String dataFormatString = cell.getCellStyle().getDataFormatString();

                        if (numericValue >= -99.99 && numericValue <= 99.99) {
                            if (numericValue == 0){
                                valor = dataFormatter.formatRawCellContents(cell.getNumericCellValue(), cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString());

                            }else {
                                boolean isTwoDigitsOrLess = Math.abs(numericValue) < 100 && Math.abs(numericValue) % 1 !=0;
                                if (isTwoDigitsOrLess) {
                                    valor = String.format("%.2f%%", numericValue/* / 100*/);
                                }else {
                                    valor = String.valueOf(numericValue);
                                }
                            }
                        }else {
                            valor = dataFormatter.formatRawCellContents(cell.getNumericCellValue(), cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString());
                        }
                    }
                    break;
                case BOOLEAN:
                    valor = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case BLANK:
                case _NONE:
                case ERROR:
                default:
                    valor = /*dataFormatter.formatCellValue(cell)*/"0";
            }

            return valor;
        } catch (Exception e) {
            return "";
        }
    }

    public static String evaluarFormulas(Cell cell) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);
            if (cellValue.getCellType() == CellType.NUMERIC) {
                double valor = cellValue.getNumberValue();
                //System.out.println("El valor de la fórmula en A5 es: " + valor);
                return Double.toString(valor);
            } else if (cellValue.getCellType() == CellType.STRING) {
                String valor = cellValue.getStringValue();
                //System.out.println("El valor de la fórmula en A5 es: " + valor);
                return valor;
            }else {
                return cellValue.formatAsString();
            }
        } catch (Exception e) {
            return "";
        }
    }

    public static String obtenerValorCeldaString(Cell cell) {
        try {
            DataFormatter dataFormatter = new DataFormatter();
            String valor = dataFormatter.formatCellValue(cell);
            return valor;
        } catch (Exception e) {
            return "";
        }
    }

    public static String evaluarFormula(Cell cell) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);

            if (cellValue.getCellType() == CellType.FORMULA) {
                // Si la celda contiene una fórmula, obtén su valor calculado
                if (cellValue.getCellType() == CellType.NUMERIC) {
                    double valor = cellValue.getNumberValue();
                    return Double.toString(valor);
                } else if (cellValue.getCellType() == CellType.STRING) {
                    return cellValue.getStringValue();
                }
            } else {
                // Si no es una fórmula, obtén el valor directo de la celda
                System.out.println("NO ES FORMULA");
                DataFormatter dataFormatter = new DataFormatter();
                String valor = dataFormatter.formatCellValue(cell);
                return valor;
            }
        } catch (Exception e) {
            return "";
        }

        return ""; // Valor por defecto si no se pudo obtener el valor de la fórmula
    }

    public static List<Map<String, String>> createMapList(List<Map<String, String>> originalList, String keyHeader, String valueHeader) {
        List<Map<String, String>> mapList = new ArrayList<>();

        for (Map<String, String> originalMap : originalList) {
            String key = originalMap.get(keyHeader);
            String value = originalMap.get(valueHeader);

            Map<String, String> newMap = new HashMap<>();
            newMap.put(key, value);

            mapList.add(newMap);
        }

        return mapList;
    }

    public static List<Map<String, String>> obtenerValoresPorFilas(Workbook workbook1, Workbook workbook2, String sheetName1, String sheetName2, String header1, String header2) throws IOException {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();
        Sheet sheet1 = workbook1.getSheet(sheetName1);
        Sheet sheet2 = workbook2.getSheet(sheetName2);

        List<String> encabezados = getHeadersMasterfile(sheet1, sheet2);

        int indexHeader1 = encabezados.indexOf(header1);
        int indexHeader2 = encabezados.indexOf(header2);

        if (indexHeader1 == -1 || indexHeader2 == -1) {
            // Los encabezados especificados no se encontraron en la lista de encabezados
            // Puedes manejar esta situación como desees, por ejemplo, lanzando una excepción
            throw new IllegalArgumentException("Los encabezados especificados no se encontraron en la hoja.");
        }

        Iterator<Row> rowIterator = sheet1.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            Map<String, String> fila = new HashMap<>();

            try {
                if (indexHeader1 >= 0 && indexHeader1 < valoresFila.size() &&
                        indexHeader2 >= 0 && indexHeader2 < valoresFila.size()) {
                    fila.put(header1, valoresFila.get(indexHeader1));
                    fila.put(header2, valoresFila.get(indexHeader2));
                } else {
                    System.err.println("En la fila [" + row.getRowNum() + "] no se encuentran los datos completos. " +
                            "\n Por favor rellene con [0] o con [NA] según el campo que falte numérico o caracteres respectivamente");
                }
            } catch (Exception e) {
                throw new RuntimeException(e);
            }


            valoresPorFilas.add(fila);
        }

        return valoresPorFilas;
    }

    public static List<Map<String, String>> obtenerValoresPorFilas(Workbook workbook1, Workbook workbook2, String sheetName1, String sheetName2) throws IOException {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();
        Sheet sheet1 = workbook1.getSheet(sheetName1);
        Sheet sheet2 = workbook2.getSheet(sheetName2);

        List<String> encabezados = getHeadersMasterfile(sheet1, sheet2);

        Iterator<Row> rowIterator = sheet1.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            Map<String, String> fila = new HashMap<>();
            for (int i = 0; i < encabezados.size() && i < valoresFila.size(); i++) {
                String encabezado = encabezados.get(i);
                String valor = valoresFila.get(i);
                fila.put(encabezado, valor);
            }

            valoresPorFilas.add(fila);
        }

        return valoresPorFilas;
    }
    public static List<Map<String, String>> obtenerValoresPorFilas(Sheet sheet, Sheet sheet2) throws IOException {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();
        List<String> encabezados = getHeadersMasterfile(sheet, sheet2);

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            Map<String, String> fila = new HashMap<>();
            for (int i = 0; i < encabezados.size() && i < valoresFila.size(); i++) {
                String encabezado = encabezados.get(i);
                String valor = valoresFila.get(i);
                fila.put(encabezado, valor);
            }

            valoresPorFilas.add(fila);
        }

        return valoresPorFilas;
    }

    public static List<Map<String, String>> obtenerValoresPorFilas(Sheet sheet, List<String> encabezados) {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            Map<String, String> fila = new HashMap<>();
            for (int i = 0; i < encabezados.size() && i < valoresFila.size(); i++) {
                String encabezado = encabezados.get(i);
                String valor = valoresFila.get(i);
                fila.put(encabezado, valor);
            }

            valoresPorFilas.add(fila);
        }

        return valoresPorFilas;
    }

    public static Map<String, String> obtenerValoresPorEncabezado(Sheet sheet, String encabezadoCodCiudad, String encabezadoFecha) {
        Map<String, String> valoresPorCodCiudad = new HashMap<>();

        List<String> encabezados = obtenerValoresFila(sheet.getRow(0)); // Obtener encabezados de la primera fila
        int columnaCodCiudad = -1;
        int columnaFecha = -1;

        // Encontrar las columnas de los encabezados específicos
        for (int i = 0; i < encabezados.size(); i++) {
            String encabezado = encabezados.get(i);
            if (encabezado.equals(encabezadoCodCiudad)) {
                columnaCodCiudad = i;
            }
            if (encabezado.equals(encabezadoFecha)) {
                columnaFecha = i;
            }
        }

        if (columnaCodCiudad == -1 || columnaFecha == -1) {
            return valoresPorCodCiudad; // No se encontraron los encabezados especificados
        }

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String codCiudad = obtenerValorVisibleCelda(row.getCell(columnaCodCiudad));
            String valorFecha = obtenerValorVisibleCelda(row.getCell(columnaFecha));
            valoresPorCodCiudad.put(codCiudad, valorFecha);
        }

        return valoresPorCodCiudad;
    }


}
